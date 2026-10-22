"""
word_splitter_core.py
=====================
Core logic for analyzing and splitting Microsoft Word (.docx) documents.
Supports splitting by:
  - 'memory' (Equal Memory / File Size on disk - Recommended)
  - 'words' (Balanced Word Count)
  - 'blocks' (Equal Content Blocks)
  - 'headings' (By Chapters / Headings)
Preserves formatting, styles, headers, footers, tables, and embedded media
by cloning the original document package and trimming XML body elements.
"""

import os
import shutil
import zipfile
import re
import xml.etree.ElementTree as ET
from copy import deepcopy
from typing import List, Dict, Any, Optional, Callable
import docx
from docx.oxml import OxmlElement

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


class WordSplitter:
    """Handles parsing, partition calculation, and splitting of .docx documents."""

    @staticmethod
    def is_docx(file_path: str) -> bool:
        """Check if file exists and has .docx extension."""
        return os.path.isfile(file_path) and file_path.lower().endswith(".docx")

    @classmethod
    def _expand_drawings(cls, body):
        """
        If a paragraph contains multiple embedded drawings/screenshots,
        unpacks each drawing (and its trailing whitespace/breaks) into its own paragraph.
        This allows memory-based and element-based splitting to distribute individual
        images evenly across parts rather than trapping multiple screens in one giant block.
        """
        for p in list(body):
            if not p.tag.endswith("p"):
                continue
            drawings = p.xpath(".//w:drawing")
            if len(drawings) <= 1:
                continue

            pPr = p.find(f"{{{W_NS}}}pPr")
            pPr_copy = deepcopy(pPr) if pPr is not None else None

            children = list(p)
            chunks = []
            curr_chunk = []
            for c in children:
                if c.tag.endswith("pPr"):
                    continue
                curr_chunk.append(c)
                if len(c.xpath(".//w:drawing")) > 0:
                    chunks.append(curr_chunk)
                    curr_chunk = []
            if curr_chunk:
                if chunks:
                    chunks[-1].extend(curr_chunk)
                else:
                    chunks.append(curr_chunk)

            p_index = body.index(p)
            for i, chunk in enumerate(chunks):
                new_p = OxmlElement("w:p")
                if pPr_copy is not None:
                    new_p.append(deepcopy(pPr_copy))
                for elem in chunk:
                    new_p.append(elem)
                body.insert(p_index + i, new_p)

            body.remove(p)

    @classmethod
    def _get_media_mapping(cls, file_path: str):
        """Read media sizes and relationship mapping from docx zip package."""
        media_sizes = {}
        rid_to_target = {}

        try:
            with zipfile.ZipFile(file_path, "r") as z:
                for info in z.infolist():
                    if info.filename.startswith("word/media/"):
                        # Use compressed size to reflect actual storage memory
                        media_sizes[info.filename] = info.compress_size

                if "word/_rels/document.xml.rels" in z.namelist():
                    rels_xml = z.read("word/_rels/document.xml.rels")
                    rels_root = ET.fromstring(rels_xml)
                    for rel in rels_root:
                        rid = rel.attrib.get("Id")
                        tgt = rel.attrib.get("Target", "")
                        norm_tgt = os.path.normpath(os.path.join("word", tgt)).replace("\\", "/")
                        if norm_tgt.startswith("/"):
                            norm_tgt = norm_tgt[1:]
                        rid_to_target[rid] = norm_tgt
        except Exception:
            pass

        return media_sizes, rid_to_target

    @classmethod
    def analyze_document(cls, file_path: str) -> Dict[str, Any]:
        """
        Analyze a .docx document and return detailed statistics:
        - Total body blocks (paragraphs and tables)
        - Paragraphs count, table count, drawings count
        - Total word count & character count
        - Total memory footprint (bytes of text + media)
        - Per-block weights for memory balancing
        - Headings list (level, text, block index)
        - Section count
        """
        if not os.path.isfile(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")

        media_sizes, rid_to_target = cls._get_media_mapping(file_path)

        doc = docx.Document(file_path)
        body = doc._body._body
        cls._expand_drawings(body)

        blocks = []
        headings = []
        page_breaks = []
        total_words = 0
        total_chars = 0
        total_drawings = 0
        total_memory_bytes = 0
        para_count = 0
        tbl_count = 0

        for idx, child in enumerate(body):
            tag_name = child.tag.split("}")[-1] if "}" in child.tag else child.tag

            if tag_name == "p":
                para_count += 1
                texts = child.xpath(".//w:t")
                text = "".join(t.text or "" for t in texts)
                words = len(text.split())
                chars = len(text)
                total_words += words
                total_chars += chars

                # Calculate memory footprint of this block
                xml_str = ET.tostring(child, encoding="utf-8").decode("utf-8")
                rids = set(re.findall(r'(?:embed|id|link)="([^"]+)"', xml_str))
                media_bytes = sum(media_sizes.get(rid_to_target.get(r, ""), 0) for r in rids)
                text_bytes = len(xml_str.encode("utf-8"))
                block_weight = media_bytes + text_bytes
                total_memory_bytes += block_weight

                drawings_count = len(child.xpath(".//w:drawing"))
                total_drawings += drawings_count

                # Check if paragraph has heading style
                style_el = child.xpath(".//w:pPr/w:pStyle")
                style_val = style_el[0].get(f"{{{W_NS}}}val") if style_el else ""
                is_heading = False
                heading_level = 0
                if style_val and ("heading" in style_val.lower() or "heading" in (child.attrib.get("style", "").lower())):
                    is_heading = True
                    match = re.search(r"\d+", style_val)
                    heading_level = int(match.group()) if match else 1

                has_page_break = bool(child.xpath(".//w:br[@w:type='page']") or child.xpath(".//w:lastRenderedPageBreak"))
                if has_page_break:
                    page_breaks.append(idx)

                block_info = {
                    "type": "p",
                    "index": len(blocks),
                    "xml_element": child,
                    "words": words,
                    "chars": chars,
                    "drawings": drawings_count,
                    "weight_bytes": block_weight,
                    "text_preview": text[:60].strip() if text else (f"[Image / Drawing]" if drawings_count else ""),
                    "is_heading": is_heading,
                    "heading_level": heading_level,
                    "has_page_break": has_page_break,
                }
                blocks.append(block_info)

                if is_heading:
                    headings.append({
                        "block_index": len(blocks) - 1,
                        "level": heading_level,
                        "text": text.strip()[:80]
                    })

            elif tag_name == "tbl":
                tbl_count += 1
                texts = child.xpath(".//w:t")
                text = "".join(t.text or "" for t in texts)
                words = len(text.split())
                chars = len(text)
                total_words += words
                total_chars += chars

                xml_str = ET.tostring(child, encoding="utf-8").decode("utf-8")
                rids = set(re.findall(r'(?:embed|id|link)="([^"]+)"', xml_str))
                media_bytes = sum(media_sizes.get(rid_to_target.get(r, ""), 0) for r in rids)
                text_bytes = len(xml_str.encode("utf-8"))
                block_weight = media_bytes + text_bytes
                total_memory_bytes += block_weight

                drawings_count = len(child.xpath(".//w:drawing"))
                total_drawings += drawings_count

                block_info = {
                    "type": "tbl",
                    "index": len(blocks),
                    "xml_element": child,
                    "words": words,
                    "chars": chars,
                    "drawings": drawings_count,
                    "weight_bytes": block_weight,
                    "text_preview": f"[Table: {words} words]",
                    "is_heading": False,
                    "heading_level": 0,
                    "has_page_break": False,
                }
                blocks.append(block_info)

            elif tag_name == "sdt":
                texts = child.xpath(".//w:t")
                text = "".join(t.text or "" for t in texts)
                words = len(text.split())
                chars = len(text)
                total_words += words
                total_chars += chars

                xml_str = ET.tostring(child, encoding="utf-8").decode("utf-8")
                rids = set(re.findall(r'(?:embed|id|link)="([^"]+)"', xml_str))
                media_bytes = sum(media_sizes.get(rid_to_target.get(r, ""), 0) for r in rids)
                text_bytes = len(xml_str.encode("utf-8"))
                block_weight = media_bytes + text_bytes
                total_memory_bytes += block_weight

                drawings_count = len(child.xpath(".//w:drawing"))
                total_drawings += drawings_count

                block_info = {
                    "type": "sdt",
                    "index": len(blocks),
                    "xml_element": child,
                    "words": words,
                    "chars": chars,
                    "drawings": drawings_count,
                    "weight_bytes": block_weight,
                    "text_preview": f"[Control: {text[:40].strip()}]",
                    "is_heading": False,
                    "heading_level": 0,
                    "has_page_break": False,
                }
                blocks.append(block_info)

        file_size = os.path.getsize(file_path)

        return {
            "file_path": file_path,
            "file_name": os.path.basename(file_path),
            "file_size": file_size,
            "total_blocks": len(blocks),
            "total_paragraphs": para_count,
            "total_tables": tbl_count,
            "total_drawings": total_drawings,
            "total_words": total_words,
            "total_chars": total_chars,
            "total_memory_bytes": total_memory_bytes,
            "sections_count": len(doc.sections),
            "headings": headings,
            "page_breaks": page_breaks,
            "blocks": blocks,
        }

    @classmethod
    def calculate_partitions(
        cls,
        analysis: Dict[str, Any],
        num_parts: int,
        method: str = "memory"
    ) -> List[Dict[str, Any]]:
        """
        Calculate the slice ranges [start_block, end_block) for each part.

        Methods:
        - 'memory': Splits into parts with equal memory/file size (Default & Recommended)
        - 'words': Splits by balanced word count
        - 'blocks': Splits into equal block count (paragraphs + tables)
        - 'headings': Splits before top-level headings
        """
        blocks = analysis["blocks"]
        total_blocks = len(blocks)
        total_words = analysis["total_words"]
        total_memory = analysis.get("total_memory_bytes", 0)

        if total_blocks == 0:
            raise ValueError("The document contains no readable text, paragraphs, or tables.")

        if num_parts <= 0:
            raise ValueError("Number of parts must be at least 1.")

        actual_parts = min(num_parts, total_blocks)
        partitions = []

        if method == "blocks":
            base_size = total_blocks // actual_parts
            remainder = total_blocks % actual_parts
            start = 0
            for i in range(actual_parts):
                count = base_size + (1 if i < remainder else 0)
                end = start + count
                part_blocks = blocks[start:end]
                partitions.append({
                    "part_number": i + 1,
                    "start_idx": start,
                    "end_idx": end,
                    "block_count": count,
                    "words": sum(b["words"] for b in part_blocks),
                    "chars": sum(b["chars"] for b in part_blocks),
                    "weight_bytes": sum(b["weight_bytes"] for b in part_blocks),
                })
                start = end

        elif method == "headings" and len(analysis["headings"]) > 1:
            h_indices = [h["block_index"] for h in analysis["headings"]]
            if len(h_indices) <= actual_parts:
                split_points = sorted(list(set([0] + h_indices + [total_blocks])))
            else:
                step = len(h_indices) / actual_parts
                selected = [h_indices[int(round(i * step))] for i in range(actual_parts)]
                split_points = sorted(list(set([0] + selected + [total_blocks])))

            for i in range(len(split_points) - 1):
                start = split_points[i]
                end = split_points[i + 1]
                if start >= end:
                    continue
                part_blocks = blocks[start:end]
                partitions.append({
                    "part_number": len(partitions) + 1,
                    "start_idx": start,
                    "end_idx": end,
                    "block_count": end - start,
                    "words": sum(b["words"] for b in part_blocks),
                    "chars": sum(b["chars"] for b in part_blocks),
                    "weight_bytes": sum(b["weight_bytes"] for b in part_blocks),
                })

        elif method == "words" and total_words > 0:
            prefix_words = [0]
            for b in blocks:
                prefix_words.append(prefix_words[-1] + b["words"])

            target_per_part = total_words / actual_parts
            cut_points = [0]
            curr_cut = 0
            for p in range(1, actual_parts):
                ideal_cum = p * target_per_part
                max_allowed = total_blocks - (actual_parts - p)
                min_allowed = curr_cut + 1
                best_idx = min_allowed
                best_diff = abs(prefix_words[min_allowed] - ideal_cum)
                for candidate in range(min_allowed, max_allowed + 1):
                    diff = abs(prefix_words[candidate] - ideal_cum)
                    if diff < best_diff:
                        best_diff = diff
                        best_idx = candidate
                    elif diff > best_diff and prefix_words[candidate] > ideal_cum:
                        break
                cut_points.append(best_idx)
                curr_cut = best_idx
            cut_points.append(total_blocks)

            for i in range(len(cut_points) - 1):
                start = cut_points[i]
                end = cut_points[i + 1]
                part_blocks = blocks[start:end]
                partitions.append({
                    "part_number": i + 1,
                    "start_idx": start,
                    "end_idx": end,
                    "block_count": end - start,
                    "words": sum(b["words"] for b in part_blocks),
                    "chars": sum(b["chars"] for b in part_blocks),
                    "weight_bytes": sum(b["weight_bytes"] for b in part_blocks),
                })

        else:
            # Method == 'memory' (Equal Memory / File Size on Disk)
            # Prefix sums of memory byte weights
            prefix_bytes = [0]
            for b in blocks:
                prefix_bytes.append(prefix_bytes[-1] + b["weight_bytes"])

            target_per_part = total_memory / actual_parts

            # Optimal partitioning: for small N (<= 5), dynamic search finds exact minimum variance
            if actual_parts <= 4 and total_blocks <= 1000:
                cut_points = cls._find_optimal_cuts_variance(prefix_bytes, total_blocks, actual_parts, target_per_part)
            else:
                cut_points = [0]
                curr_cut = 0
                for p in range(1, actual_parts):
                    ideal_cum = p * target_per_part
                    max_allowed = total_blocks - (actual_parts - p)
                    min_allowed = curr_cut + 1
                    best_idx = min_allowed
                    best_diff = abs(prefix_bytes[min_allowed] - ideal_cum)
                    for candidate in range(min_allowed, max_allowed + 1):
                        diff = abs(prefix_bytes[candidate] - ideal_cum)
                        if diff < best_diff:
                            best_diff = diff
                            best_idx = candidate
                        elif diff > best_diff and prefix_bytes[candidate] > ideal_cum:
                            break
                    cut_points.append(best_idx)
                    curr_cut = best_idx
                cut_points.append(total_blocks)

            for i in range(len(cut_points) - 1):
                start = cut_points[i]
                end = cut_points[i + 1]
                part_blocks = blocks[start:end]
                partitions.append({
                    "part_number": i + 1,
                    "start_idx": start,
                    "end_idx": end,
                    "block_count": end - start,
                    "words": sum(b["words"] for b in part_blocks),
                    "chars": sum(b["chars"] for b in part_blocks),
                    "weight_bytes": sum(b["weight_bytes"] for b in part_blocks),
                })

        return partitions

    @classmethod
    def _find_optimal_cuts_variance(cls, prefix: List[int], N: int, K: int, target: float) -> List[int]:
        """Exhaustively find cut points for small K (2, 3, 4) that minimize variance from target."""
        if K == 1:
            return [0, N]
        if K == 2:
            best_c = 1
            best_diff = float("inf")
            for c in range(1, N):
                diff = (prefix[c] - target) ** 2 + ((prefix[N] - prefix[c]) - target) ** 2
                if diff < best_diff:
                    best_diff = diff
                    best_c = c
            return [0, best_c, N]
        if K == 3:
            best_cuts = (1, 2)
            best_diff = float("inf")
            for c1 in range(1, N - 1):
                w1 = prefix[c1]
                for c2 in range(c1 + 1, N):
                    w2 = prefix[c2] - prefix[c1]
                    w3 = prefix[N] - prefix[c2]
                    diff = (w1 - target)**2 + (w2 - target)**2 + (w3 - target)**2
                    if diff < best_diff:
                        best_diff = diff
                        best_cuts = (c1, c2)
            return [0, best_cuts[0], best_cuts[1], N]
        if K == 4:
            best_cuts = (1, 2, 3)
            best_diff = float("inf")
            for c1 in range(1, N - 2):
                w1 = prefix[c1]
                for c2 in range(c1 + 1, N - 1):
                    w2 = prefix[c2] - prefix[c1]
                    for c3 in range(c2 + 1, N):
                        w3 = prefix[c3] - prefix[c2]
                        w4 = prefix[N] - prefix[c3]
                        diff = (w1 - target)**2 + (w2 - target)**2 + (w3 - target)**2 + (w4 - target)**2
                        if diff < best_diff:
                            best_diff = diff
                            best_cuts = (c1, c2, c3)
            return [0, best_cuts[0], best_cuts[1], best_cuts[2], N]
        return [0, N]

    @classmethod
    def split_document(
        cls,
        file_path: str,
        num_parts: int,
        output_dir: Optional[str] = None,
        method: str = "memory",
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
    ) -> Dict[str, Any]:
        """
        Splits a .docx file into `num_parts` separate .docx files.
        Preserves all styles, formatting, tables, images, and section setups.
        """
        analysis = cls.analyze_document(file_path)
        partitions = cls.calculate_partitions(analysis, num_parts, method=method)

        base_name, _ = os.path.splitext(os.path.basename(file_path))

        if not output_dir:
            file_dir = os.path.dirname(os.path.abspath(file_path))
            output_dir = os.path.join(file_dir, f"{base_name}_split")

        os.makedirs(output_dir, exist_ok=True)

        generated_files = []
        total_parts = len(partitions)
        padding = len(str(total_parts))

        for part_info in partitions:
            p_num = part_info["part_number"]
            str_p_num = str(p_num).zfill(padding)
            out_filename = f"{base_name}_part_{str_p_num}.docx"
            out_path = os.path.join(output_dir, out_filename)

            # 1. Clone original document package to output destination
            shutil.copyfile(file_path, out_path)

            # 2. Open cloned document and expand multi-drawing paragraphs deterministically
            doc = docx.Document(out_path)
            body = doc._body._body
            cls._expand_drawings(body)

            # Find all relevant block children in body
            block_elements = [
                c for c in body
                if c.tag.endswith("p") or c.tag.endswith("tbl") or c.tag.endswith("sdt")
            ]

            start_idx = part_info["start_idx"]
            end_idx = part_info["end_idx"]

            # Set of elements to keep
            keep_elements = set(block_elements[start_idx:end_idx])

            for el in block_elements:
                if el not in keep_elements:
                    body.remove(el)

            # 3. Handle trailing section break
            kept_paras = [c for c in body if c.tag.endswith("p")]
            if kept_paras:
                last_p = kept_paras[-1]
                pPr = last_p.find(f"{{{W_NS}}}pPr")
                if pPr is not None:
                    sectPr = pPr.find(f"{{{W_NS}}}sectPr")
                    if sectPr is not None:
                        body_sectPr = body.find(f"{{{W_NS}}}sectPr")
                        if body_sectPr is not None:
                            body.replace(body_sectPr, sectPr)
                        else:
                            body.append(sectPr)

            # 4. Save trimmed part
            doc.save(out_path)

            # 5. Prune unreferenced media files to optimize file size
            cls._prune_unreferenced_media(out_path)

            file_size = os.path.getsize(out_path)
            file_meta = {
                "part_number": p_num,
                "file_name": out_filename,
                "file_path": out_path,
                "file_size": file_size,
                "start_idx": start_idx,
                "end_idx": end_idx,
                "block_count": part_info["block_count"],
                "words": part_info["words"],
                "chars": part_info["chars"],
                "weight_bytes": part_info.get("weight_bytes", file_size),
            }
            generated_files.append(file_meta)

            if progress_callback:
                progress_callback(p_num, total_parts, out_path)

        return {
            "source_file": file_path,
            "output_directory": output_dir,
            "method": method,
            "requested_parts": num_parts,
            "actual_parts": total_parts,
            "total_source_words": analysis["total_words"],
            "total_source_blocks": analysis["total_blocks"],
            "parts": generated_files,
        }

    @classmethod
    def _prune_unreferenced_media(cls, docx_path: str):
        """
        Prunes unused images and media from the docx package to keep split file sizes minimal.
        Preserves images used in the document body as well as headers and footers.
        """
        temp_path = docx_path + ".pruned.tmp"
        try:
            with zipfile.ZipFile(docx_path, "r") as zin:
                namelist = zin.namelist()
                has_media = any(name.startswith("word/media/") for name in namelist)
                if not has_media or "word/document.xml" not in namelist:
                    return

                xml_files_to_check = [
                    name for name in namelist
                    if name == "word/document.xml" or name.startswith("word/header") or name.startswith("word/footer")
                ]

                used_rids = set()
                for xml_name in xml_files_to_check:
                    try:
                        content = zin.read(xml_name).decode("utf-8")
                        matches = re.findall(r'(?:embed|id|link)="([^"]+)"', content)
                        used_rids.update(matches)
                    except Exception:
                        pass

                keep_targets = set()
                rels_to_update = {}

                for rels_name in [name for name in namelist if name.endswith(".rels")]:
                    try:
                        rels_content = zin.read(rels_name)
                        root = ET.fromstring(rels_content)
                        modified = False

                        for rel in list(root):
                            rid = rel.attrib.get("Id")
                            target = rel.attrib.get("Target", "")
                            rel_type = rel.attrib.get("Type", "")

                            is_media = "image" in rel_type.lower() or "media" in target.lower()
                            if is_media:
                                if rid in used_rids:
                                    norm_target = os.path.normpath(
                                        os.path.join(os.path.dirname(rels_name).replace("_rels", ""), target)
                                    ).replace("\\", "/")
                                    if norm_target.startswith("/"):
                                        norm_target = norm_target[1:]
                                    keep_targets.add(norm_target)
                                else:
                                    root.remove(rel)
                                    modified = True
                            else:
                                norm_target = os.path.normpath(
                                    os.path.join(os.path.dirname(rels_name).replace("_rels", ""), target)
                                ).replace("\\", "/")
                                if norm_target.startswith("/"):
                                    norm_target = norm_target[1:]
                                keep_targets.add(norm_target)

                        if modified:
                            rels_to_update[rels_name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                    except Exception:
                        pass

                with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        fname = item.filename
                        if fname in rels_to_update:
                            zout.writestr(fname, rels_to_update[fname])
                        elif fname.startswith("word/media/"):
                            if fname in keep_targets:
                                zout.writestr(fname, zin.read(fname))
                        else:
                            zout.writestr(fname, zin.read(fname))

            if os.path.exists(temp_path):
                shutil.move(temp_path, docx_path)
        except Exception:
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass

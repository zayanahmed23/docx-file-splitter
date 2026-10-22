"""
word_splitter_cli.py
====================
Interactive Command-Line Interface for the Word File Splitter.
Prompts the user for the .docx file and the number of parts to split it into.
"""

import os
import sys
import argparse
from typing import Optional
from word_splitter_core import WordSplitter


def format_size(bytes_num: int) -> str:
    """Format bytes to human readable string."""
    for unit in ["B", "KB", "MB", "GB"]:
        if bytes_num < 1024.0:
            return f"{bytes_num:.1f} {unit}" if unit != "B" else f"{bytes_num} B"
        bytes_num /= 1024.0
    return f"{bytes_num:.1f} TB"


def clean_input_path(raw_path: str) -> str:
    """Strip quotes and whitespace from drag-and-dropped paths in terminal."""
    path = raw_path.strip()
    if (path.startswith('"') and path.endswith('"')) or (path.startswith("'") and path.endswith("'")):
        path = path[1:-1].strip()
    return os.path.abspath(os.path.expanduser(path))


def pick_file_via_dialog() -> Optional[str]:
    """Launch native OS file picker dialog if GUI display is available."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected = filedialog.askopenfilename(
            title="Select Word Document (.docx) to Split",
            filetypes=[("Word Document (*.docx)", "*.docx"), ("All Files (*.*)", "*.*")]
        )
        root.destroy()
        return selected if selected else None
    except Exception:
        return None


def prompt_for_file() -> str:
    """Interactive prompt for Word document path with drag & drop and file browser."""
    while True:
        try:
            print("\n" + "=" * 64)
            print("  Step 1: Select Word Document")
            print("=" * 64)
            print("📂 Enter path to your Word document (.docx), drag & drop it here,")
            raw = input("   or press Enter to open File Browser: ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\nOperation cancelled.")
            sys.exit(0)

        if not raw:
            print("   Opening file browser...")
            chosen = pick_file_via_dialog()
            if chosen:
                clean_path = clean_input_path(chosen)
                if os.path.exists(clean_path) and clean_path.lower().endswith(".docx"):
                    return clean_path
                elif clean_path.lower().endswith(".doc"):
                    print("⚠️  This is an older Word 97-2003 format (.doc).")
                    print("   Please open it in Microsoft Word and 'Save As' .docx first.")
                    continue
            else:
                print("⚠️  No file selected. Please enter a path or drag & drop a .docx file.")
                continue

        clean_path = clean_input_path(raw)

        if not os.path.exists(clean_path):
            print(f"❌ File not found at: {clean_path}")
            print("   Please check the path and try again.")
            continue

        if not clean_path.lower().endswith(".docx"):
            if clean_path.lower().endswith(".doc"):
                print("⚠️  This is an older Word 97-2003 format (.doc).")
                print("   Please open it in Microsoft Word and 'Save As' .docx first.")
            else:
                print(f"❌ Unsupported file type. Please provide a Word document (.docx).")
            continue

        return clean_path


def prompt_for_num_segments(max_blocks: int) -> int:
    """Interactive prompt for number of segments based on memory distribution."""
    while True:
        try:
            print("\n" + "=" * 64)
            print("  Step 2: Number of Segments (Memory Distribution)")
            print("=" * 64)
            print("⚡ Strategy: Memory Distribution")
            print("   (Balances embedded images & content weight so each segment has equal MB)")
            raw = input(f"\n✂️  Into how many segments do you want to split it? [e.g. 2, 3, 5] (Default: 2): ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\nOperation cancelled.")
            sys.exit(0)

        if not raw:
            return 2

        try:
            parts = int(raw)
            if parts < 1:
                print("⚠️  Number of segments must be at least 1.")
                continue
            if parts == 1:
                print("ℹ️  Splitting into 1 segment will create a single clone of the document.")
                return 1
            if parts > max_blocks:
                print(f"⚠️  Document only has {max_blocks} content blocks (paragraphs/tables).")
                print(f"   The document will be split into at most {max_blocks} segments.")
                return max_blocks
            return parts
        except ValueError:
            print("❌ Invalid input. Please enter a valid whole number (e.g., 2, 3, 4).")


def prompt_for_method(headings_count: int) -> str:
    """Advanced interactive prompt for splitting method."""
    print("\n⚙️  Select Splitting Method:")
    print("   [1] Equal Memory / File Size (Recommended - divides into equal MB)")
    print("   [2] Balanced Word Count (even text distribution)")
    print("   [3] Equal Content Blocks (Paragraphs & Tables)")
    if headings_count > 1:
        print(f"   [4] By Headings ({headings_count} chapters/headings found)")

    try:
        choice = input("   Press Enter for [1], or enter choice (1-4): ").strip()
    except (KeyboardInterrupt, EOFError):
        print("\nOperation cancelled.")
        sys.exit(0)

    if choice == "2":
        return "words"
    elif choice == "3":
        return "blocks"
    elif choice == "4" and headings_count > 1:
        return "headings"
    return "memory"


def run_interactive(advanced: bool = False):
    """
    Main interactive loop.
    Easy 2-step wizard by default:
      1. Word file (.docx)
      2. Number of segments (Memory Distribution)
    """
    print("=" * 64)
    print("              📄 WORD FILE SPLITTER ✂️              ")
    print("=" * 64)
    print("Splits Microsoft Word (.docx) documents by Memory Distribution")
    print("while preserving 100% of formatting, tables, images, and styles.\n")

    # Step 1: Prompt for file
    file_path = prompt_for_file()

    # Step 2: Analyze file
    print("\n🔍 Analyzing document structure and memory footprint...")
    try:
        analysis = WordSplitter.analyze_document(file_path)
    except Exception as e:
        print(f"❌ Error reading document: {e}")
        sys.exit(1)

    print("\n" + "=" * 64)
    print(f"  📄 File:        {analysis['file_name']}")
    print(f"  💾 Total Size:  {format_size(analysis['file_size'])}")
    print(f"  📝 Words:       {analysis['total_words']:,}")
    print(f"  📦 Blocks:      {analysis['total_blocks']:,} ({analysis['total_paragraphs']} paragraphs, {analysis['total_tables']} tables)")
    if analysis.get('headings'):
        print(f"  🔖 Headings:    {len(analysis['headings'])} found")
    print("=" * 64)

    if analysis['total_blocks'] == 0:
        print("❌ This document contains no readable paragraphs or tables to split.")
        sys.exit(1)

    # Step 3: Prompt for number of segments
    num_segments = prompt_for_num_segments(analysis['total_blocks'])

    # Method & output folder
    method = "memory"
    base_name, _ = os.path.splitext(analysis['file_name'])
    default_out = os.path.join(os.path.dirname(file_path), f"{base_name}_split")
    output_dir = default_out

    if advanced:
        method = prompt_for_method(len(analysis.get('headings', [])))
        print(f"\n📁 Output directory:")
        print(f"   Default: {default_out}")
        try:
            custom_out = input("   Press Enter to use default, or enter custom path: ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\nOperation cancelled.")
            sys.exit(0)
        if custom_out:
            output_dir = clean_input_path(custom_out)

    # Step 4: Perform splitting with live progress
    print("\n" + "-" * 64)
    print(f"🚀 Splitting '{analysis['file_name']}' into {num_segments} segments using '{method}' distribution...")
    print("-" * 64)

    def on_progress(current: int, total: int, part_path: str):
        fname = os.path.basename(part_path)
        pct = int((current / total) * 100)
        bar_len = 24
        filled = int(bar_len * current / total)
        bar = "█" * filled + "░" * (bar_len - filled)
        print(f"  [{bar}] {pct}% - Created Segment {current}/{total}: {fname}")

    try:
        result = WordSplitter.split_document(
            file_path=file_path,
            num_parts=num_segments,
            output_dir=output_dir,
            method=method,
            progress_callback=on_progress,
        )
    except Exception as e:
        print(f"\n❌ Error during splitting: {e}")
        sys.exit(1)

    # Step 5: Results summary
    print("\n" + "=" * 64)
    print("✨ SPLIT COMPLETED SUCCESSFULLY!")
    print("=" * 64)
    print(f"📁 Destination Folder:\n   {result['output_directory']}\n")

    print(f"{'Segment':<10} | {'Filename':<32} | {'Words':>8} | {'Size':>10}")
    print("-" * 68)
    for p in result['parts']:
        print(f"Part {p['part_number']:<5} | {p['file_name']:<32} | {p['words']:>8,} | {format_size(p['file_size']):>10}")
    print("-" * 68)

    # Step 6: Offer to open folder
    try:
        open_folder = input("\n📂 Open output folder in File Explorer? [Y/n]: ").strip().lower()
        if open_folder in ("", "y", "yes"):
            if sys.platform == "win32":
                os.startfile(result['output_directory'])
            elif sys.platform == "darwin":
                import subprocess
                subprocess.run(["open", result['output_directory']])
            else:
                import subprocess
                subprocess.run(["xdg-open", result['output_directory']])
    except Exception:
        pass

    print("\nThank you for using Word File Splitter!\n")


def main():
    parser = argparse.ArgumentParser(
        description="Split a Microsoft Word (.docx) document into N segments based on memory distribution."
    )
    parser.add_argument("file", nargs="?", help="Path to input .docx document")
    parser.add_argument("parts", nargs="?", type=int, help="Number of segments to split into")
    parser.add_argument("-f", "--file-path", dest="flag_file", help="Path to input .docx document")
    parser.add_argument("-n", "--num-parts", "--num-segments", dest="flag_parts", type=int, help="Number of segments")
    parser.add_argument("-m", "--method", choices=["memory", "words", "blocks", "headings"], default="memory", help="Split method (default: memory)")
    parser.add_argument("-o", "--output", help="Output directory")
    parser.add_argument("-a", "--advanced", action="store_true", help="Launch interactive mode with advanced options")

    args = parser.parse_args()

    input_file = args.flag_file or args.file
    input_parts = args.flag_parts or args.parts

    # If no file was supplied as argument, run interactive mode
    if not input_file:
        run_interactive(advanced=args.advanced)
        return

    # Non-interactive CLI mode
    input_file = clean_input_path(input_file)
    if not os.path.isfile(input_file):
        print(f"❌ File not found: {input_file}", file=sys.stderr)
        sys.exit(1)

    analysis = WordSplitter.analyze_document(input_file)

    if not input_parts:
        input_parts = prompt_for_num_segments(analysis["total_blocks"])

    output_dir = args.output
    if output_dir:
        output_dir = clean_input_path(output_dir)

    print(f"Splitting '{analysis['file_name']}' into {input_parts} segments using method '{args.method}'...")
    result = WordSplitter.split_document(
        file_path=input_file,
        num_parts=input_parts,
        output_dir=output_dir,
        method=args.method,
    )
    print(f"Successfully generated {len(result['parts'])} segments in: {result['output_directory']}")
    for p in result["parts"]:
        print(f" - Part {p['part_number']}: {p['file_name']} ({p['words']:,} words, {format_size(p['file_size'])})")


if __name__ == "__main__":
    main()

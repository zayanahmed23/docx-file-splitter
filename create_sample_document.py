"""
create_sample_document.py
=========================
Generates a realistic sample Word document with multiple chapters,
formatting, tables, and bullet lists for testing the Word File Splitter.
"""

import os
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def create_sample(path: str = "example_document.docx"):
    doc = docx.Document()

    # Title
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title.add_run("The Complete Guide to Artificial Intelligence Systems")
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(30, 58, 138)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = sub.add_run("Volume I: Foundations, Architectures, and Modern Applications")
    run_sub.font.size = Pt(13)
    run_sub.font.italic = True
    run_sub.font.color.rgb = RGBColor(100, 116, 139)

    doc.add_paragraph()  # spacer

    # Chapter 1
    doc.add_heading("Chapter 1: Introduction to Machine Intelligence", level=1)
    p1 = doc.add_paragraph(
        "Artificial Intelligence represents one of the most transformative technological developments "
        "of the twenty-first century. Rooted in logic, computational linguistics, and neurobiology, modern AI systems "
        "have transitioned from theoretical laboratory experiments into fundamental infrastructure supporting global commerce, "
        "medicine, engineering, and scientific research."
    )
    p2 = doc.add_paragraph(
        "Historically, early computing pioneers envisioned machines capable of general inference. "
        "Alan Turing proposed the imitation test in 1950, while John McCarthy coined the term 'artificial intelligence' "
        "in 1956 during the Dartmouth Summer Research Project on Artificial Intelligence."
    )
    r2 = p2.add_run(" Early symbolic systems demonstrated impressive performance on formal mathematics and logic puzzles.")
    r2.bold = True

    # Chapter 2
    doc.add_heading("Chapter 2: Deep Learning and Neural Network Architectures", level=1)
    doc.add_paragraph(
        "The renaissance of deep learning over the past decade can be attributed to three converging factors: "
        "the availability of vast digitized datasets, algorithmic advancements in optimization techniques, "
        "and parallel processing hardware such as modern Graphical Processing Units (GPUs) and Tensor Processing Units (TPUs)."
    )

    # Table of Network Types
    table = doc.add_table(rows=4, cols=3)
    table.style = 'Table Grid'
    headers = ["Architecture", "Primary Use Case", "Key Characteristics"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        for p in cell.paragraphs:
            for r in p.runs:
                r.bold = True

    rows_data = [
        ("Feedforward Dense", "Tabular and structured data", "Fully connected layers, nonlinear activations"),
        ("Convolutional (CNN)", "Computer vision and image analysis", "Spatial hierarchies, translation equivariance"),
        ("Transformers", "Large language models and multimodal", "Multi-head self-attention, positional encoding")
    ]
    for row_idx, row_values in enumerate(rows_data, start=1):
        for col_idx, val in enumerate(row_values):
            table.cell(row_idx, col_idx).text = val

    doc.add_paragraph()  # spacer

    p3 = doc.add_paragraph(
        "Transformers in particular have revolutionized sequential and contextual reasoning. By replacing recurrent structures "
        "with self-attention mechanisms, modern foundation models can process entire contexts simultaneously, unlocking emergent "
        "zero-shot reasoning capabilities across a wide array of domains."
    )

    # Chapter 3
    doc.add_heading("Chapter 3: Practical Implementation and Operational Deployment", level=1)
    doc.add_paragraph(
        "Deploying intelligence models at scale requires robust operational pipelines, often referred to as MLOps. "
        "Key engineering challenges include latency minimization, quantizing float weights to int8/int4 representations, "
        "managing key-value caches in autoregressive inference, and orchestrating distributed microservices."
    )
    doc.add_paragraph(
        "Monitoring production drift and maintaining retrieval augmentations ensures systems stay factually grounded, "
        "minimizing hallucinations while providing transparent citations to authoritative enterprise data stores."
    )

    # Chapter 4
    doc.add_heading("Chapter 4: Ethics, Alignment, and Future Horizons", level=1)
    doc.add_paragraph(
        "As computational models assume greater autonomy, aligning agentic workflows with human values and safety standards "
        "is paramount. Techniques such as Reinforcement Learning from Human Feedback (RLHF), constitutional fine-tuning, "
        "and verifiable sandboxing provide guardrails against unintended behavior."
    )
    doc.add_paragraph(
        "Looking forward, the frontier of AI research is focusing on autonomous reasoning loops, test-time compute scaling, "
        "neuro-symbolic integration, and automated scientific discovery across chemistry, biology, and materials science."
    )

    doc.save(path)
    print(f"Sample document created successfully at: {os.path.abspath(path)}")

if __name__ == "__main__":
    create_sample()

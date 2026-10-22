"""
splitter.py
===========
Unified entry point for Word Document Splitter.
Intelligently handles GUI mode, CLI mode, headless environments, and flags.

Usage:
  python splitter.py                     # Launches Desktop GUI (or CLI fallback in headless mode)
  python splitter.py --gui (-g)          # Explicitly launches Desktop GUI
  python splitter.py --cli (-c)          # Explicitly launches Interactive Terminal CLI
  python splitter.py doc.docx 3          # Splits doc.docx into 3 parts
  python splitter.py -f doc.docx -n 3    # Non-interactive CLI split with options
  python splitter.py --help (-h)         # Displays help and usage guide
"""

import sys
import os

HELP_TEXT = """Word Document Splitter
======================
Fast, lossless Microsoft Word (.docx) splitter that preserves 100% of formatting,
styles, tables, headers/footers, and embedded images.

Usage:
  python splitter.py [options]
  python splitter.py <file> <parts>

Options:
  -h, --help                 Show this help message and exit
  -g, --gui                  Launch the desktop graphical interface (GUI)
  -c, --cli                  Launch the interactive terminal interface (CLI)
  -f, --file-path PATH       Path to the input .docx file
  -n, --num-parts N          Number of parts to split into
  -m, --method METHOD        Splitting strategy:
                             'memory'   - Equal file size / image weight (default)
                             'words'    - Balanced word count
                             'blocks'   - Equal content paragraphs & tables
                             'headings' - Split by major headings / chapters
  -o, --output DIR           Custom output directory for split files

Examples:
  python splitter.py                         # Launch GUI
  python splitter.py --cli                   # Interactive CLI prompt
  python splitter.py report.docx 4           # Split into 4 parts
  python splitter.py -f book.docx -n 3 -m words
"""


def main():
    # If help flag requested
    if "-h" in sys.argv or "--help" in sys.argv:
        print(HELP_TEXT)
        return

    # If explicit CLI flag is present
    if "--cli" in sys.argv or "-c" in sys.argv:
        sys.argv = [arg for arg in sys.argv if arg not in ("--cli", "-c")]
        import word_splitter_cli
        word_splitter_cli.main()
        return

    # If explicit GUI flag is present
    if "--gui" in sys.argv or "-g" in sys.argv:
        import word_splitter_gui
        word_splitter_gui.main()
        return

    # If CLI arguments or options are passed
    cli_flags = {"-f", "--file-path", "-n", "--num-parts", "--num-segments", "-m", "--method", "-o", "--output", "-a", "--advanced"}
    has_cli_flag = any(arg in cli_flags for arg in sys.argv[1:])
    has_positional = len(sys.argv) > 1 and not sys.argv[1].startswith("-")

    if has_cli_flag or has_positional:
        import word_splitter_cli
        word_splitter_cli.main()
        return

    # Otherwise, by default, try to launch the GUI (with headless fallback)
    try:
        import word_splitter_gui
        word_splitter_gui.main()
    except Exception as e:
        print(f"Notice: GUI display unavailable ({e}). Falling back to interactive CLI...\n")
        import word_splitter_cli
        word_splitter_cli.run_interactive()


if __name__ == "__main__":
    main()

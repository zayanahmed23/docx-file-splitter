"""
main.py
=======
Convenience entry point for Word Document Splitter.
Runs the easy interactive prompt by default, or accepts CLI/GUI flags.
"""
import sys

def main():
    if "--gui" in sys.argv or "-g" in sys.argv:
        import word_splitter_gui
        word_splitter_gui.main()
    else:
        import word_splitter_cli
        word_splitter_cli.main()

if __name__ == "__main__":
    main()

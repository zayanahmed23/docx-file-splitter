# Contributing to Word Document Splitter

Thank you for your interest in contributing to **Word Document Splitter**! We welcome community contributions, bug reports, and feature suggestions.

## 🤝 Code of Conduct
Please be respectful, kind, and collaborative in all interactions.

## 🛠️ Development Setup

1. **Fork and Clone the Repository**:
   ```bash
   git clone https://github.com/your-username/Word-file-splitter.git
   cd Word-file-splitter
   ```

2. **Create a Virtual Environment**:
   ```bash
   python -m venv .venv
   # On Windows:
   .venv\Scripts\activate
   # On macOS/Linux:
   source .venv/bin/activate
   ```

3. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   pip install -e .
   ```

4. **Run the Test Suite**:
   ```bash
   python test_splitter.py
   ```

5. **Generate a Sample Document for Manual Testing**:
   ```bash
   python create_sample_document.py
   ```

## 🔒 Privacy & Document Security Rule
- **Never commit `.docx`, `.doc`, or `.pdf` files containing personal or proprietary data.**
- The `.gitignore` is configured to ignore all `.docx`, `.doc`, and output directories automatically.
- Always use `create_sample_document.py` or synthetic test data when writing new tests or reproducing issues.

## 📬 Submitting Changes
1. Create a feature branch (`git checkout -b feature/amazing-feature`).
2. Make your improvements with clean, readable code and comments.
3. Ensure all tests pass: `python test_splitter.py`.
4. Commit your changes (`git commit -m "feat: add amazing feature"`).
5. Push to your branch and open a Pull Request.

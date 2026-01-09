# Contributing to CSV to XLSX Converter

First off, thank you for considering contributing to CSV to XLSX Converter! 🎉

## Code of Conduct

This project and everyone participating in it is governed by our Code of Conduct. By participating, you are expected to uphold this code.

## How Can I Contribute?

### Reporting Bugs

Before creating bug reports, please check the existing issues as you might find out that you don't need to create one.

**When you are creating a bug report, please include as many details as possible:**

- **Use a clear and descriptive title**
- **Describe the exact steps to reproduce the problem**
- **Provide specific examples**
- **Describe the behavior you observed and what you expected**
- **Include screenshots if possible**
- **Include your environment details:**
  - Windows version
  - Python version
  - Package versions (`pip freeze`)

### Suggesting Enhancements

Enhancement suggestions are tracked as GitHub issues.

**When creating an enhancement suggestion, please include:**

- **Use a clear and descriptive title**
- **Provide a step-by-step description of the suggested enhancement**
- **Explain why this enhancement would be useful**
- **List any alternatives you've considered**

### Pull Requests

1. Fork the repo and create your branch from `main`
2. If you've added code that should be tested, add tests
3. Ensure your code follows the existing style
4. Make sure your code lints
5. Issue that pull request!

## Development Setup

1. Fork and clone the repository
```bash
git clone https://github.com/yourusername/csv-to-xlsx-converter.git
cd csv-to-xlsx-converter
```

2. Create a virtual environment
```bash
python -m venv venv
venv\Scripts\activate  # Windows
```

3. Install dependencies
```bash
pip install -r requirements.txt
pip install -r requirements-dev.txt  # If available
```

4. Run the application
```bash
python main.py
```

## Style Guidelines

### Python Style Guide

- Follow PEP 8
- Use meaningful variable names
- Add docstrings to functions and classes
- Keep functions small and focused
- Use type hints where appropriate

### Git Commit Messages

- Use the present tense ("Add feature" not "Added feature")
- Use the imperative mood ("Move cursor to..." not "Moves cursor to...")
- Limit the first line to 72 characters or less
- Reference issues and pull requests liberally after the first line

### Example commit messages:
```
Add zebra stripe styling option

- Implement alternating row colors
- Add checkbox in styling tab
- Update documentation

Fixes #123
```

## Project Structure

```
csv-to-xlsx-converter/
├── main.py              # Main application entry point
├── requirements.txt     # Production dependencies
├── README.md           # Project documentation
├── LICENSE             # MIT License
├── .gitignore          # Git ignore rules
├── CHANGELOG.md        # Version history
├── CONTRIBUTING.md     # This file
├── setup.py            # Package setup
├── build_exe.bat       # Windows build script
└── run.bat             # Windows run script
```

## Testing

Currently, the project doesn't have automated tests. This is an area where contributions would be very welcome!

Suggested testing areas:
- CSV parsing with various encodings
- Excel file generation
- UI component functionality
- Error handling

## Documentation

Improvements to documentation are always welcome. This includes:
- README.md
- Code comments
- Docstrings
- Wiki pages

## Questions?

Feel free to open an issue with your question or reach out to the maintainers.

Thank you for contributing! 🙏

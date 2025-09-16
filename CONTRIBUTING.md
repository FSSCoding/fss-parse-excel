# Contributing to Excel Toolkit

Thank you for considering contributing to the Excel Toolkit! This document outlines the guidelines for contributing to this project.

## 🚀 Quick Start

1. Fork the repository
2. Clone your fork: `git clone https://github.com/yourusername/fss-parse-excel.git`
3. Create a feature branch: `git checkout -b feature/amazing-feature`
4. Make your changes
5. Run tests: `python -m pytest tests/`
6. Commit your changes: `git commit -m 'Add amazing feature'`
7. Push to your branch: `git push origin feature/amazing-feature`
8. Open a Pull Request

## 🏗️ Development Setup

### Prerequisites
- Python 3.8 or higher
- pip package manager
- Virtual environment (recommended)

### Installation
```bash
# Clone the repository
git clone https://github.com/FSSCoding/fss-parse-excel.git
cd fss-parse-excel

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install in development mode
pip install -e .
pip install -e ".[dev]"

# Run tests
python -m pytest tests/
```

## 📝 Code Style

We follow Python best practices and use automated tools:

### Formatting
- **Black**: Code formatting
- **isort**: Import sorting
- **flake8**: Linting

```bash
# Format code
black src/ tests/
isort src/ tests/

# Check linting
flake8 src/ tests/
```

### Type Hints
- Use type hints for all public APIs
- Run mypy for type checking: `mypy src/`

## 🧪 Testing

### Running Tests
```bash
# Run all tests
python -m pytest tests/

# Run with coverage
python -m pytest tests/ --cov=src --cov-report=html

# Run specific test file
python -m pytest tests/test_cell_manager.py -v
```

### Writing Tests
- Write tests for all new features
- Maintain test coverage above 90%
- Use descriptive test names
- Include edge cases and error conditions

Example test structure:
```python
def test_set_cell_value_success():
    """Test successful cell value setting."""
    # Arrange
    manager = CellManager("test.xlsx")
    
    # Act
    result = manager.set_cell_value("A1", "test_value")
    
    # Assert
    assert result is True
    assert manager.get_cell_value("A1") == "test_value"
```

## 📚 Documentation

### Docstrings
Use Google-style docstrings:

```python
def process_range(self, range_ref: str, operation: str) -> bool:
    """Process a range of cells with the specified operation.
    
    Args:
        range_ref: Excel range reference (e.g., 'A1:C10')
        operation: Operation to perform ('clear', 'format', etc.)
        
    Returns:
        True if successful, False otherwise.
        
    Raises:
        ValueError: If range_ref is invalid.
        FileNotFoundError: If Excel file doesn't exist.
    """
```

### README Updates
- Update README.md for new features
- Include usage examples
- Update the feature list

## 🎯 Feature Guidelines

### New Features
- Must be CLI agent focused
- Include comprehensive tests
- Add documentation and examples
- Follow the existing architecture patterns

### Architecture Principles
1. **Safety First**: Always use safety mechanisms
2. **Precise Control**: Support exact cell/range operations  
3. **Error Handling**: Graceful failure with clear messages
4. **Performance**: Efficient operations for large files
5. **CLI Friendly**: Rich console output and clear interfaces

## 🐛 Bug Reports

When reporting bugs, include:

1. **Environment**: Python version, OS, package versions
2. **Reproduction**: Minimal code to reproduce the issue
3. **Expected vs Actual**: What you expected vs what happened
4. **Error Messages**: Full traceback if applicable

Template:
```markdown
## Environment
- Python: 3.9.7
- OS: Ubuntu 20.04
- Excel Toolkit: 1.0.0

## Reproduction
```python
from excel_toolkit import CellManager
manager = CellManager("test.xlsx")
# Steps that cause the issue
```

## Expected Behavior
Should do X

## Actual Behavior  
Does Y instead

## Error Messages
```
Traceback (most recent call last):
  ...
```
```

## 📋 Pull Request Guidelines

### Before Submitting
- [ ] Tests pass locally
- [ ] Code follows style guidelines  
- [ ] Documentation is updated
- [ ] CHANGELOG.md is updated (if applicable)
- [ ] Commit messages are descriptive

### PR Template
```markdown
## Description
Brief description of changes

## Type of Change
- [ ] Bug fix
- [ ] New feature  
- [ ] Breaking change
- [ ] Documentation update

## Testing
- [ ] Unit tests added/updated
- [ ] Integration tests pass
- [ ] Manual testing completed

## Checklist
- [ ] Code follows style guidelines
- [ ] Self-review completed
- [ ] Documentation updated
- [ ] No breaking changes (or clearly marked)
```

## 🏷️ Versioning

We use [Semantic Versioning](https://semver.org/):
- **MAJOR**: Breaking changes
- **MINOR**: New features (backward compatible)
- **PATCH**: Bug fixes (backward compatible)

## 📞 Getting Help

- **Issues**: GitHub Issues for bugs and feature requests
- **Discussions**: GitHub Discussions for questions and ideas
- **Email**: development@fsscoding.com

## 🎖️ Recognition

Contributors will be acknowledged in:
- CONTRIBUTORS.md file
- Release notes for significant contributions
- Special thanks in documentation

## 📜 Code of Conduct

This project follows the [Contributor Covenant](https://www.contributor-covenant.org/) Code of Conduct. Please read and follow these guidelines to ensure a welcoming environment for all contributors.

---

**Thank you for contributing to Excel Toolkit!** 🚀
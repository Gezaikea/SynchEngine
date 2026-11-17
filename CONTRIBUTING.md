# Contributing to PPT ↔ Excel Sync Engine

Thank you for your interest in contributing! Here's how to get started.

## 🤝 Code of Conduct

Be respectful, inclusive, and collaborative. We're all here to build something great together.

## 🐛 Reporting Issues

Found a bug? Help us fix it!

### Before Submitting
- Check [existing issues](../../issues) to avoid duplicates
- Test with the latest version
- Gather relevant information

### Issue Template
```
**Description**
Clear description of the problem

**Steps to Reproduce**
1. Step 1
2. Step 2
3. ...

**Expected Behavior**
What should happen

**Actual Behavior**
What actually happens

**Environment**
- Python version: 3.x
- OS: Windows/Mac/Linux
- PPT/Excel versions: ...
```

## 💡 Suggesting Features

Have a great idea? We'd love to hear it!

### Feature Request Template
```
**Problem**
What problem does this solve?

**Solution**
How should it work?

**Alternative Solutions**
Any other approaches?

**Additional Context**
Any other information?
```

## 🛠️ Development Setup

### Fork & Clone
```bash
# Fork the repository on GitHub
# Clone your fork
git clone https://github.com/YOUR-USERNAME/ppt-excel-sync.git
cd ppt-excel-sync

# Add upstream remote
git remote add upstream https://github.com/original-owner/ppt-excel-sync.git
```

### Create Development Environment
```bash
# Create virtual environment
python -m venv venv

# Activate it
# Windows:
venv\Scripts\activate
# Mac/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Install development dependencies
pip install -r requirements-dev.txt
```

## 📝 Making Changes

### Code Style

We follow PEP 8 with these guidelines:

**Python**
```python
# Good
def sync_ppt_to_excel(ppt_path: str, excel_path: str) -> bool:
    """Sync PPT data to Excel."""
    reader = PPTReader(ppt_path)
    data = reader.read()
    return True

# Avoid
def sync(p, e):
    # Unclear variable names
    r = PPTReader(p)
    d = r.read()
    return 1
```

**Comments**
```python
# Good: Explain WHY, not WHAT
# We use this field mapping to handle legacy PPT naming conventions
PPT_FIELD_MAP = {...}

# Avoid: Just repeating the code
# This is the field map
FIELD_MAP = {...}
```

### Type Hints

Always include type hints:
```python
def clean_text(value: str, level: str = 'standard') -> str:
    """Clean and normalize text input."""
    return value.strip()

def detect_changes(ppt_data: Dict, excel_data: Dict) -> List[Record]:
    """Detect changes between sources."""
    return []
```

### Documentation

```python
def write_product(self, product_id: str, product_data: Dict[str, Any]) -> bool:
    """
    Write product record to Excel.
    
    Args:
        product_id: Unique product identifier
        product_data: Dictionary of field names to values
    
    Returns:
        True if successful, False otherwise
    
    Raises:
        FileNotFoundError: If Excel file not found
        ValueError: If product_id is invalid
    
    Example:
        >>> writer = ExcelWriter('workbook.xlsx')
        >>> writer.open()
        >>> writer.write_product('prod_1', {'name': 'Product 1'})
        True
    """
```

## ✅ Creating a Pull Request

### Before Starting
1. Create an issue first (discuss the change)
2. Wait for feedback
3. Create a feature branch

### Commit Workflow
```bash
# Create feature branch
git checkout -b feature/add-data-validation

# Make changes
# Test thoroughly
# Commit with clear messages
git add .
git commit -m "Add data validation layer for percentages and choices"

# Keep commits atomic and logical
git commit -m "Extract percentage numbers from PPT text fields"
git commit -m "Add placeholder text detection"
git commit -m "Add text cleaning for special characters"

# Push to your fork
git push origin feature/add-data-validation
```

### Commit Message Format
```
[TYPE] Brief description

Longer explanation of what changed and why.

- Point 1
- Point 2

Fixes #123
```

**Types:**
- `[FEATURE]` - New feature
- `[FIX]` - Bug fix
- `[DOCS]` - Documentation
- `[REFACTOR]` - Code refactoring
- `[TEST]` - Tests
- `[PERF]` - Performance improvement

### Pull Request Template
```markdown
## Description
Clear description of changes

## Related Issue
Fixes #123

## Type of Change
- [ ] Bug fix
- [ ] New feature
- [ ] Breaking change
- [ ] Documentation update

## How to Test
Steps to verify the changes

## Checklist
- [ ] Code follows style guidelines
- [ ] Comments added for complex logic
- [ ] Documentation updated
- [ ] Tests pass locally
- [ ] No new warnings generated
```

## 🧪 Testing

### Running Tests
```bash
# Run all tests
python -m pytest

# Run specific test file
python -m pytest tests/test_sync_engine.py

# Run with coverage
pytest --cov=sync_engine tests/
```

### Writing Tests
```python
import unittest
from sync_engine_v2 import PPTReader

class TestPPTReader(unittest.TestCase):
    
    def setUp(self):
        self.ppt_path = 'tests/fixtures/sample.pptx'
        self.reader = PPTReader(self.ppt_path)
    
    def test_read_extracts_all_fields(self):
        """Test that read() extracts all 47 fields"""
        data = self.reader.read()
        self.assertEqual(len(data), 47)
    
    def test_read_handles_empty_ppt(self):
        """Test that read() handles empty PPT gracefully"""
        empty_ppt = 'tests/fixtures/empty.pptx'
        reader = PPTReader(empty_ppt)
        data = reader.read()
        self.assertEqual(data, {})
```

## 📚 Documentation

### Updating README
- Keep it concise and up-to-date
- Add examples for new features
- Update the roadmap

### Adding Documentation
1. Create file in `docs/` folder
2. Update index in main README
3. Include examples and links

### Code Examples
Always include working examples:
```python
# Example usage
from sync_engine_v2 import PPTReader

reader = PPTReader('template.pptx')
data = reader.read()
print(data)
```

## 📋 Checklist Before Submitting PR

- [ ] Code follows style guidelines (PEP 8)
- [ ] Type hints added
- [ ] Docstrings written
- [ ] Tests added/updated
- [ ] All tests pass
- [ ] No new warnings
- [ ] Documentation updated
- [ ] Commit messages clear
- [ ] No unrelated changes
- [ ] No merge conflicts

## 🚀 Release Process

Maintainers will:
1. Review and test PR thoroughly
2. Merge to main branch
3. Update version number
4. Create release notes
5. Publish to PyPI (future)

## 🎓 Project Structure

```
ppt-excel-sync/
├── sync_engine_v2.py           # Core reader & change detector
├── sync_engine_writer.py       # Excel writer
├── sync_full.py               # Full automation
├── sync_engine_validation.py  # Data validation (new feature)
├── docs/
│   ├── ARCHITECTURE.md
│   ├── USAGE_GUIDE.md
│   └── API_REFERENCE.md
├── tests/
│   ├── test_sync_engine_v2.py
│   ├── test_sync_writer.py
│   ├── fixtures/
│   │   ├── sample.pptx
│   │   └── sample.xlsx
│   └── __init__.py
├── requirements.txt
├── requirements-dev.txt
└── README.md
```

## 🏆 Areas to Contribute

### High Priority
- [ ] Data validation layer (`sync_engine_validation.py`)
- [ ] Unit tests (90%+ coverage)
- [ ] Error handling improvements
- [ ] Documentation

### Medium Priority
- [ ] Performance optimization
- [ ] Additional field types
- [ ] Better logging
- [ ] Configuration options

### Nice to Have
- [ ] CLI improvements
- [ ] GUI prototype
- [ ] Integration examples
- [ ] Video tutorials

## 💬 Getting Help

- **Questions?** Use GitHub Discussions
- **Stuck?** Check existing issues and docs
- **Need feedback?** Create a draft PR early

## 📞 Maintainers

- **Primary:** Gezaikea (@Gezaikea)

## ⭐ Thank You!

Thank you for contributing to making this project better!

---

**Happy coding!** 🚀

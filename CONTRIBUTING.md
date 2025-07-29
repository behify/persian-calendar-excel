# Contributing to Persian Calendar Excel

<div align="center">
  <img src="https://raw.githubusercontent.com/behify/persian-calendar-excel/main/assets/logo.png" alt="Persian Calendar Excel Logo" width="150">
</div>

Thank you for your interest in contributing to this project!

## How to Contribute

### 1. Reporting Issues

If you find a bug:
- Check the [Issues page](https://github.com/behify/persian-calendar-excel/issues) first
- Look for similar existing issues
- Create a new issue with complete details

**Issue Template:**
```
### Bug Description
Clear description of the bug

### Steps to Reproduce
1. First step
2. Second step
3. Unexpected result

### Expected Behavior
What should have happened

### Environment
- Excel Version: 
- Windows/Mac Version: 
- Package Version: 
```

### 2. Feature Requests

For new features:
- First discuss in [Discussions](https://github.com/behify/persian-calendar-excel/discussions)
- After discussion, create an Issue with `enhancement` label

### 3. Code Contributions

#### Initial Steps
1. **Fork** the repository
2. **Clone** your fork:
   ```bash
   git clone https://github.com/your-username/persian-calendar-excel.git
   ```
3. Create a new **branch**:
   ```bash
   git checkout -b feature/amazing-feature
   ```

#### Coding Standards

##### VBA Code Style
```vba
' Use clear, descriptive function names
Public Function CalculatePersianDays() As Long
    ' Add comments for complex logic
    Dim totalDays As Long
    
    ' Use meaningful variable names
    For yearIndex = 1 To targetYear
        ' Implementation
    Next yearIndex
End Function
```

##### File Structure
- Each module in separate `.bas` file
- Clear and logical naming
- Complete header documentation in each file

#### Testing

Before submitting Pull Request:
1. Run all existing tests:
   ```vba
   TestPersianCalendarFunctions()
   ```
2. Add tests for new features
3. Ensure no existing functionality is broken

#### Pull Request

1. Commit your changes:
   ```bash
   git commit -m "Add: Description of changes"
   ```

2. Push to your branch:
   ```bash
   git push origin feature/amazing-feature
   ```

3. Create Pull Request with:
   - Clear title
   - Complete description of changes
   - List of tests performed
   - Screenshots/examples if needed

**Pull Request Template:**
```
## Changes
- Main changes description
- Added features
- Fixed issues

## Testing
- [ ] All existing tests pass
- [ ] New tests added
- [ ] Manual testing in Excel completed

## Screenshots/Examples
(if applicable)

Closes #123
```

## Project Structure

```
persian-calendar-excel/
├── src/                    # Main VBA files
├── examples/              # Usage examples
├── docs/                  # Documentation
├── tests/                 # Additional tests
└── tools/                 # Helper tools
```

## Types of Contributions

### High Priority
- Fix date conversion bugs
- Improve calculation accuracy
- Fix Excel compatibility issues

### Medium Priority
- Add new functions
- Improve performance
- Enhance documentation

### Low Priority
- Improve code style
- Add examples
- Translate documentation

## Non-Technical Contributions

- **Documentation**: Improve README, API docs
- **Translation**: Translate to other languages
- **Testing**: Test in different environments
- **Design**: Improve documentation appearance

## Questions?

- [GitHub Discussions](https://github.com/behify/persian-calendar-excel/discussions)
- [Issues](https://github.com/behify/persian-calendar-excel/issues)

## Code of Conduct

- Respect all contributors
- Provide constructive feedback
- Maintain friendly and professional environment

## Development Setup

1. Clone the repository
2. Open Excel and VBA Editor (Alt+F11)
3. Import all .bas files from src/ directory
4. Run tests to verify setup

## Release Process

1. Update CHANGELOG.md
2. Update version numbers
3. Run full test suite
4. Create release tag
5. Update documentation

Thank you for contributing!
# Contributing to MultiPageImporter

Thank you for your interest in contributing to MultiPageImporter.

## Reporting Issues

### Bug Reports

If the script is not working correctly:

1. Go to [Issues](https://github.com/101sam/ID-MultiPageImporter/issues)
2. Click **New Issue**
3. Select **Bug Report**
4. Provide:
   - InDesign version
   - Operating system
   - Script version (check top of .jsx file)
   - File type being imported (PDF, AI, INDD, etc.)
   - Steps to reproduce
   - Full error message
   - Screenshots if applicable

### Feature Requests

For new features or improvements:

1. Go to [Issues](https://github.com/101sam/ID-MultiPageImporter/issues)
2. Click **New Issue**
3. Select **Feature Request**
4. Describe the feature and use case

### Before Opening an Issue

- Check [existing issues](https://github.com/101sam/ID-MultiPageImporter/issues) for duplicates
- Check the [Known Issues section](README.md#known-issues--future-enhancements) in README

## Submitting Code Changes

### Fork and Clone

1. Fork this repository
2. Clone your fork:
   ```bash
   git clone https://github.com/YOUR-USERNAME/ID-MultiPageImporter.git
   ```

### Create a Branch

```bash
git checkout -b fix/description-of-fix
```

Branch naming:
- `fix/` - Bug fixes
- `feat/` - New features
- `docs/` - Documentation

### Make Changes

1. Make your changes to `MultiPageImporter.jsx`
2. Test thoroughly in InDesign
3. Update version number and changelog in both the script header and README.md

### Code Guidelines

- Do NOT use `exit()` - use `return` instead (causes crashes in InDesign 2024+)
- Maintain backward compatibility with CS3+
- Add comments for complex logic
- Test on both macOS and Windows if possible

### Submit Pull Request

1. Push your branch
2. Create a Pull Request
3. Describe your changes
4. Reference any related issues

## Testing

Before submitting:

- Test with various PDF files (different page counts, sizes)
- Test with AI files (single and multi-artboard)
- Test with INDD/INDT/IDML files
- Test creating new document vs. placing in existing
- Test different crop options
- Verify no `exit()` calls in code

## Questions

Open an issue with the **question** label if you need help.

---

Thank you for contributing.

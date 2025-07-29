# Security Policy

## Supported Versions

We provide security updates for the following versions:

| Version | Supported          |
| ------- | ------------------ |
| 1.0.x   | :white_check_mark: |

## Reporting a Vulnerability

If you discover a security vulnerability in this project, please report it responsibly:

### Private Reporting
- **Email**: Create an issue in the repository with the "security" label
- **Response Time**: We aim to respond within 48 hours
- **Disclosure**: We follow responsible disclosure practices

### What to Include
When reporting a security issue, please include:

1. **Description**: Clear description of the vulnerability
2. **Impact**: What could an attacker achieve?
3. **Reproduction**: Steps to reproduce the issue
4. **Environment**: Excel version, OS, and other relevant details
5. **Suggested Fix**: If you have ideas for fixes

### What Qualifies as a Security Issue

For this VBA Excel package, security issues might include:

- **Code Injection**: VBA code that could execute malicious commands
- **File System Access**: Unauthorized access to files outside the workbook
- **Memory Issues**: Buffer overflows or memory corruption
- **Data Leakage**: Unintended exposure of sensitive information

### What is NOT a Security Issue

- General bugs that don't have security implications
- Feature requests
- Performance issues
- Compatibility problems

### Process

1. **Report**: Submit your security report
2. **Acknowledgment**: We'll acknowledge receipt within 48 hours
3. **Investigation**: We'll investigate and validate the issue
4. **Fix**: We'll develop and test a fix
5. **Disclosure**: We'll coordinate disclosure timing with you
6. **Release**: We'll release the fix and credit you (if desired)

### Security Best Practices for Users

When using this package:

- **Enable Macro Security**: Use Excel's macro security settings
- **Trusted Sources**: Only enable macros from trusted sources
- **Regular Updates**: Keep Excel and this package updated
- **Code Review**: Review any modifications you make to the code
- **Backup**: Always backup your data before using new macros

### Contact

For security-related questions or concerns:
- Create an issue with the "security" label
- Mark the issue as confidential if needed

Thank you for helping keep this project secure!
# 🤝 Contributing to OPS Communicator

Thank you for considering contributing to this project! This guide will help you get started.

## Ways to Contribute

- 🐛 **Bug Reports**: Found an issue? Let us know!
- 💡 **Feature Requests**: Have an idea for improvement?
- 📖 **Documentation**: Help improve guides and examples
- 🔧 **Code Contributions**: Fix bugs or add features
- 🧪 **Testing**: Try the system and report your experience

## Getting Started

### Development Setup

1. **Fork the Repository**
   ```
   Click the "Fork" button on GitHub
   ```

2. **Create a Test Environment**
   - Create a new Google Sheet for testing
   - Follow the [installation guide](INSTALLATION.md)
   - Copy your modified code to test changes

3. **Test Your Changes**
   - Always test in a separate Google Sheet
   - Use the simulation and test features
   - Verify edge cases work correctly

### Making Changes

1. **Create a Feature Branch**
   ```bash
   git checkout -b feature/your-feature-name
   ```

2. **Make Your Changes**
   - Keep changes focused and atomic
   - Follow existing code style
   - Add comments for complex logic

3. **Test Thoroughly**
   - Test the specific feature you changed
   - Run end-to-end tests with different scenarios
   - Check error handling

4. **Update Documentation**
   - Update README.md if needed
   - Add installation notes for new features
   - Update function comments in code

## Code Style Guidelines

### JavaScript/Google Apps Script
- Use consistent indentation (2 spaces)
- Use descriptive variable and function names
- Add JSDoc comments for complex functions
- Handle errors gracefully with try/catch
- Use console.log for debugging (not alert)

### Example:
```javascript
/**
 * Parses time strings and Date objects into 24-hour format
 * @param {string|Date} timeStr - Time in various formats
 * @return {number} Hour in 24-hour format (0-23) or -1 if invalid
 */
function parseTimeString(timeStr) {
  if (!timeStr) return -1;
  
  // Handle Date objects
  if (timeStr instanceof Date) {
    return timeStr.getHours();
  }
  
  // Handle string formats...
}
```

### HTML/CSS
- Use semantic HTML elements
- Keep CSS organized and commented
- Ensure accessibility (labels, proper form elements)
- Test in different screen sizes

## Submitting Changes

### Pull Request Process

1. **Create a Pull Request**
   - Use a descriptive title
   - Reference any related issues
   - Describe what you changed and why

2. **PR Description Template**
   ```markdown
   ## What Changed
   Brief description of your changes
   
   ## Why
   Explain the problem this solves
   
   ## Testing
   How did you test these changes?
   
   ## Screenshots (if applicable)
   Show before/after if UI changes
   ```

3. **Review Process**
   - Maintainers will review your PR
   - Address any feedback promptly
   - Keep discussions respectful and constructive

## Bug Reports

### Before Submitting
- Check existing issues for duplicates
- Test in a clean environment
- Gather relevant information

### Bug Report Template
```markdown
**Describe the Bug**
Clear description of what happened

**To Reproduce**
Steps to reproduce the issue:
1. Go to '...'
2. Click on '...'
3. See error

**Expected Behavior**
What should have happened

**Screenshots**
Add screenshots if helpful

**Environment**
- Google Apps Script version
- Browser (if relevant)
- Any relevant settings

**Error Messages**
Copy any error messages from:
- Apps Script execution transcript
- Browser console
- Email bounces
```

## Feature Requests

### Suggesting Features
- Check existing issues and discussions
- Explain the use case clearly
- Consider how it fits with existing features
- Think about potential complications

### Feature Request Template
```markdown
**Problem Statement**
What problem does this solve?

**Proposed Solution**
Describe your ideal solution

**Alternatives Considered**
Other ways this could be addressed

**Additional Context**
Any other relevant information
```

## Code Review Guidelines

### For Contributors
- Be open to feedback
- Explain your reasoning for complex decisions
- Keep discussions focused on the code
- Update your PR based on feedback

### For Reviewers
- Be constructive and helpful
- Ask questions to understand the approach
- Suggest improvements, not just problems
- Appreciate the contributor's effort

## Development Tips

### Testing Your Changes
1. **Create Test Data**
   - Set up multiple communications
   - Test different time formats
   - Try various day combinations

2. **Test Edge Cases**
   - Empty fields
   - Invalid email addresses
   - Future dates
   - Different time zones

3. **Verify Existing Features**
   - Ensure you didn't break anything
   - Test the major user flows
   - Check error handling

### Google Apps Script Specifics
- Execution time limit: 6 minutes per function
- Daily trigger limit: 20 triggers per script
- Email quota limits apply
- Timezone handling can be tricky

## Community

- Be respectful and inclusive
- Help newcomers get started
- Share knowledge and experiences
- Focus on constructive discussions

## Questions?

- Open an issue for technical questions
- Use discussions for broader topics
- Check existing documentation first

Thank you for contributing! 🎉

# Testing Guide

## Running Tests

```bash
npm test
```

## Troubleshooting

### Tests failing with "ENOENT: no such file or directory"

If all tests are failing with file not found errors, try the following steps:

1. **Clean and reinstall dependencies:**
   ```bash
   rm -rf node_modules package-lock.json
   npm install
   ```

2. **Clear vitest cache:**
   ```bash
   npx vitest --clearCache
   ```

3. **Check Node.js version:**
   The tests have been tested with Node.js 16-20. If you're using Node.js v24 or newer, there may be compatibility issues. Check your version:
   ```bash
   node --version
   ```
   
   If you're on Node.js v24+, you may need to use an older version:
   ```bash
   nvm install 20
   nvm use 20
   ```

4. **Run tests with verbose output:**
   ```bash
   npm test -- --reporter=verbose
   ```

5. **Check for permission issues:**
   On macOS/Linux, ensure the temp directory has proper permissions:
   ```bash
   ls -la $(node -e "console.log(require('os').tmpdir())")
   ```

### Common Issues

- **Node.js v24+**: The project uses `@types/node` 16.6.2 which may not be fully compatible with Node.js v24. Consider using Node.js v20 LTS.
- **macOS/Linux**: If you get permission denied errors, check that the temp directory is writable
- **Windows**: Path separators might cause issues - ensure you're using the latest version of Node.js
- **CI vs Local**: The tests use the system's temp directory, which may behave differently locally vs in CI

### Debug Mode

The tests now include automatic diagnostic output when files are missing. When a test fails, you'll see:
- Complete directory tree of what was actually created
- Expected file paths that were not found
- Progress reports from the importer (successes and failures)

This information will help identify whether the issue is:
- Files being created in the wrong location
- Files not being created at all
- Import process failing silently

If you continue to have issues, please provide:
1. Your Node.js version (`node --version`)
2. Your operating system (macOS version, etc.)
3. The full test output showing the diagnostic information

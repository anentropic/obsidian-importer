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
   The tests require Node.js 16 or higher. Check your version:
   ```bash
   node --version
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

- **macOS/Linux**: If you get permission denied errors, check that the temp directory is writable
- **Windows**: Path separators might cause issues - ensure you're using the latest version of Node.js
- **CI vs Local**: The tests use the system's temp directory, which may behave differently locally vs in CI

### Debug Mode

To see detailed output about file creation during tests, look for console.log output that shows:
- Directory trees
- Expected vs actual file paths
- Progress reports from the importer

If you continue to have issues, please provide:
1. Your Node.js version (`node --version`)
2. Your operating system
3. The full test output with `--reporter=verbose`

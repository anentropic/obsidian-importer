# Testing Guide

## Prerequisites

- **Node.js**: Version 16-20 (Node.js v20 LTS recommended)
- **npm**: Comes with Node.js

### Installing Node.js

If you don't have Node.js installed or need to switch versions:

**Using nvm (recommended):**
```bash
# Install nvm (if not already installed)
curl -o- https://raw.githubusercontent.com/nvm-sh/nvm/v0.39.0/install.sh | bash

# Install Node.js 20 LTS
nvm install 20
nvm use 20

# Verify installation
node --version  # Should show v20.x.x
```

**Direct installation:**
- Download from [nodejs.org](https://nodejs.org/) (choose the LTS version)

## Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/anentropic/obsidian-importer.git
   cd obsidian-importer
   ```

2. **Install dependencies:**
   ```bash
   npm install
   ```

## Running Tests

### Run all tests:
```bash
npm test
```

### Run tests in watch mode (auto-rerun on file changes):
```bash
npx vitest
```

### Run tests with verbose output:
```bash
npm test -- --reporter=verbose
```

### Run a specific test file:
```bash
npx vitest tests/onenote-importer.test.ts
```

### Run tests with coverage:
```bash
npx vitest --coverage
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

- **Unsupported Node.js version**: The tests are compatible with Node.js 16-20. Node.js v24+ is not supported due to incompatibility with `@types/node` 16.6.2. Use Node.js v20 LTS (recommended).
- **macOS/Linux**: If you get permission denied errors, check that the temp directory is writable
- **Windows**: Path separators might cause issues - ensure you're using a supported Node.js version (16-20)
- **CI vs Local**: The tests use the system's temp directory, which may behave differently locally vs in CI

If you continue to have issues, please provide:
1. Your Node.js version (`node --version`)
2. Your operating system (macOS version, etc.)
3. The full test output

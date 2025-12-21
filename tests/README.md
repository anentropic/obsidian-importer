# Testing Guide

## Prerequisites

- **Node.js**: Version 16-20
  - (newer node versions don't work due to `"@types/node": "16.6.2"` dependency)
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

Tests are using [Vitest](https://vitest.dev/).

### Run all tests:
```bash
npm test
```

### Run a specific test file:
```bash
npm test tests/onenote-importer.test.ts
```

/**
 * Polyfills to match Obsidian's runtime helpers
 * 
 * These extensions add Obsidian-specific methods to native prototypes
 * that are available in the Obsidian runtime but not in the test environment.
 * 
 * Use setupObsidianPolyfills() and teardownObsidianPolyfills() in beforeEach/afterEach
 * or beforeAll/afterAll hooks to properly manage the polyfill lifecycle.
 */

interface PolyfillState {
	arrayContains?: typeof Array.prototype.includes;
	stringContains?: typeof String.prototype.includes;
	htmlElementFindAll?: (selector: string) => Element[];
}

const originalMethods: PolyfillState = {};

export function setupObsidianPolyfills(): void {
	// Save original methods if they exist
	if ((Array.prototype as any).contains) {
		originalMethods.arrayContains = (Array.prototype as any).contains;
	}
	if ((String.prototype as any).contains) {
		originalMethods.stringContains = (String.prototype as any).contains;
	}
	if ((HTMLElement.prototype as any).findAll) {
		originalMethods.htmlElementFindAll = (HTMLElement.prototype as any).findAll;
	}

	// Add polyfills
	if (!(Array.prototype as any).contains) {
		// eslint-disable-next-line no-extend-native
		(Array.prototype as any).contains = function(value: any) {
			return this.includes(value);
		};
	}

	if (!(String.prototype as any).contains) {
		// eslint-disable-next-line no-extend-native
		(String.prototype as any).contains = function(value: string) {
			return this.includes(value);
		};
	}

	if (!(HTMLElement.prototype as any).findAll) {
		// eslint-disable-next-line no-extend-native
		(HTMLElement.prototype as any).findAll = function(selector: string) {
			return Array.from(this.querySelectorAll(selector));
		};
	}
}

export function teardownObsidianPolyfills(): void {
	// Restore or delete polyfills
	if (originalMethods.arrayContains !== undefined) {
		(Array.prototype as any).contains = originalMethods.arrayContains;
	} else {
		delete (Array.prototype as any).contains;
	}

	if (originalMethods.stringContains !== undefined) {
		(String.prototype as any).contains = originalMethods.stringContains;
	} else {
		delete (String.prototype as any).contains;
	}

	if (originalMethods.htmlElementFindAll !== undefined) {
		(HTMLElement.prototype as any).findAll = originalMethods.htmlElementFindAll;
	} else {
		delete (HTMLElement.prototype as any).findAll;
	}

	// Clear the saved state
	originalMethods.arrayContains = undefined;
	originalMethods.stringContains = undefined;
	originalMethods.htmlElementFindAll = undefined;
}


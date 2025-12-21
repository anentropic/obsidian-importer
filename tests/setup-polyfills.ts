/**
 * Polyfills to match Obsidian's runtime helpers
 * 
 * These extensions add Obsidian-specific methods to native prototypes
 * that are available in the Obsidian runtime but not in the test environment.
 */

export function setupObsidianPolyfills(): void {
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

// Automatically setup polyfills when this file is imported as a setup file
setupObsidianPolyfills();

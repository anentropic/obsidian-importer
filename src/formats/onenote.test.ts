import { describe, it, expect, beforeEach } from 'vitest';
import { Notebook, OnenoteSection, OnenotePage, SectionGroup } from '@microsoft/microsoft-graph-types';

/**
 * Test for OneNote importer page path resolution bug.
 * 
 * Bug description: When importing from OneNote, the first page of each section
 * is correctly placed in the section folder, but subsequent pages are placed
 * at the vault root instead.
 * 
 * Expected structure:
 * [vault]
 * ├── [notebook]
 * │   ├── [section 1]
 * │   │   ├── [first page of section 1]
 * │   │   ├── [second page of section 1]
 * │   ├── [section 2]
 * │   │   ├── [first page of section 2]
 * │   │   ├── [second page of section 2]
 * 
 * Actual (buggy) structure:
 * [vault]
 * ├── [notebook]
 * │   ├── [section 1]
 * │   │   ├── [first page of section 1]
 * │   ├── [section 2]
 * │   │   ├── [first page of section 2]
 * └── [rest of pages...]
 * 
 * ANALYSIS:
 * The path resolution logic (getEntityPathNoParent, getEntityPath, searchPages)
 * has been thoroughly tested and appears correct. All tests pass, indicating
 * the logic correctly determines the path for all pages (first and subsequent).
 * 
 * HYPOTHESIS: The bug is likely in the Obsidian vault operations in processFile():
 * 
 * ```typescript
 * if (!await this.vault.adapter.exists(outputPath)) 
 *     pageFolder = await this.vault.createFolder(outputPath);
 * else 
 *     pageFolder = this.vault.getAbstractFileByPath(outputPath) as TFolder;
 * ```
 * 
 * For the FIRST page: vault.createFolder() is called and returns the folder.
 * For SUBSEQUENT pages: vault.getAbstractFileByPath() is called.
 * 
 * If getAbstractFileByPath() returns null (due to timing/indexing issues),
 * then pageFolder becomes null, and saveAsMarkdownFile(null, ...) might
 * save the file to the vault root.
 * 
 * This cannot be easily tested without the actual Obsidian environment,
 * but the path resolution tests below confirm the issue is NOT in the
 * path determination logic.
 */

// Extract the path resolution logic from OneNoteImporter for testing
// This mirrors the logic in onenote.ts but without Obsidian dependencies

class OneNotePathResolver {
	notebooks: Notebook[] = [];

	insertPagesToSection(pages: OnenotePage[], sectionId: string, parentEntity?: Notebook | SectionGroup) {
		if (!parentEntity) {
			for (const notebook of this.notebooks) {
				this.insertPagesToSection(pages, sectionId, notebook);
			}
			return;
		}

		if (parentEntity.sectionGroups) {
			const sectionGroups: SectionGroup[] = parentEntity.sectionGroups;
			for (const sectionGroup of sectionGroups) {
				this.insertPagesToSection(pages, sectionId, sectionGroup);
			}
		}

		if (parentEntity.sections) {
			const sectionGroup = parentEntity;
			for (const section of sectionGroup.sections!) {
				if (section.id === sectionId) {
					section.pages = pages;
				}
			}
		}
	}

	getEntityPathNoParent(entityID: string, currentPath: string): string | null {
		for (const notebook of this.notebooks) {
			const path = this.getEntityPath(entityID, `${currentPath}/${notebook.displayName}`, notebook);
			if (path) return path;
		}
		return null;
	}

	getEntityPath(entityID: string, currentPath: string, parentEntity: Notebook | SectionGroup | OnenoteSection): string | null {
		let returnPath: string | null = null;

		if ('sectionGroups' in parentEntity && parentEntity.sectionGroups) {
			const path = this.searchSectionGroups(entityID, currentPath, parentEntity.sectionGroups);
			if (path !== null) returnPath = path;
		}

		if ('sections' in parentEntity && parentEntity.sections) {
			const path = this.searchSectionGroups(entityID, currentPath, parentEntity.sections);
			if (path !== null) returnPath = path;
		}

		if ('pages' in parentEntity && parentEntity.pages) {
			const path = this.searchPages(entityID, currentPath, parentEntity as OnenoteSection);
			if (path !== null) returnPath = path;
		}

		return returnPath;
	}

	private searchPages(entityID: string, currentPath: string, section: OnenoteSection): string | null {
		let returnPath: string | null = null;
		for (let i = 0; i < section.pages!.length; i++) {
			const page = section.pages![i];
			const pageContentID = page.contentUrl!.split('page-id=')[1]?.split('}')[0];

			if (page.id === entityID || pageContentID === entityID) {
				if (page.level === 0) {
					if (section.pages![i + 1] && section.pages![i + 1].level !== 0) {
						returnPath = `${currentPath}/${page.title}`;
					}
					else returnPath = currentPath;
				}
				else {
					returnPath = currentPath;

					for (let i = section.pages!.indexOf(page) - 1; i >= 0; i--) {
						if (section.pages![i].level === page.level! - 1) {
							returnPath += '/' + section.pages![i].title;
							break;
						}
					}
				}
				break;
			}
		}
		return returnPath;
	}

	private searchSectionGroups(entityID: string, currentPath: string, sectionGroups: SectionGroup[] | OnenoteSection[]): string | null {
		let returnPath: string | null = null;
		for (const sectionGroup of sectionGroups) {
			if (sectionGroup.id === entityID) returnPath = `${currentPath}/${sectionGroup.displayName}`;
			else {
				const foundPath = this.getEntityPath(entityID, `${currentPath}/${sectionGroup.displayName}`, sectionGroup);
				if (foundPath) {
					returnPath = foundPath;
					break;
				}
			}
		}
		return returnPath;
	}
}

describe('OneNote Page Path Resolution Bug', () => {
	let resolver: OneNotePathResolver;

	beforeEach(() => {
		resolver = new OneNotePathResolver();
	});

	describe('when importing pages from a notebook with multiple sections', () => {
		const section1Pages: OnenotePage[] = [
			{ id: 'page-1-1', title: 'First Page Section 1', level: 0, contentUrl: 'https://example.com/page-id={page-1-1}' },
			{ id: 'page-1-2', title: 'Second Page Section 1', level: 0, contentUrl: 'https://example.com/page-id={page-1-2}' },
			{ id: 'page-1-3', title: 'Third Page Section 1', level: 0, contentUrl: 'https://example.com/page-id={page-1-3}' },
		];

		const section2Pages: OnenotePage[] = [
			{ id: 'page-2-1', title: 'First Page Section 2', level: 0, contentUrl: 'https://example.com/page-id={page-2-1}' },
			{ id: 'page-2-2', title: 'Second Page Section 2', level: 0, contentUrl: 'https://example.com/page-id={page-2-2}' },
		];

		beforeEach(() => {
			// Set up notebook structure: Notebook > Section 1, Section 2
			const notebook: Notebook = {
				id: 'notebook-1',
				displayName: 'My Notebook',
				sections: [
					{ id: 'section-1', displayName: 'Section 1' },
					{ id: 'section-2', displayName: 'Section 2' },
				],
			};

			resolver.notebooks = [notebook];

			// Simulate what happens during import:
			// 1. Pages for section 1 are fetched and inserted
			resolver.insertPagesToSection(section1Pages, 'section-1');
			// 2. Pages for section 2 are fetched and inserted
			resolver.insertPagesToSection(section2Pages, 'section-2');
		});

		it('should resolve the first page of section 1 to the correct path', () => {
			const path = resolver.getEntityPathNoParent('page-1-1', 'vault');
			expect(path).toBe('vault/My Notebook/Section 1');
		});

		it('should resolve the second page of section 1 to the correct path', () => {
			const path = resolver.getEntityPathNoParent('page-1-2', 'vault');
			expect(path).toBe('vault/My Notebook/Section 1');
		});

		it('should resolve the third page of section 1 to the correct path', () => {
			const path = resolver.getEntityPathNoParent('page-1-3', 'vault');
			expect(path).toBe('vault/My Notebook/Section 1');
		});

		it('should resolve the first page of section 2 to the correct path', () => {
			const path = resolver.getEntityPathNoParent('page-2-1', 'vault');
			expect(path).toBe('vault/My Notebook/Section 2');
		});

		it('should resolve the second page of section 2 to the correct path', () => {
			const path = resolver.getEntityPathNoParent('page-2-2', 'vault');
			expect(path).toBe('vault/My Notebook/Section 2');
		});

		it('should resolve all pages in section 1 to paths within that section', () => {
			const paths = section1Pages.map(page => resolver.getEntityPathNoParent(page.id!, 'vault'));
			
			// All pages should be in section 1, not at vault root
			for (const path of paths) {
				expect(path).not.toBeNull();
				expect(path).toContain('Section 1');
				expect(path).not.toBe('vault'); // Should NOT be at vault root
			}
		});

		it('should resolve all pages in section 2 to paths within that section', () => {
			const paths = section2Pages.map(page => resolver.getEntityPathNoParent(page.id!, 'vault'));
			
			// All pages should be in section 2, not at vault root
			for (const path of paths) {
				expect(path).not.toBeNull();
				expect(path).toContain('Section 2');
				expect(path).not.toBe('vault'); // Should NOT be at vault root
			}
		});
	});

	describe('when processing sections sequentially (mimicking actual import flow)', () => {
		it('should maintain correct paths even when sections are processed in order', () => {
			// This test mimics the actual import flow more closely
			const notebook: Notebook = {
				id: 'notebook-1',
				displayName: 'My Notebook',
				sections: [
					{ id: 'section-1', displayName: 'Section 1' },
					{ id: 'section-2', displayName: 'Section 2' },
				],
			};

			resolver.notebooks = [notebook];

			// Section 1 pages
			const section1Pages: OnenotePage[] = [
				{ id: 'page-1-1', title: 'Page 1-1', level: 0, contentUrl: 'https://example.com/page-id={page-1-1}' },
				{ id: 'page-1-2', title: 'Page 1-2', level: 0, contentUrl: 'https://example.com/page-id={page-1-2}' },
				{ id: 'page-1-3', title: 'Page 1-3', level: 0, contentUrl: 'https://example.com/page-id={page-1-3}' },
			];

			// Process section 1
			resolver.insertPagesToSection(section1Pages, 'section-1');

			// Check paths for section 1 pages BEFORE processing section 2
			expect(resolver.getEntityPathNoParent('page-1-1', 'vault')).toBe('vault/My Notebook/Section 1');
			expect(resolver.getEntityPathNoParent('page-1-2', 'vault')).toBe('vault/My Notebook/Section 1');
			expect(resolver.getEntityPathNoParent('page-1-3', 'vault')).toBe('vault/My Notebook/Section 1');

			// Section 2 pages
			const section2Pages: OnenotePage[] = [
				{ id: 'page-2-1', title: 'Page 2-1', level: 0, contentUrl: 'https://example.com/page-id={page-2-1}' },
				{ id: 'page-2-2', title: 'Page 2-2', level: 0, contentUrl: 'https://example.com/page-id={page-2-2}' },
			];

			// Process section 2
			resolver.insertPagesToSection(section2Pages, 'section-2');

			// Check paths for ALL pages AFTER processing both sections
			// This is where the bug would manifest - section 1 pages might lose their paths
			expect(resolver.getEntityPathNoParent('page-1-1', 'vault')).toBe('vault/My Notebook/Section 1');
			expect(resolver.getEntityPathNoParent('page-1-2', 'vault')).toBe('vault/My Notebook/Section 1');
			expect(resolver.getEntityPathNoParent('page-1-3', 'vault')).toBe('vault/My Notebook/Section 1');
			expect(resolver.getEntityPathNoParent('page-2-1', 'vault')).toBe('vault/My Notebook/Section 2');
			expect(resolver.getEntityPathNoParent('page-2-2', 'vault')).toBe('vault/My Notebook/Section 2');
		});
	});

	describe('edge cases that could cause pages to be misplaced', () => {
		it('should return correct path when page ID contains special characters', () => {
			const notebook: Notebook = {
				id: 'notebook-1',
				displayName: 'My Notebook',
				sections: [
					{ id: 'section-1', displayName: 'Section 1' },
				],
			};

			const resolver = new OneNotePathResolver();
			resolver.notebooks = [notebook];

			// Real OneNote page IDs look like GUIDs
			const pages: OnenotePage[] = [
				{ 
					id: '1-abc123-def456-789', 
					title: 'Page 1', 
					level: 0, 
					contentUrl: 'https://graph.microsoft.com/v1.0/users/me/onenote/pages/1-abc123-def456-789/content?page-id={1-abc123-def456-789}' 
				},
				{ 
					id: '2-xyz789-uvw012-345', 
					title: 'Page 2', 
					level: 0, 
					contentUrl: 'https://graph.microsoft.com/v1.0/users/me/onenote/pages/2-xyz789-uvw012-345/content?page-id={2-xyz789-uvw012-345}' 
				},
			];

			resolver.insertPagesToSection(pages, 'section-1');

			expect(resolver.getEntityPathNoParent('1-abc123-def456-789', 'vault')).toBe('vault/My Notebook/Section 1');
			expect(resolver.getEntityPathNoParent('2-xyz789-uvw012-345', 'vault')).toBe('vault/My Notebook/Section 1');
		});

		it('should handle multiple notebooks correctly', () => {
			const resolver = new OneNotePathResolver();
			
			// Set up multiple notebooks (as the user reported having)
			resolver.notebooks = [
				{
					id: 'notebook-1',
					displayName: 'Notebook 1',
					sections: [
						{ id: 'section-1-1', displayName: 'Section A' },
					],
				},
				{
					id: 'notebook-2',
					displayName: 'Notebook 2',
					sections: [
						{ id: 'section-2-1', displayName: 'Section B' },
					],
				},
			];

			// User only imports from Notebook 1
			const pagesFromNotebook1: OnenotePage[] = [
				{ id: 'page-1', title: 'First Page', level: 0, contentUrl: 'https://example.com/page-id={page-1}' },
				{ id: 'page-2', title: 'Second Page', level: 0, contentUrl: 'https://example.com/page-id={page-2}' },
				{ id: 'page-3', title: 'Third Page', level: 0, contentUrl: 'https://example.com/page-id={page-3}' },
			];

			resolver.insertPagesToSection(pagesFromNotebook1, 'section-1-1');

			// All pages should be found in Notebook 1/Section A
			expect(resolver.getEntityPathNoParent('page-1', 'OneNote')).toBe('OneNote/Notebook 1/Section A');
			expect(resolver.getEntityPathNoParent('page-2', 'OneNote')).toBe('OneNote/Notebook 1/Section A');
			expect(resolver.getEntityPathNoParent('page-3', 'OneNote')).toBe('OneNote/Notebook 1/Section A');
		});

		it('should NOT return null for any page in a section with multiple pages', () => {
			const resolver = new OneNotePathResolver();
			
			resolver.notebooks = [{
				id: 'notebook-1',
				displayName: 'My Notebook',
				sections: [
					{ id: 'section-1', displayName: 'My Section' },
				],
			}];

			const pages: OnenotePage[] = [
				{ id: 'page-1', title: 'First', level: 0, contentUrl: 'https://example.com/page-id={page-1}' },
				{ id: 'page-2', title: 'Second', level: 0, contentUrl: 'https://example.com/page-id={page-2}' },
				{ id: 'page-3', title: 'Third', level: 0, contentUrl: 'https://example.com/page-id={page-3}' },
				{ id: 'page-4', title: 'Fourth', level: 0, contentUrl: 'https://example.com/page-id={page-4}' },
				{ id: 'page-5', title: 'Fifth', level: 0, contentUrl: 'https://example.com/page-id={page-5}' },
			];

			resolver.insertPagesToSection(pages, 'section-1');

			// CRITICAL: Every single page should have a valid path, not null
			for (const page of pages) {
				const path = resolver.getEntityPathNoParent(page.id!, 'vault');
				expect(path).not.toBeNull();
				expect(path).toBe('vault/My Notebook/My Section');
			}
		});

		it('should handle empty currentPath (vault root)', () => {
			const resolver = new OneNotePathResolver();
			
			resolver.notebooks = [{
				id: 'notebook-1',
				displayName: 'My Notebook',
				sections: [
					{ id: 'section-1', displayName: 'My Section' },
				],
			}];

			const pages: OnenotePage[] = [
				{ id: 'page-1', title: 'First', level: 0, contentUrl: 'https://example.com/page-id={page-1}' },
				{ id: 'page-2', title: 'Second', level: 0, contentUrl: 'https://example.com/page-id={page-2}' },
			];

			resolver.insertPagesToSection(pages, 'section-1');

			// Even with empty currentPath, pages should be found
			const path1 = resolver.getEntityPathNoParent('page-1', '');
			const path2 = resolver.getEntityPathNoParent('page-2', '');

			expect(path1).not.toBeNull();
			expect(path2).not.toBeNull();
			expect(path1).toContain('My Notebook');
			expect(path2).toContain('My Notebook');
		});

		it('should return consistent paths for all pages in the same section', () => {
			// This test verifies that all pages in a section get the EXACT same path
			// If the bug is in path resolution, the first page would get a different path
			// than subsequent pages
			const resolver = new OneNotePathResolver();
			
			resolver.notebooks = [{
				id: 'notebook-1',
				displayName: 'My Notebook',
				sections: [
					{ id: 'section-1', displayName: 'Section 1' },
					{ id: 'section-2', displayName: 'Section 2' },
				],
			}];

			// Pages with all level 0 (no sub-pages)
			const section1Pages: OnenotePage[] = [
				{ id: 'p1', title: 'Page 1', level: 0, contentUrl: 'https://example.com/page-id={p1}' },
				{ id: 'p2', title: 'Page 2', level: 0, contentUrl: 'https://example.com/page-id={p2}' },
				{ id: 'p3', title: 'Page 3', level: 0, contentUrl: 'https://example.com/page-id={p3}' },
				{ id: 'p4', title: 'Page 4', level: 0, contentUrl: 'https://example.com/page-id={p4}' },
				{ id: 'p5', title: 'Page 5', level: 0, contentUrl: 'https://example.com/page-id={p5}' },
			];

			resolver.insertPagesToSection(section1Pages, 'section-1');

			// Get paths for all pages
			const paths = section1Pages.map(p => resolver.getEntityPathNoParent(p.id!, 'vault'));

			// All paths should be identical (all pages go to same section folder)
			const expectedPath = 'vault/My Notebook/Section 1';
			for (let i = 0; i < paths.length; i++) {
				expect(paths[i]).toBe(expectedPath);
			}
		});

		it('should handle pages where level is undefined or null', () => {
			// The Microsoft Graph API might not always return level for all pages
			// This tests what happens when level is missing
			const resolver = new OneNotePathResolver();
			
			resolver.notebooks = [{
				id: 'notebook-1',
				displayName: 'My Notebook',
				sections: [
					{ id: 'section-1', displayName: 'Section 1' },
				],
			}];

			// Pages where level might be undefined (simulating API response quirks)
			const pages: OnenotePage[] = [
				{ id: 'p1', title: 'Page 1', level: 0, contentUrl: 'https://example.com/page-id={p1}' },
				{ id: 'p2', title: 'Page 2', level: undefined, contentUrl: 'https://example.com/page-id={p2}' },
				{ id: 'p3', title: 'Page 3', level: null, contentUrl: 'https://example.com/page-id={p3}' },
			];

			resolver.insertPagesToSection(pages, 'section-1');

			// All pages should still resolve to the section path, not fail or go elsewhere
			const path1 = resolver.getEntityPathNoParent('p1', 'vault');
			const path2 = resolver.getEntityPathNoParent('p2', 'vault');
			const path3 = resolver.getEntityPathNoParent('p3', 'vault');

			expect(path1).not.toBeNull();
			expect(path2).not.toBeNull();
			expect(path3).not.toBeNull();

			// They should all be in the same section
			expect(path1).toContain('Section 1');
			expect(path2).toContain('Section 1');
			expect(path3).toContain('Section 1');
		});
	});
});

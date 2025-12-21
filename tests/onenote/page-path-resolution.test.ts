import { describe, it, expect, beforeEach, vi } from 'vitest';
import type { OnenotePage, Notebook, OnenoteSection, SectionGroup } from '@microsoft/microsoft-graph-types';

/**
 * Test for the bug where pages after the first one in each section are placed
 * at the vault root instead of in their section folders.
 * 
 * The bug scenario:
 * - User imports a notebook with multiple sections
 * - Each section has multiple pages
 * - First page of each section is correctly placed in section folder
 * - Rest of pages end up at vault root
 * 
 * Expected structure:
 *   vault/
 *   ├─ notebook/
 *   │  ├─ section 1/
 *   │  │  ├─ page 1 of section 1
 *   │  │  ├─ page 2 of section 1
 *   │  ├─ section 2/
 *   │  │  ├─ page 1 of section 2
 *   │  │  ├─ page 2 of section 2
 * 
 * Actual (buggy) structure:
 *   vault/
 *   ├─ notebook/
 *   │  ├─ section 1/
 *   │  │  ├─ page 1 of section 1
 *   │  ├─ section 2/
 *   │  │  ├─ page 1 of section 2
 *   ├─ page 2 of section 1  (wrong!)
 *   ├─ page 2 of section 2  (wrong!)
 * 
 * ROOT CAUSE HYPOTHESIS:
 * The issue is NOT in the path resolution logic (tests below confirm it works correctly).
 * The bug is likely in processFile() where:
 * 1. For the first page, the folder doesn't exist, so vault.createFolder() is called
 *    which creates the folder AND returns the TFolder object.
 * 2. For subsequent pages, the folder exists on disk (checked via vault.adapter.exists()),
 *    but vault.getAbstractFileByPath() may return null if Obsidian's internal cache
 *    hasn't been updated yet.
 * 3. When getAbstractFileByPath returns null, the pageFolder becomes null (cast as TFolder),
 *    and saveAsMarkdownFile(null, ...) creates the file at the vault root.
 */

// Helper function to create mock pages
function createMockPage(id: string, title: string, level: number = 0): OnenotePage {
	return {
		id,
		title,
		level,
		contentUrl: `https://graph.microsoft.com/v1.0/me/onenote/pages/${id}/content?page-id={${id}}`,
		createdDateTime: '2024-01-01T00:00:00Z',
		lastModifiedDateTime: '2024-01-01T00:00:00Z',
	};
}

// Helper function to create mock sections
function createMockSection(id: string, displayName: string, pages?: OnenotePage[]): OnenoteSection {
	const section: OnenoteSection = {
		id,
		displayName,
	};
	if (pages) {
		section.pages = pages;
	}
	return section;
}

// Helper function to create mock notebook
function createMockNotebook(id: string, displayName: string, sections: OnenoteSection[]): Notebook {
	return {
		id,
		displayName,
		sections,
	};
}

/**
 * Extracts and tests the path resolution logic from OneNoteImporter.
 * 
 * This replicates the getEntityPath, getEntityPathNoParent, searchSectionGroups,
 * and searchPages methods to test them in isolation.
 */
class PathResolver {
	notebooks: Notebook[] = [];
	
	sanitizeFilePath(path: string): string {
		return path.replace(/[:|?<>*\\]/g, '');
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

		if (returnPath) {
			returnPath = this.sanitizeFilePath(returnPath);
		}

		return returnPath;
	}

	private searchPages(entityID: string, currentPath: string, section: OnenoteSection): string | null {
		let returnPath: string | null = null;
		// Check if the target page is in the current entity's pages
		for (let i = 0; i < section.pages!.length; i++) {
			const page = section.pages![i];
			const pageContentID = page.contentUrl!.split('page-id=')[1]?.split('}')[0];

			if (page.id === entityID || pageContentID === entityID) {
				if (page.level === 0) {
					/* Checks if we have a page leveled below this one.
					 * without this line, leveled notes are more scattered:
					 * ...Section/Example.md, *but* ...Section/Example/Lower level.md
					 * with this line both files are in one neat directory:
					 * ...Section/Example/Page.md and ...Section/Example/Lower level.md
					 */
					if (section.pages![i + 1] && section.pages![i + 1].level !== 0) {
						returnPath = `${currentPath}/${page.title}`;
					}
					else returnPath = currentPath;
				}
				else {
					returnPath = currentPath;

					// Iterate backward to find the parent page
					for (let j = section.pages!.indexOf(page) - 1; j >= 0; j--) {
						if (section.pages![j].level === page.level! - 1) {
							returnPath += '/' + section.pages![j].title;
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
		// Recursively search in section groups
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

	/**
	 * Simulates insertPagesToSection - associates pages with their section
	 */
	insertPagesToSection(pages: OnenotePage[], sectionId: string, parentEntity?: Notebook | SectionGroup) {
		if (!parentEntity) {
			for (const notebook of this.notebooks) {
				this.insertPagesToSection(pages, sectionId, notebook);
			}
			return;
		}

		if (parentEntity.sectionGroups) {
			// Recursively search in section groups
			const sectionGroups: SectionGroup[] = parentEntity.sectionGroups;
			for (const sectionGroup of sectionGroups) {
				this.insertPagesToSection(pages, sectionId, sectionGroup);
			}
		}

		if (parentEntity.sections) {
			// Recursively search in sections
			const sectionGroup = parentEntity;
			for (const section of sectionGroup.sections!) {
				if (section.id === sectionId) {
					section.pages = pages;
				}
			}
		}
	}
}

describe('OneNote Page Path Resolution', () => {
	let resolver: PathResolver;
	
	beforeEach(() => {
		resolver = new PathResolver();
	});
	
	describe('Bug reproduction: pages placed at root instead of section', () => {
		it('should place all pages from a section in that section folder', () => {
			// Setup: notebook with 2 sections, each with 3 pages
			const section1Pages = [
				createMockPage('page-1-1', 'First Page Section 1'),
				createMockPage('page-1-2', 'Second Page Section 1'),
				createMockPage('page-1-3', 'Third Page Section 1'),
			];
			
			const section2Pages = [
				createMockPage('page-2-1', 'First Page Section 2'),
				createMockPage('page-2-2', 'Second Page Section 2'),
				createMockPage('page-2-3', 'Third Page Section 2'),
			];
			
			const section1 = createMockSection('section-1', 'Section 1');
			const section2 = createMockSection('section-2', 'Section 2');
			
			const notebook = createMockNotebook('notebook-1', 'My Notebook', [section1, section2]);
			
			resolver.notebooks = [notebook];
			
			// Simulate the import process for section 1
			resolver.insertPagesToSection(section1Pages, 'section-1');
			
			// Verify that pages were inserted into section 1
			expect(section1.pages).toBeDefined();
			expect(section1.pages!.length).toBe(3);
			
			// Now test path resolution for all pages in section 1
			const basePath = 'vault';
			
			// First page should be in section folder
			const path1 = resolver.getEntityPathNoParent('page-1-1', basePath);
			expect(path1).toBe('vault/My Notebook/Section 1');
			
			// Second page should ALSO be in section folder
			const path2 = resolver.getEntityPathNoParent('page-1-2', basePath);
			expect(path2).toBe('vault/My Notebook/Section 1');
			
			// Third page should ALSO be in section folder
			const path3 = resolver.getEntityPathNoParent('page-1-3', basePath);
			expect(path3).toBe('vault/My Notebook/Section 1');
		});
		
		it('simulates the full import flow for one section at a time', () => {
			// This test simulates the actual import flow more closely:
			// 1. Notebook with sections is loaded (sections don't have pages yet)
			// 2. For each selected section:
			//    a. Pages are fetched
			//    b. insertPagesToSection is called
			//    c. Each page is processed (getEntityPathNoParent is called)
			
			const section1 = createMockSection('section-1', 'Section 1');
			const section2 = createMockSection('section-2', 'Section 2');
			const notebook = createMockNotebook('notebook-1', 'My Notebook', [section1, section2]);
			
			resolver.notebooks = [notebook];
			
			// Simulate processing section 1
			const section1Pages = [
				createMockPage('page-1-1', 'First Page Section 1'),
				createMockPage('page-1-2', 'Second Page Section 1'),
				createMockPage('page-1-3', 'Third Page Section 1'),
			];
			
			// Step 2b: Insert pages into section 1
			resolver.insertPagesToSection(section1Pages, 'section-1');
			
			// Step 2c: Process each page - get path for each
			const paths1: (string | null)[] = [];
			for (const page of section1Pages) {
				const path = resolver.getEntityPathNoParent(page.id!, 'vault');
				paths1.push(path);
			}
			
			// All pages from section 1 should have the same section path
			expect(paths1[0]).toBe('vault/My Notebook/Section 1');
			expect(paths1[1]).toBe('vault/My Notebook/Section 1');
			expect(paths1[2]).toBe('vault/My Notebook/Section 1');
			
			// Now simulate processing section 2
			const section2Pages = [
				createMockPage('page-2-1', 'First Page Section 2'),
				createMockPage('page-2-2', 'Second Page Section 2'),
			];
			
			resolver.insertPagesToSection(section2Pages, 'section-2');
			
			const paths2: (string | null)[] = [];
			for (const page of section2Pages) {
				const path = resolver.getEntityPathNoParent(page.id!, 'vault');
				paths2.push(path);
			}
			
			// All pages from section 2 should have their section path
			expect(paths2[0]).toBe('vault/My Notebook/Section 2');
			expect(paths2[1]).toBe('vault/My Notebook/Section 2');
			
			// Verify section 1 pages are still accessible
			expect(resolver.getEntityPathNoParent('page-1-1', 'vault')).toBe('vault/My Notebook/Section 1');
			expect(resolver.getEntityPathNoParent('page-1-2', 'vault')).toBe('vault/My Notebook/Section 1');
		});
		
		it('should handle multiple sections independently', () => {
			// Setup: notebook with 2 sections, each with 2 pages
			const section1Pages = [
				createMockPage('page-1-1', 'Page 1 of Section 1'),
				createMockPage('page-1-2', 'Page 2 of Section 1'),
			];
			
			const section2Pages = [
				createMockPage('page-2-1', 'Page 1 of Section 2'),
				createMockPage('page-2-2', 'Page 2 of Section 2'),
			];
			
			const section1 = createMockSection('section-1', 'Section 1');
			const section2 = createMockSection('section-2', 'Section 2');
			
			const notebook = createMockNotebook('notebook-1', 'My Notebook', [section1, section2]);
			
			resolver.notebooks = [notebook];
			
			// Simulate importing section 1 first
			resolver.insertPagesToSection(section1Pages, 'section-1');
			
			// Check section 1 pages
			expect(resolver.getEntityPathNoParent('page-1-1', 'vault')).toBe('vault/My Notebook/Section 1');
			expect(resolver.getEntityPathNoParent('page-1-2', 'vault')).toBe('vault/My Notebook/Section 1');
			
			// Now simulate importing section 2
			resolver.insertPagesToSection(section2Pages, 'section-2');
			
			// Check section 2 pages
			expect(resolver.getEntityPathNoParent('page-2-1', 'vault')).toBe('vault/My Notebook/Section 2');
			expect(resolver.getEntityPathNoParent('page-2-2', 'vault')).toBe('vault/My Notebook/Section 2');
			
			// Verify section 1 pages are still correctly resolved
			expect(resolver.getEntityPathNoParent('page-1-1', 'vault')).toBe('vault/My Notebook/Section 1');
			expect(resolver.getEntityPathNoParent('page-1-2', 'vault')).toBe('vault/My Notebook/Section 1');
		});
		
		it('should return null for pages not found in any section', () => {
			const section1 = createMockSection('section-1', 'Section 1');
			const notebook = createMockNotebook('notebook-1', 'My Notebook', [section1]);
			
			resolver.notebooks = [notebook];
			
			// Don't insert any pages
			
			// Try to find a non-existent page
			const path = resolver.getEntityPathNoParent('non-existent-page', 'vault');
			expect(path).toBeNull();
		});
		
		it('should correctly handle searching when pages property does not exist on section', () => {
			// This simulates the case where sections are fetched from API without pages property
			const section1 = createMockSection('section-1', 'Section 1');
			const section2 = createMockSection('section-2', 'Section 2');
			
			const notebook = createMockNotebook('notebook-1', 'My Notebook', [section1, section2]);
			
			resolver.notebooks = [notebook];
			
			// Before inserting pages, neither section has 'pages' property
			expect('pages' in section1).toBe(false);
			expect('pages' in section2).toBe(false);
			
			// Insert pages only into section 1
			const section1Pages = [
				createMockPage('page-1-1', 'First Page'),
				createMockPage('page-1-2', 'Second Page'),
			];
			resolver.insertPagesToSection(section1Pages, 'section-1');
			
			// Now section 1 has pages, section 2 doesn't
			expect('pages' in section1).toBe(true);
			expect('pages' in section2).toBe(false);
			
			// Search should find pages in section 1
			expect(resolver.getEntityPathNoParent('page-1-1', 'vault')).toBe('vault/My Notebook/Section 1');
			expect(resolver.getEntityPathNoParent('page-1-2', 'vault')).toBe('vault/My Notebook/Section 1');
		});
	});
	
	describe('Page levels (subpages)', () => {
		it('should handle pages with subpages correctly', () => {
			// Page with level 1 subpage
			const pages = [
				createMockPage('page-1', 'Parent Page', 0),
				createMockPage('page-2', 'Child Page', 1),
			];
			
			const section = createMockSection('section-1', 'Section 1');
			const notebook = createMockNotebook('notebook-1', 'My Notebook', [section]);
			
			resolver.notebooks = [notebook];
			resolver.insertPagesToSection(pages, 'section-1');
			
			// Parent page should be in its own folder (because it has a subpage)
			expect(resolver.getEntityPathNoParent('page-1', 'vault')).toBe('vault/My Notebook/Section 1/Parent Page');
			
			// Child page should be in parent's folder
			expect(resolver.getEntityPathNoParent('page-2', 'vault')).toBe('vault/My Notebook/Section 1/Parent Page');
		});
	});
});

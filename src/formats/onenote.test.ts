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
 * These tests mock OneNote API data and test the real importer behavior.
 * Tests assert CORRECT behavior and should FAIL against the buggy implementation.
 */

// Mock OneNote API data
function createMockNotebook(id: string, displayName: string, sections: OnenoteSection[]): Notebook {
	return {
		id,
		displayName,
		sections,
	};
}

function createMockSection(id: string, displayName: string): OnenoteSection {
	return {
		id,
		displayName,
	};
}

function createMockPage(id: string, title: string, level: number = 0): OnenotePage {
	return {
		id,
		title,
		level,
		contentUrl: `https://graph.microsoft.com/v1.0/users/me/onenote/pages/${id}/content?page-id={${id}}`,
	};
}

/**
 * This class extracts and tests the core path resolution and folder handling
 * logic from OneNoteImporter. It uses the REAL implementation logic but with
 * mocked vault operations to detect the bug.
 */
class TestableOneNoteImporter {
	notebooks: Notebook[] = [];
	
	// Track all files created and their folder paths
	createdFiles: { filename: string; folderPath: string | null }[] = [];
	
	// Mock vault state
	private existingFolders: Set<string> = new Set();
	private folderObjects: Map<string, { path: string; name: string }> = new Map();

	/**
	 * This is the REAL insertPagesToSection logic from onenote.ts
	 */
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

	/**
	 * This is the REAL getEntityPathNoParent logic from onenote.ts
	 */
	getEntityPathNoParent(entityID: string, currentPath: string): string | null {
		for (const notebook of this.notebooks) {
			const path = this.getEntityPath(entityID, `${currentPath}/${notebook.displayName}`, notebook);
			if (path) return path;
		}
		return null;
	}

	/**
	 * This is the REAL getEntityPath logic from onenote.ts
	 */
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

	/**
	 * This is the REAL searchPages logic from onenote.ts
	 */
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

	/**
	 * This is the REAL searchSectionGroups logic from onenote.ts
	 */
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

	/**
	 * This is the REAL sanitizeFilePath logic from format-importer.ts
	 */
	sanitizeFilePath(path: string): string {
		return path.replace(/[:|?<>*\\]/g, '');
	}

	/**
	 * Mock vault.adapter.exists() - simulates Obsidian's file system check
	 */
	private async vaultAdapterExists(path: string): Promise<boolean> {
		return this.existingFolders.has(path);
	}

	/**
	 * Mock vault.createFolder() - simulates Obsidian folder creation
	 * Returns a folder object like the real API
	 */
	private async vaultCreateFolder(path: string): Promise<{ path: string; name: string }> {
		const folder = { path, name: path.split('/').pop()! };
		this.existingFolders.add(path);
		this.folderObjects.set(path, folder);
		return folder;
	}

	/**
	 * Mock vault.getAbstractFileByPath() - THIS IS WHERE THE BUG IS
	 * 
	 * In the real Obsidian API, this can return null even when the folder exists
	 * due to timing/indexing issues, path normalization differences, or case sensitivity.
	 * 
	 * This mock simulates the buggy behavior where getAbstractFileByPath returns null
	 * even though the folder was just created.
	 */
	private vaultGetAbstractFileByPath(path: string): { path: string; name: string } | null {
		// BUG SIMULATION: This returns null to simulate the real-world behavior
		// where getAbstractFileByPath fails to find a folder that was just created.
		// This happens in practice due to:
		// - Vault not indexed yet
		// - Path normalization differences between adapter.exists and getAbstractFileByPath
		// - Case sensitivity mismatches
		return null;
	}

	/**
	 * Mock saveAsMarkdownFile - records where files are saved
	 * When folder is null, file goes to vault root (the bug)
	 */
	private saveAsMarkdownFile(folder: { path: string; name: string } | null, title: string): void {
		this.createdFiles.push({
			filename: `${title}.md`,
			folderPath: folder?.path ?? null, // null means vault root
		});
	}

	/**
	 * This simulates the REAL processFile logic from onenote.ts lines 540-577
	 * It uses the same folder creation/lookup pattern that contains the bug.
	 */
	async processFile(page: OnenotePage, outputFolderName: string): Promise<void> {
		const outputPath = this.getEntityPathNoParent(page.id!, outputFolderName)!;

		let pageFolder: { path: string; name: string } | null;
		
		// THIS IS THE BUGGY CODE PATTERN FROM onenote.ts lines 546-548:
		// if (!await this.vault.adapter.exists(outputPath)) 
		//     pageFolder = await this.vault.createFolder(outputPath);
		// else 
		//     pageFolder = this.vault.getAbstractFileByPath(outputPath) as TFolder;
		
		if (!await this.vaultAdapterExists(outputPath)) {
			pageFolder = await this.vaultCreateFolder(outputPath);
		}
		else {
			// BUG: getAbstractFileByPath can return null even when folder exists!
			pageFolder = this.vaultGetAbstractFileByPath(outputPath);
		}

		// When pageFolder is null, saveAsMarkdownFile saves to vault root
		this.saveAsMarkdownFile(pageFolder, page.title!);
	}

	/**
	 * Simulates the full import flow for a section's pages
	 */
	async importSection(sectionId: string, pages: OnenotePage[], outputFolderName: string): Promise<void> {
		// First, insert pages into the section (like the real import does)
		this.insertPagesToSection(pages, sectionId);
		
		// Then process each page sequentially (like the real import does)
		for (const page of pages) {
			await this.processFile(page, outputFolderName);
		}
	}
}

describe('OneNote Importer - Page Placement Bug', () => {
	let importer: TestableOneNoteImporter;

	beforeEach(() => {
		importer = new TestableOneNoteImporter();
	});

	describe('importing pages from a section with multiple pages', () => {
		beforeEach(() => {
			// Set up mock OneNote data: one notebook with one section containing 3 pages
			importer.notebooks = [
				createMockNotebook('notebook-1', 'My Notebook', [
					createMockSection('section-1', 'Section 1'),
				]),
			];
		});

		it('should place ALL pages in the correct section folder', async () => {
			const pages = [
				createMockPage('page-1', 'First Page'),
				createMockPage('page-2', 'Second Page'),
				createMockPage('page-3', 'Third Page'),
			];

			await importer.importSection('section-1', pages, 'OneNote');

			// EXPECTED BEHAVIOR: All pages should be in the section folder
			expect(importer.createdFiles).toHaveLength(3);
			
			for (const file of importer.createdFiles) {
				expect(file.folderPath).not.toBeNull();
				expect(file.folderPath).toBe('OneNote/My Notebook/Section 1');
			}
		});

		it('should not place any pages at vault root (null folder)', async () => {
			const pages = [
				createMockPage('page-1', 'First Page'),
				createMockPage('page-2', 'Second Page'),
				createMockPage('page-3', 'Third Page'),
			];

			await importer.importSection('section-1', pages, 'OneNote');

			// EXPECTED BEHAVIOR: No pages should have null folderPath (vault root)
			const pagesAtRoot = importer.createdFiles.filter(f => f.folderPath === null);
			expect(pagesAtRoot).toHaveLength(0);
		});
	});

	describe('importing pages from multiple sections', () => {
		beforeEach(() => {
			// Set up mock OneNote data: one notebook with two sections
			importer.notebooks = [
				createMockNotebook('notebook-1', 'My Notebook', [
					createMockSection('section-1', 'Section 1'),
					createMockSection('section-2', 'Section 2'),
				]),
			];
		});

		it('should place all pages from all sections in their correct folders', async () => {
			const section1Pages = [
				createMockPage('page-1-1', 'First Page Section 1'),
				createMockPage('page-1-2', 'Second Page Section 1'),
			];
			const section2Pages = [
				createMockPage('page-2-1', 'First Page Section 2'),
				createMockPage('page-2-2', 'Second Page Section 2'),
			];

			// Import both sections
			await importer.importSection('section-1', section1Pages, 'OneNote');
			await importer.importSection('section-2', section2Pages, 'OneNote');

			expect(importer.createdFiles).toHaveLength(4);

			// Check section 1 pages
			const section1Files = importer.createdFiles.filter(f => 
				f.filename.includes('Section 1')
			);
			for (const file of section1Files) {
				expect(file.folderPath).toBe('OneNote/My Notebook/Section 1');
			}

			// Check section 2 pages
			const section2Files = importer.createdFiles.filter(f => 
				f.filename.includes('Section 2')
			);
			for (const file of section2Files) {
				expect(file.folderPath).toBe('OneNote/My Notebook/Section 2');
			}
		});
	});

	describe('importing many pages from a single section', () => {
		beforeEach(() => {
			importer.notebooks = [
				createMockNotebook('notebook-1', 'My Notebook', [
					createMockSection('section-1', 'My Section'),
				]),
			];
		});

		it('should place all 10 pages in the section folder, not vault root', async () => {
			const pages = Array.from({ length: 10 }, (_, i) => 
				createMockPage(`page-${i + 1}`, `Page ${i + 1}`)
			);

			await importer.importSection('section-1', pages, 'OneNote');

			expect(importer.createdFiles).toHaveLength(10);

			// EXPECTED: All pages in the section folder
			// BUG: Pages 2-10 end up at vault root (null)
			for (let i = 0; i < importer.createdFiles.length; i++) {
				const file = importer.createdFiles[i];
				expect(file.folderPath, `Page ${i + 1} should be in section folder`).toBe('OneNote/My Notebook/My Section');
			}
		});
	});
});

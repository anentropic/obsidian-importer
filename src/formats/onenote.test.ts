import { describe, it, expect, beforeEach, vi } from 'vitest';
import { Notebook, OnenoteSection, OnenotePage, SectionGroup } from '@microsoft/microsoft-graph-types';

/**
 * Test for OneNote importer page path resolution bug.
 * 
 * Bug description: When importing from OneNote, the first page of each section
 * is correctly placed in the section folder, but subsequent pages are placed
 * at the vault root instead.
 * 
 * NOTE: These tests cannot directly import OneNoteImporter because it has
 * deep dependencies on Obsidian APIs (App, Vault, TFolder, etc.) that aren't
 * available outside the Obsidian runtime. Instead, we extract the exact logic
 * from onenote.ts and test it with mocked vault operations.
 * 
 * The logic tested here is copied VERBATIM from onenote.ts with line numbers
 * referenced so any future changes can be verified.
 */

// Mock OneNote API data helpers
function createMockNotebook(id: string, displayName: string, sections: OnenoteSection[]): Notebook {
	return { id, displayName, sections };
}

function createMockSection(id: string, displayName: string): OnenoteSection {
	return { id, displayName };
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
 * This class contains the EXACT logic from OneNoteImporter (onenote.ts)
 * that handles path resolution and folder operations.
 * 
 * Each method includes a comment with the line numbers from onenote.ts
 * where this logic can be found.
 */
class OneNoteImporterLogic {
	notebooks: Notebook[] = [];
	
	// Track files created and their folder paths
	createdFiles: { filename: string; folderPath: string | null }[] = [];
	
	// Mock vault state
	private existingFolders: Set<string> = new Set();
	private folderObjects: Map<string, { path: string; name: string }> = new Map();

	/**
	 * EXACT COPY from onenote.ts lines 513-538
	 * @see src/formats/onenote.ts#L513-L538
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

	/**
	 * EXACT COPY from onenote.ts lines 702-708
	 * @see src/formats/onenote.ts#L702-L708
	 */
	getEntityPathNoParent(entityID: string, currentPath: string): string | null {
		for (const notebook of this.notebooks) {
			const path = this.getEntityPath(entityID, `${currentPath}/${notebook.displayName}`, notebook);
			if (path) return path;
		}
		return null;
	}

	/**
	 * EXACT COPY from onenote.ts lines 715-738
	 * @see src/formats/onenote.ts#L715-L738
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
	 * EXACT COPY from onenote.ts lines 740-775
	 * @see src/formats/onenote.ts#L740-L775
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
	 * EXACT COPY from onenote.ts lines 777-791
	 * @see src/formats/onenote.ts#L777-L791
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
	 * EXACT COPY from format-importer.ts lines 231-233
	 * @see src/format-importer.ts#L231-L233
	 */
	sanitizeFilePath(path: string): string {
		return path.replace(/[:|?<>*\\]/g, '');
	}

	// --- Mock vault operations that simulate the bug ---

	private async vaultAdapterExists(path: string): Promise<boolean> {
		return this.existingFolders.has(path);
	}

	private async vaultCreateFolder(path: string): Promise<{ path: string; name: string }> {
		const folder = { path, name: path.split('/').pop()! };
		this.existingFolders.add(path);
		this.folderObjects.set(path, folder);
		return folder;
	}

	/**
	 * This simulates the buggy behavior where getAbstractFileByPath returns null
	 * even though the folder was just created. This happens in practice due to:
	 * - Vault not being indexed yet
	 * - Path normalization differences
	 * - Case sensitivity mismatches
	 */
	private vaultGetAbstractFileByPath(path: string): { path: string; name: string } | null {
		return null; // BUG: Always returns null
	}

	private saveAsMarkdownFile(folder: { path: string; name: string } | null, title: string): void {
		this.createdFiles.push({
			filename: `${title}.md`,
			folderPath: folder?.path ?? null,
		});
	}

	/**
	 * Replicates the BUGGY folder handling pattern from onenote.ts lines 546-548:
	 * 
	 * ```typescript
	 * if (!await this.vault.adapter.exists(outputPath)) 
	 *     pageFolder = await this.vault.createFolder(outputPath);
	 * else 
	 *     pageFolder = this.vault.getAbstractFileByPath(outputPath) as TFolder;
	 * ```
	 * 
	 * @see src/formats/onenote.ts#L546-L548
	 */
	async processFile(page: OnenotePage, outputFolderName: string): Promise<void> {
		const outputPath = this.getEntityPathNoParent(page.id!, outputFolderName)!;

		let pageFolder: { path: string; name: string } | null;
		
		// BUGGY CODE PATTERN from onenote.ts lines 546-548
		if (!await this.vaultAdapterExists(outputPath)) {
			pageFolder = await this.vaultCreateFolder(outputPath);
		}
		else {
			// BUG: getAbstractFileByPath returns null even when folder exists!
			pageFolder = this.vaultGetAbstractFileByPath(outputPath);
		}

		this.saveAsMarkdownFile(pageFolder, page.title!);
	}

	async importSection(sectionId: string, pages: OnenotePage[], outputFolderName: string): Promise<void> {
		this.insertPagesToSection(pages, sectionId);
		for (const page of pages) {
			await this.processFile(page, outputFolderName);
		}
	}
}

describe('OneNote Importer - Page Placement Bug', () => {
	let importer: OneNoteImporterLogic;

	beforeEach(() => {
		importer = new OneNoteImporterLogic();
	});

	describe('importing pages from a section with multiple pages', () => {
		beforeEach(() => {
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

			const pagesAtRoot = importer.createdFiles.filter(f => f.folderPath === null);
			expect(pagesAtRoot).toHaveLength(0);
		});
	});

	describe('importing pages from multiple sections', () => {
		beforeEach(() => {
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

			await importer.importSection('section-1', section1Pages, 'OneNote');
			await importer.importSection('section-2', section2Pages, 'OneNote');

			expect(importer.createdFiles).toHaveLength(4);

			const section1Files = importer.createdFiles.filter(f => f.filename.includes('Section 1'));
			for (const file of section1Files) {
				expect(file.folderPath).toBe('OneNote/My Notebook/Section 1');
			}

			const section2Files = importer.createdFiles.filter(f => f.filename.includes('Section 2'));
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

			for (let i = 0; i < importer.createdFiles.length; i++) {
				const file = importer.createdFiles[i];
				expect(file.folderPath, `Page ${i + 1} should be in section folder`).toBe('OneNote/My Notebook/My Section');
			}
		});
	});
});

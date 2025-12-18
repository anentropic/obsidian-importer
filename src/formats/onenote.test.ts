import { describe, it, expect, beforeEach } from 'vitest';
import { Notebook, OnenoteSection, OnenotePage } from '@microsoft/microsoft-graph-types';
import { OneNotePathResolver } from './onenote/path-resolver';

/**
 * Test for OneNote importer page path resolution bug.
 * 
 * Bug description: When importing from OneNote, the first page of each section
 * is correctly placed in the section folder, but subsequent pages are placed
 * at the vault root instead.
 * 
 * This test uses the actual OneNotePathResolver implementation for path
 * resolution logic, and mocks only the vault operations that simulate the bug.
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
 * Test harness that uses the actual OneNotePathResolver for path resolution
 * while providing mock vault operations to test the page placement bug.
 */
class OneNoteImporterTestHarness {
	// Use the actual path resolver implementation
	pathResolver: OneNotePathResolver = new OneNotePathResolver();
	
	// Convenience getter/setter to access notebooks via pathResolver
	get notebooks(): Notebook[] {
		return this.pathResolver.notebooks;
	}
	set notebooks(value: Notebook[]) {
		this.pathResolver.notebooks = value;
	}
	
	// Track files created and their folder paths
	createdFiles: { filename: string, folderPath: string | null }[] = [];
	
	// Mock vault state
	private existingFolders: Set<string> = new Set();
	private folderObjects: Map<string, { path: string, name: string }> = new Map();

	// --- Mock vault operations that simulate the bug ---

	private async vaultAdapterExists(path: string): Promise<boolean> {
		return this.existingFolders.has(path);
	}

	private async vaultCreateFolder(path: string): Promise<{ path: string, name: string }> {
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
	private vaultGetAbstractFileByPath(path: string): { path: string, name: string } | null {
		return null; // BUG: Always returns null
	}

	private saveAsMarkdownFile(folder: { path: string, name: string } | null, title: string): void {
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
		const outputPath = this.pathResolver.getEntityPathNoParent(page.id!, outputFolderName)!;

		let pageFolder: { path: string, name: string } | null;
		
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
		this.pathResolver.insertPagesToSection(pages, sectionId);
		for (const page of pages) {
			await this.processFile(page, outputFolderName);
		}
	}
}

describe('OneNote Importer - Page Placement Bug', () => {
	let importer: OneNoteImporterTestHarness;

	beforeEach(() => {
		importer = new OneNoteImporterTestHarness();
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

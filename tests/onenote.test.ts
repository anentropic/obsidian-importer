import { describe, it, expect, beforeEach, vi } from 'vitest';
// Mocks must be declared before importing modules that depend on them.
vi.mock('obsidian', () => {
	class FakeSetting {
		setName() { return this; }
		setDesc() { return this; }
		addToggle() { return this; }
		addButton() { return this; }
		setCta() { return this; }
		setButtonText() { return this; }
		onClick() { return this; }
		addText() { return this; }
	}

	return {
		Notice: class {},
		Setting: FakeSetting,
		TFolder: class {},
		TFile: class {},
		Vault: class {},
		App: class {},
		Plugin: class {},
		Modal: class {},
		normalizePath: (p: string) => p,
		Platform: { isDesktopApp: false },
		moment: () => ({
			format: () => '',
		}),
		htmlToMarkdown: (html: unknown) => (typeof html === 'string' ? html : String(html)),
		requestUrl: async () => ({ json: async () => ({}) }),
	};
});

// Stub zip dependency used by other format importers pulled in transitively.
vi.mock('zip', () => ({
	readZip: async () => ({}),
	ZipEntryFile: class {},
}));

vi.mock('../src/util', async () => {
	const actual = await vi.importActual<typeof import('../src/util')>('../src/util');
	return {
		...actual,
		parseHTML: (html: string) => ({
			outerHTML: html,
			querySelectorAll: () => [],
		}) as unknown as HTMLElement,
	};
});

import { Notebook, OnenoteSection, OnenotePage } from '@microsoft/microsoft-graph-types';
import { OneNoteImporter } from '../src/formats/onenote';
import { OneNotePathResolver } from '../src/formats/onenote/path-resolver';

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

type MockFolder = { path: string, name: string };

/**
 * Build a lightweight OneNoteImporter instance that uses the real folder
 * resolution logic from processFile but swaps out heavy dependencies with
 * simple stubs suited for unit testing.
 */
function createTestImporter(
	createdFiles: { filename: string, folderPath: string | null }[],
	existingFolders: Set<string>,
): OneNoteImporter & { outputFolderName: string } {
	const importer = Object.create(OneNoteImporter.prototype) as OneNoteImporter & { outputFolderName: string };

	// Use the real path resolver / notebook accessors.
	importer.pathResolver = new OneNotePathResolver();
	importer.outputFolderName = 'OneNote';

	// Minimal vault mock: replicates the bug by returning null from getAbstractFileByPath
	// even when the folder exists.
	importer.vault = {
		adapter: {
			exists: async (path: string) => existingFolders.has(path),
		},
		createFolder: async (path: string) => {
			existingFolders.add(path);
			return { path, name: path.split('/').pop()! } as MockFolder;
		},
		getAbstractFileByPath: (_path: string) => null,
		append: async () => {},
	} as any;

	// Lightweight impls for the rest of processFile's collaborators.
	importer.convertFormat = (content: string) => ({ html: content }) as any;
	importer.getOutputFolder = async () => ({ name: importer.outputFolderName } as any);
	importer.convertTags = (element: any) => (typeof element?.outerHTML === 'string' ? element.outerHTML : String(element ?? ''));
	importer.getAllAttachments = async (_progress: unknown, pageHTML: string) => pageHTML as any;
	importer.combineCodeBlocksAsNecessary = () => {};
	importer.styledElementToHTML = () => {};
	importer.convertInternalLinks = () => {};
	importer.convertDrawings = () => {};
	importer.convertMathML = () => {};
	importer.removeExtraListItemParagraphs = () => {};
	importer.escapeTextNodes = () => {};
	importer.saveAsMarkdownFile = async (folder: MockFolder | null, title: string) => {
		createdFiles.push({
			filename: `${title}.md`,
			folderPath: folder?.path ?? null,
		});
		return { path: folder?.path ?? null } as any;
	};

	return importer;
}

/**
 * Test harness that uses the actual OneNotePathResolver for path resolution
 * while providing mock vault operations to test the page placement bug.
 */
class OneNoteImporterTestHarness {
	createdFiles: { filename: string, folderPath: string | null }[] = [];
	private existingFolders: Set<string> = new Set();
	private importer = createTestImporter(this.createdFiles, this.existingFolders);

	get notebooks(): Notebook[] {
		return this.importer.notebooks;
	}

	set notebooks(value: Notebook[]) {
		this.importer.notebooks = value;
	}

	async importSection(sectionId: string, pages: OnenotePage[], outputFolderName: string): Promise<void> {
		this.importer.outputFolderName = outputFolderName;
		this.importer.insertPagesToSection(pages, sectionId);
		for (const page of pages) {
			await this.importer.processFile({
				reportNoteSuccess: () => {},
				reportFailed: () => {},
			} as any, '<div></div>', page);
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

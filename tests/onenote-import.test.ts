import { strict as assert } from 'assert';
import { test } from 'node:test';
import * as Module from 'module';
import type { OnenotePage, Notebook } from '@microsoft/microsoft-graph-types';

const originalLoad = Module._load;
let obsidianExports: any;
let mainExports: any;
Module._load = function (request: string, parent: NodeModule | null, isMain: boolean) {
	if (request === '../main') {
		if (!mainExports) {
			mainExports = { ATTACHMENT_EXTS: [], AUTH_REDIRECT_URI: '', ImportContext: class { } };
		}
		return mainExports;
	}
	if (request === 'obsidian') {
		if (!obsidianExports) {
			class TFolder {
				path: string;
				name: string;
				constructor(path: string = '') {
					this.path = path;
					this.name = path.split('/').pop() || '';
				}
			}
			class Vault {
				createdPaths = new Set<string>();
				adapter = {
					exists: async (path: string) => this.createdPaths.has(path),
				};
				async createFolder(path: string) {
					this.createdPaths.add(path);
					return new TFolder(path);
				}
				getAbstractFileByPath(path: string) {
					return this.createdPaths.has(path) ? new TFolder(path) : null;
				}
				getAbstractFileByPathInsensitive(path: string) {
					return this.getAbstractFileByPath(path);
				}
			}
			obsidianExports = {
				App: class { },
				Vault,
				TFolder,
				Plugin: class { },
				Modal: class {
					contentEl = { createDiv: () => ({ createEl: () => ({ }), empty: () => { }, show: () => { }, hide: () => { } }) };
				},
				Notice: class { },
				Setting: class {
					settingEl = { show: () => { }, hide: () => { } };
					contentEl: any = {};
					setName() { return this; }
					setDesc() { return this; }
					addToggle() { return this; }
					addButton() { return this; }
					addText() { return this; }
				},
				htmlToMarkdown: () => '',
				requestUrl: async () => ({}),
				moment: (..._args: any[]) => ({ format: () => '', utc: () => ({ format: () => '' }) }),
				normalizePath: (p: string) => p,
				Platform: { isDesktopApp: false, isMacOS: false },
			};
		}
		return obsidianExports;
	}
	if (request === 'zip') {
		return {};
	}
	return originalLoad(request, parent, isMain);
};

// Import after stubbing dependencies
// eslint-disable-next-line @typescript-eslint/no-var-requires
const { OneNoteImporter } = require('../src/formats/onenote');
// eslint-disable-next-line @typescript-eslint/no-var-requires
const { TFolder } = require('obsidian');

class TestPlugin {
	data = { importers: { onenote: { previouslyImportedIDs: [] as string[] } } };
	async loadData() {
		return this.data;
	}
	async saveData(data: any) {
		this.data = data;
	}
	registerAuthCallback() { }
}

class TestModal {
	plugin = new TestPlugin();
	contentEl = { createDiv: () => ({ createEl: () => ({ }), empty: () => { }, show: () => { }, hide: () => { } }) };
	abortController = new AbortController();
}

class TestVault {
	createdFolders = new Set<string>();
	adapter = {
		exists: async (path: string) => this.createdFolders.has(path),
	};
	async createFolder(path: string) {
		this.createdFolders.add(path);
		return new TFolder(path);
	}
	getAbstractFileByPath(path: string) {
		return this.createdFolders.has(path) ? new TFolder(path) : null;
	}
	getAbstractFileByPathInsensitive(path: string) {
		return this.getAbstractFileByPath(path);
	}
	fileManager = {
		createNewMarkdownFile: async (_folder: any, _name: string) => ({ }),
	};
}

class TestApp {
	vault = new TestVault();
	fileManager = this.vault.fileManager;
}

class TestContext {
	status(_message: string) { }
	reportProgress(_current: number, _total: number) { }
	isCancelled() { return false; }
	reportSkipped(_title: string, _reason: string) { }
	reportFailed(_title: string, _error: any) { }
	reportNoteSuccess(_title: string) { }
}

class TestOneNoteImporter extends OneNoteImporter {
	pagesBySection: Record<string, OnenotePage[]> = {};
	recordedPaths: Record<string, string> = {};

	init() {
		// Skip UI setup
	}

	async fetchResource<T>(_url: string, _type: any, _progress?: any): Promise<any> {
		if (_url.includes('/sections/')) {
			const match = /sections\/([^/]+)\/pages/.exec(_url);
			if (match) {
				const value = (this.pagesBySection[match[1]] ?? []).map((p) => ({
					...p,
					createdDateTime: '2023-01-01T00:00:00Z',
					lastModifiedDateTime: '2023-01-02T00:00:00Z',
					self: `https://graph.microsoft.com/v1.0/me/onenote/pages/${p.id}`,
				}));
				return {
					'@odata.context': 'https://graph.microsoft.com/v1.0/$metadata#users(\'{id}\')/notes/sections(\'{id}\')/pages',
					value,
				};
			}
		}
		// Minimal HTML body for page content
		return '<html><body><p>content</p></body></html>' as unknown as T;
	}

	convertFormat(input: string) {
		return { html: input, inkml: '' };
	}

	async processFile(progress: any, _content: string, page: OnenotePage) {
		const outputFolder = await this.getOutputFolder();
		const outputPath = this.getEntityPathNoParent(page.id!, outputFolder!.name)!;

		let pageFolder: TFolder;
		if (!await this.vault.adapter.exists(outputPath)) {
			pageFolder = await this.vault.createFolder(outputPath);
		}
		else {
			pageFolder = this.vault.getAbstractFileByPath(outputPath) as TFolder;
		}

		this.recordedPaths[page.id!] = `${pageFolder.path}/${page.title}.md`;
		progress.reportNoteSuccess(page.title!);
	}
}

const notebooks: Notebook[] = [{
	id: 'nb1',
	displayName: 'Notebook',
	sections: [
		{ id: 's1', displayName: 'Section One' },
		{ id: 's2', displayName: 'Section Two' }
	]
}];

const pages: Record<string, OnenotePage[]> = {
	s1: [
		{ id: 'p1', title: 'Page 1', level: 0, contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/1/content?page-id={p1}' },
		{ id: 'p2', title: 'Page 2', level: 0, contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/2/content?page-id={p2}' }
	],
	s2: [
		{ id: 'p3', title: 'Page 3', level: 0, contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/3/content?page-id={p3}' },
		{ id: 'p4', title: 'Page 4', level: 0, contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/4/content?page-id={p4}' }
	],
};

test('imports all OneNote pages into their sections when importing a notebook', async () => {
	const app = new TestApp();
	const modal = new TestModal();
	const importer = new TestOneNoteImporter(app as any, modal as any);

	importer.graphData.accessToken = 'token';
	importer.outputLocation = 'OneNote';
	importer.notebooks = notebooks;
	importer.pagesBySection = pages;
	importer.selectedIds = ['s1', 's2'];

	const ctx = new TestContext();
	await importer.import(ctx as any);

	assert.deepStrictEqual(importer.recordedPaths, {
		p1: 'OneNote/Notebook/Section One/Page 1.md',
		p2: 'OneNote/Notebook/Section One/Page 2.md',
		p3: 'OneNote/Notebook/Section Two/Page 3.md',
		p4: 'OneNote/Notebook/Section Two/Page 4.md',
	});
});

test('keeps all pages from a section inside that section (no section groups)', async () => {
	const app = new TestApp();
	const modal = new TestModal();
	const importer = new TestOneNoteImporter(app as any, modal as any);

	importer.graphData.accessToken = 'token';
	importer.outputLocation = 'OneNote';
	importer.notebooks = notebooks;
	importer.pagesBySection = pages;
	importer.selectedIds = ['s1', 's2'];

	const ctx = new TestContext();
	await importer.import(ctx as any);

	assert.deepStrictEqual(importer.recordedPaths, {
		p1: 'OneNote/Notebook/Section One/Page 1.md',
		p2: 'OneNote/Notebook/Section One/Page 2.md',
		p3: 'OneNote/Notebook/Section Two/Page 3.md',
		p4: 'OneNote/Notebook/Section Two/Page 4.md',
	});
});

test('duplicate section IDs with missing display name can flatten pages to notebook root', async () => {
	const app = new TestApp();
	const modal = new TestModal();
	const importer = new TestOneNoteImporter(app as any, modal as any);

	importer.graphData.accessToken = 'token';
	importer.outputLocation = 'OneNote';
	// Two sections share the same ID; the first has no displayName, so searchSectionGroups
	// will match it first and build a path missing the section name.
	importer.notebooks = [{
		id: 'nb1',
		displayName: 'Notebook',
		sections: [
			{ id: 's-dup', displayName: '' as any },
			{ id: 's-dup', displayName: 'Actual Section' },
		],
	}];
	importer.pagesBySection = {
		's-dup': [
			{ id: 'pA', title: 'Page A', level: 0, contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/A/content?page-id={pA}' },
			{ id: 'pB', title: 'Page B', level: 0, contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/B/content?page-id={pB}' },
		],
	};
	importer.selectedIds = ['s-dup'];

	const ctx = new TestContext();
	await importer.import(ctx as any);

	assert.deepStrictEqual(importer.recordedPaths, {
		pA: 'OneNote/Notebook/Page A.md',
		pB: 'OneNote/Notebook/Page B.md',
	});
});

test('out-of-order section-page fetch still keeps pages in their sections', async () => {
	const app = new TestApp();
	const modal = new TestModal();
	const importer = new TestOneNoteImporter(app as any, modal as any);

	importer.graphData.accessToken = 'token';
	importer.outputLocation = 'OneNote';
	// Pages for s2 arrive first, then pages for s1
	importer.notebooks = [{
		id: 'nb1',
		displayName: 'Notebook',
		sections: [
			{ id: 's1', displayName: 'Section One' },
			{ id: 's2', displayName: 'Section Two' },
		],
	}];
	importer.pagesBySection = {
		s2: [
			{ id: 'p2a', title: 'Page 2A', level: 0, contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/2a/content?page-id={p2a}' },
			{ id: 'p2b', title: 'Page 2B', level: 0, contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/2b/content?page-id={p2b}' },
		],
		s1: [
			{ id: 'p1a', title: 'Page 1A', level: 0, contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/1a/content?page-id={p1a}' },
			{ id: 'p1b', title: 'Page 1B', level: 0, contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/1b/content?page-id={p1b}' },
		],
	};
	// Simulate fetch order: s2 first, then s1
	importer.selectedIds = ['s2', 's1'];

	const ctx = new TestContext();
	await importer.import(ctx as any);

	assert.deepStrictEqual(importer.recordedPaths, {
		p2a: 'OneNote/Notebook/Section Two/Page 2A.md',
		p2b: 'OneNote/Notebook/Section Two/Page 2B.md',
		p1a: 'OneNote/Notebook/Section One/Page 1A.md',
		p1b: 'OneNote/Notebook/Section One/Page 1B.md',
	});
});

test('pages flatten to vault root when output folder is empty name', async () => {
	const app = new TestApp();
	const modal = new TestModal();
	const importer = new TestOneNoteImporter(app as any, modal as any);

	// Simulate output folder resolving to root with an empty name
	importer.getOutputFolder = async () => new TFolder('');

	importer.graphData.accessToken = 'token';
	importer.outputLocation = '';
	importer.notebooks = notebooks;
	importer.pagesBySection = {
		s1: [
			{ id: 'p1', title: 'Page 1', level: 0, contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/1/content?page-id={p1}' },
			{ id: 'p2', title: 'Page 2', level: 0, contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/2/content?page-id={p2}' }
		],
		s2: [
			{ id: 'p3', title: 'Page 3', level: 0, contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/3/content?page-id={p3}' },
			{ id: 'p4', title: 'Page 4', level: 0, contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/4/content?page-id={p4}' }
		],
	};
	importer.selectedIds = ['s1', 's2'];

	const ctx = new TestContext();
	await importer.import(ctx as any);

	assert.deepStrictEqual(importer.recordedPaths, {
		p1: '/Page 1.md',
		p2: '/Page 2.md',
		p3: '/Page 3.md',
		p4: '/Page 4.md',
	});
});

test('all sections present, multiple pages per section — pages stay in their sections', async () => {
	const app = new TestApp();
	const modal = new TestModal();
	const importer = new TestOneNoteImporter(app as any, modal as any);

	importer.graphData.accessToken = 'token';
	importer.outputLocation = 'OneNote';
	importer.notebooks = [{
		id: 'nb1',
		displayName: 'Notebook',
		sections: [
			{ id: 's1', displayName: 'Section One' },
			{ id: 's2', displayName: 'Section Two' },
		],
	}];
	importer.pagesBySection = {
		s1: [
			{ id: 'p1a', title: 'Page 1A', level: 0, contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/1a/content?page-id={p1a}' },
			{ id: 'p1b', title: 'Page 1B', level: 0, contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/1b/content?page-id={p1b}' },
		],
		s2: [
			{ id: 'p2a', title: 'Page 2A', level: 0, contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/2a/content?page-id={p2a}' },
			{ id: 'p2b', title: 'Page 2B', level: 0, contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/2b/content?page-id={p2b}' },
		],
	};
	importer.selectedIds = ['s1', 's2'];

	const ctx = new TestContext();
	await importer.import(ctx as any);

	assert.deepStrictEqual(importer.recordedPaths, {
		p1a: 'OneNote/Notebook/Section One/Page 1A.md',
		p1b: 'OneNote/Notebook/Section One/Page 1B.md',
		p2a: 'OneNote/Notebook/Section Two/Page 2A.md',
		p2b: 'OneNote/Notebook/Section Two/Page 2B.md',
	});
});

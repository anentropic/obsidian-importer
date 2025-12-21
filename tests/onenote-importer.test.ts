import { describe, it, expect, vi, beforeEach, afterEach, beforeAll, afterAll } from 'vitest';
import path from 'node:path';
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import os from 'node:os';
import { Notice } from 'obsidian';
import type { ImporterData } from '../src/main';
import { OneNoteImporter } from '../src/formats/onenote';
import { setupObsidianPolyfills, teardownObsidianPolyfills } from './setup-polyfills';
import { MOCK_DATE_STRING } from './obsidian-mock';
import { createRequestMock } from './request-mock';


class TestableOneNoteImporter extends OneNoteImporter {
	init(): void {
		// Skip UI setup; the tests set state directly
		this.outputLocation = 'OneNote';
	}
}

interface TestHarness {
	importer: TestableOneNoteImporter;
	pluginData: ImporterData;
	root: string;
}

const createHarness = async (): Promise<TestHarness> => {
	const root = await fsp.mkdtemp(path.join(os.tmpdir(), 'onenote-importer-'));
	const { Vault, App } = await import('obsidian');
	// @ts-ignore mocked class
	const vault = new Vault(root);
	// @ts-ignore mocked class
	const fileManager = new (await import('obsidian')).FileManager(vault);
	const app = new App(vault, fileManager);
	const pluginData: ImporterData = { importers: { onenote: { previouslyImportedIDs: [] } } };
	const plugin = {
		loadData: vi.fn().mockResolvedValue(pluginData),
		saveData: vi.fn(async (data: ImporterData) => Object.assign(pluginData, data)),
		registerAuthCallback: vi.fn(),
	};
	const modal = { contentEl: document.createElement('div'), plugin, abortController: new AbortController() } as any;
	const importer = new TestableOneNoteImporter(app as any, modal);
	importer.graphData.accessToken = 'token';
	return { importer, pluginData, root };
};

const createProgress = () => {
	return {
		status: vi.fn(),
		reportProgress: vi.fn(),
		reportSkipped: vi.fn(),
		reportFailed: vi.fn(),
		reportNoteSuccess: vi.fn(),
		reportAttachmentSuccess: vi.fn(),
		isCancelled: () => false,
	};
};

const debugTestFailure = async (root: string, expectedPaths: string[], progress: any) => {
	const listDir = async (dir: string, prefix = '', depth = 0): Promise<string> => {
		if (depth > 10) return `${prefix}[max depth reached]\n`;
		try {
			const items = await fsp.readdir(dir, { withFileTypes: true });
			let result = '';
			for (const item of items) {
				result += `${prefix}${item.isDirectory() ? '[DIR] ' : '[FILE]'} ${item.name}\n`;
				if (item.isDirectory()) {
					result += await listDir(path.join(dir, item.name), prefix + '  ', depth + 1);
				}
			}
			return result;
		} catch (e) {
			return `${prefix}Error reading directory: ${e}\n`;
		}
	};

	const dirTree = await listDir(root);
	console.log('Directory tree:');
	console.log(dirTree);
	console.log('Expected paths:');
	expectedPaths.forEach((p, i) => console.log(`  [${i}]:`, p));
	console.log('Progress reports:');
	console.log('  reportNoteSuccess calls:', progress.reportNoteSuccess.mock.calls);
	console.log('  reportFailed calls:', progress.reportFailed.mock.calls);
};

const buildMultipartContent = (html: string): string => {
	const boundary = '--batch_12345';
	return [
		`${boundary}`,
		'Content-Type: text/html; charset=utf-8',
		'',
		html,
		`${boundary}`,
		'Content-Type: application/inkml+xml',
		'',
		'<inkml></inkml>',
		'',
	].join('\n');
};

describe('OneNoteImporter integration', () => {
	let tmpRoot: string;
	const requestMock = createRequestMock();

	beforeAll(() => {
		setupObsidianPolyfills();
	});

	afterAll(() => {
		teardownObsidianPolyfills();
	});

	beforeEach(() => {
		Notice.messages = [];
		requestMock.start();
	});

	afterEach(async () => {
		vi.restoreAllMocks();
		requestMock.stop();
		requestMock.clearHandlers();
		if (tmpRoot) {
			await fsp.rm(tmpRoot, { recursive: true, force: true });
		}
	});

	it('imports a simple page and writes markdown', async () => {
		const { importer, root } = await createHarness();
		tmpRoot = root;
		importer.selectedIds = ['section-1'];
		importer.notebooks = [
			{
				id: 'notebook-1',
				displayName: 'Test Notebook',
				sections: [
					{ id: 'section-1', displayName: 'Section A', pages: [] },
				],
			},
		] as any;

		const htmlBody = '<html><body><p>Hello OneNote</p></body></html>';
		const content = buildMultipartContent(htmlBody);

		const pagesResponse = {
			value: [
				{
					id: 'page-1',
					title: 'My Note',
					createdDateTime: '2023-01-01T00:00:00Z',
					lastModifiedDateTime: '2023-01-02T00:00:00Z',
					level: 0,
					order: 0,
					contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/page-1/content?includeInkML=true',
				},
			],
		};

		requestMock.addHandler(async (url: string | URL) => {
			const target = url.toString();
			if (target.includes('/sections/section-1/pages')) {
				return new Response(JSON.stringify(pagesResponse), { status: 200 });
			}
			if (target.includes('/pages/page-1/content')) {
				return new Response(content, { status: 200 });
			}
			return new Response('not found', { status: 404 });
		});

		const progress = createProgress();
		await importer.import(progress as any);

		const notePath = path.join(root, 'OneNote', 'Test Notebook', 'Section A', 'My Note.md');
		const md = await fsp.readFile(notePath, 'utf8');
		expect(md).toContain('Hello OneNote');
		expect(progress.reportNoteSuccess).toHaveBeenCalledWith('My Note');
	});

	it('downloads attachments and rewrites embeds', async () => {
		const { importer, root } = await createHarness();
		tmpRoot = root;
		importer.selectedIds = ['section-2'];
		importer.notebooks = [
			{
				id: 'notebook-2',
				displayName: 'Work Notebook',
				sections: [
					{ id: 'section-2', displayName: 'Attachments', pages: [] },
				],
			},
		] as any;

		const htmlBody = [
			'<html><body>',
			'<object data-attachment="report.pdf" data="https://files.example.com/report.pdf"></object>',
			'<img data-fullres-src="https://files.example.com/photo" data-fullres-src-type="image/png" alt="Found via OCR" />',
			'</body></html>',
		].join('');
		const content = buildMultipartContent(htmlBody);

		const pagesResponse = {
			value: [
				{
					id: 'page-2',
					title: 'Page With Attachments',
					createdDateTime: '2023-01-01T00:00:00Z',
					lastModifiedDateTime: '2023-01-02T00:00:00Z',
					level: 0,
					order: 0,
					contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/page-2/content?includeInkML=true',
				},
			],
		};

		const binaryBuffer = Buffer.from('file');
		requestMock.addHandler(async (url: string | URL) => {
			const target = url.toString();
			if (target.includes('/sections/section-2/pages')) {
				return new Response(JSON.stringify(pagesResponse), { status: 200 });
			}
			if (target.includes('/pages/page-2/content')) {
				return new Response(content, { status: 200 });
			}
			if (target.includes('report.pdf') || target.includes('photo')) {
				return new Response(binaryBuffer, { status: 200 });
			}
			return new Response('not found', { status: 404 });
		});

		const progress = createProgress();
		await importer.import(progress as any);

		const notePath = path.join(root, 'OneNote', 'Work Notebook', 'Attachments', 'Page With Attachments.md');
		const md = await fsp.readFile(notePath, 'utf8');
		const attachmentPath = path.join(root, 'OneNote', 'report.pdf');
		const imagePath = path.join(root, 'OneNote', `Exported image ${MOCK_DATE_STRING}-0.png`);

		expect(fs.existsSync(attachmentPath)).toBe(true);
		expect(fs.existsSync(imagePath)).toBe(true);
		expect(md).toContain('![Found via OCR]');
		expect(progress.reportAttachmentSuccess).toHaveBeenCalledTimes(2);
	});

	it('imports notebook with multiple sections', async () => {
		const { importer, root } = await createHarness();
		tmpRoot = root;
		importer.selectedIds = ['section-a', 'section-b'];
		importer.notebooks = [
			{
				id: 'notebook-3',
				displayName: 'Multi-Section Notebook',
				sections: [
					{ id: 'section-a', displayName: 'First Section', pages: [] },
					{ id: 'section-b', displayName: 'Second Section', pages: [] },
				],
			},
		] as any;

		const htmlBodyA = '<html><body><p>Content from first section</p></body></html>';
		const contentA = buildMultipartContent(htmlBodyA);
		const htmlBodyB = '<html><body><p>Content from second section</p></body></html>';
		const contentB = buildMultipartContent(htmlBodyB);

		const pagesResponseA = {
			value: [
				{
					id: 'page-a1',
					title: 'Page A1',
					createdDateTime: '2023-01-01T00:00:00Z',
					lastModifiedDateTime: '2023-01-02T00:00:00Z',
					level: 0,
					order: 0,
					contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/page-a1/content?includeInkML=true',
				},
				{
					id: 'page-a2',
					title: 'Page A2',
					createdDateTime: '2023-01-01T00:00:00Z',
					lastModifiedDateTime: '2023-01-02T00:00:00Z',
					level: 0,
					order: 1,
					contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/page-a2/content?includeInkML=true',
				},
			],
		};

		const pagesResponseB = {
			value: [
				{
					id: 'page-b1',
					title: 'Page B1',
					createdDateTime: '2023-01-01T00:00:00Z',
					lastModifiedDateTime: '2023-01-02T00:00:00Z',
					level: 0,
					order: 0,
					contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/page-b1/content?includeInkML=true',
				},
			],
		};

		requestMock.addHandler(async (url: string | URL) => {
			const target = url.toString();
			if (target.includes('/sections/section-a/pages')) {
				return new Response(JSON.stringify(pagesResponseA), { status: 200 });
			}
			if (target.includes('/sections/section-b/pages')) {
				return new Response(JSON.stringify(pagesResponseB), { status: 200 });
			}
			if (target.includes('/pages/page-a1/content') || target.includes('/pages/page-a2/content')) {
				return new Response(contentA, { status: 200 });
			}
			if (target.includes('/pages/page-b1/content')) {
				return new Response(contentB, { status: 200 });
			}
			return new Response('not found', { status: 404 });
		});

		const progress = createProgress();
		await importer.import(progress as any);

		// Verify pages from first section
		const notePathA1 = path.join(root, 'OneNote', 'Multi-Section Notebook', 'First Section', 'Page A1.md');
		const notePathA2 = path.join(root, 'OneNote', 'Multi-Section Notebook', 'First Section', 'Page A2.md');
		const mdA1 = await fsp.readFile(notePathA1, 'utf8');
		const mdA2 = await fsp.readFile(notePathA2, 'utf8');
		expect(mdA1).toContain('Content from first section');
		expect(mdA2).toContain('Content from first section');

		// Verify pages from second section
		const notePathB1 = path.join(root, 'OneNote', 'Multi-Section Notebook', 'Second Section', 'Page B1.md');
		const mdB1 = await fsp.readFile(notePathB1, 'utf8');
		expect(mdB1).toContain('Content from second section');

		expect(progress.reportNoteSuccess).toHaveBeenCalledWith('Page A1');
		expect(progress.reportNoteSuccess).toHaveBeenCalledWith('Page A2');
		expect(progress.reportNoteSuccess).toHaveBeenCalledWith('Page B1');
	});

	it('imports notebook with section group', async () => {
		const { importer, root } = await createHarness();
		tmpRoot = root;
		importer.selectedIds = ['section-sg1', 'section-sg2'];
		importer.notebooks = [
			{
				id: 'notebook-4',
				displayName: 'Notebook With Groups',
				sectionGroups: [
					{
						id: 'group-1',
						displayName: 'My Section Group',
						sections: [
							{ id: 'section-sg1', displayName: 'Grouped Section 1', pages: [] },
							{ id: 'section-sg2', displayName: 'Grouped Section 2', pages: [] },
						],
					},
				],
			},
		] as any;

		const htmlBodySG1 = '<html><body><p>Content from grouped section 1</p></body></html>';
		const contentSG1 = buildMultipartContent(htmlBodySG1);
		const htmlBodySG2 = '<html><body><p>Content from grouped section 2</p></body></html>';
		const contentSG2 = buildMultipartContent(htmlBodySG2);

		const pagesResponseSG1 = {
			value: [
				{
					id: 'page-sg1',
					title: 'Page SG1',
					createdDateTime: '2023-01-01T00:00:00Z',
					lastModifiedDateTime: '2023-01-02T00:00:00Z',
					level: 0,
					order: 0,
					contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/page-sg1/content?includeInkML=true',
				},
			],
		};

		const pagesResponseSG2 = {
			value: [
				{
					id: 'page-sg2',
					title: 'Page SG2',
					createdDateTime: '2023-01-01T00:00:00Z',
					lastModifiedDateTime: '2023-01-02T00:00:00Z',
					level: 0,
					order: 0,
					contentUrl: 'https://graph.microsoft.com/v1.0/me/onenote/pages/page-sg2/content?includeInkML=true',
				},
			],
		};

		requestMock.addHandler(async (url: string | URL) => {
			const target = url.toString();
			if (target.includes('/sections/section-sg1/pages')) {
				return new Response(JSON.stringify(pagesResponseSG1), { status: 200 });
			}
			if (target.includes('/sections/section-sg2/pages')) {
				return new Response(JSON.stringify(pagesResponseSG2), { status: 200 });
			}
			if (target.includes('/pages/page-sg1/content')) {
				return new Response(contentSG1, { status: 200 });
			}
			if (target.includes('/pages/page-sg2/content')) {
				return new Response(contentSG2, { status: 200 });
			}
			return new Response('not found', { status: 404 });
		});

		const progress = createProgress();
		await importer.import(progress as any);

		// Verify pages from section group
		const notePathSG1 = path.join(root, 'OneNote', 'Notebook With Groups', 'My Section Group', 'Grouped Section 1', 'Page SG1.md');
		const notePathSG2 = path.join(root, 'OneNote', 'Notebook With Groups', 'My Section Group', 'Grouped Section 2', 'Page SG2.md');
		
		// If files don't exist, print diagnostics
		if (!fs.existsSync(notePathSG1) || !fs.existsSync(notePathSG2)) {
			await debugTestFailure(root, [notePathSG1, notePathSG2], progress);
		}
		
		const mdSG1 = await fsp.readFile(notePathSG1, 'utf8');
		const mdSG2 = await fsp.readFile(notePathSG2, 'utf8');
		expect(mdSG1).toContain('Content from grouped section 1');
		expect(mdSG2).toContain('Content from grouped section 2');

		expect(progress.reportNoteSuccess).toHaveBeenCalledWith('Page SG1');
		expect(progress.reportNoteSuccess).toHaveBeenCalledWith('Page SG2');
	});
});

export class Notice {}

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

export const Setting = FakeSetting;
export class TFolder {}
export const moment = () => ({ format: () => '' });
export const htmlToMarkdown = (html: unknown) => (typeof html === 'string' ? html : String(html));
export const requestUrl = async () => ({ json: async () => ({}) });

// Minimal types to satisfy tests
export type DataWriteOptions = { ctime?: number; mtime?: number };

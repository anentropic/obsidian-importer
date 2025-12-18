export async function resolveFolderWithCache<TFolderType>(
	cache: Map<string, TFolderType>,
	createFolder: (path: string) => Promise<TFolderType>,
	outputPath: string
): Promise<TFolderType> {
	if (!cache.has(outputPath)) {
		cache.set(outputPath, await createFolder(outputPath));
	}
	const folder = cache.get(outputPath);
	if (!folder) {
		throw new Error(`Failed to resolve folder at "${outputPath}"`);
	}
	return folder;
}

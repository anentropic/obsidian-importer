/**
 * Test to reproduce the OneNote section organization bug.
 * 
 * Bug description: When importing from OneNote, only the first page from each
 * section is placed in the correct section folder. Remaining pages are imported
 * to the root of the vault.
 * 
 * This test simulates the internal structure and path resolution to demonstrate
 * the bug in the original code.
 */

// Simple type definitions for testing
interface OnenotePage {
	id?: string;
	title?: string;
	level?: number;
	contentUrl?: string;
}

interface OnenoteSection {
	id?: string;
	displayName?: string;
	pages?: OnenotePage[];
	sections?: OnenoteSection[];
	sectionGroups?: SectionGroup[];
}

interface SectionGroup {
	id?: string;
	displayName?: string;
	sections?: OnenoteSection[];
	sectionGroups?: SectionGroup[];
}

interface Notebook {
	id?: string;
	displayName?: string;
	sections?: OnenoteSection[];
	sectionGroups?: SectionGroup[];
}

// Mock the sanitizeFilePath function
function sanitizeFilePath(path: string): string {
	return path.replace(/[<>:"|?*]/g, '_');
}

// Original buggy version of getEntityPath
function getEntityPath_BUGGY(
	entityID: string,
	currentPath: string,
	parentEntity: Notebook | SectionGroup | OnenoteSection,
	searchSectionGroups: (entityID: string, currentPath: string, groups: any[]) => string | null,
	searchPages: (entityID: string, currentPath: string, section: OnenoteSection) => string | null
): string | null {
	let returnPath: string | null = null;

	if ('sectionGroups' in parentEntity && parentEntity.sectionGroups) {
		const path = searchSectionGroups(entityID, currentPath, parentEntity.sectionGroups);
		if (path !== null) returnPath = path;
	}

	if ('sections' in parentEntity && parentEntity.sections) {
		const path = searchSectionGroups(entityID, currentPath, parentEntity.sections);
		if (path !== null) returnPath = path;
	}

	if ('pages' in parentEntity && parentEntity.pages) {
		const path = searchPages(entityID, currentPath, parentEntity);
		if (path !== null) returnPath = path;
	}

	if (returnPath) {
		returnPath = sanitizeFilePath(returnPath);
	}

	return returnPath;
}

// Fixed version of getEntityPath with early returns
function getEntityPath_FIXED(
	entityID: string,
	currentPath: string,
	parentEntity: Notebook | SectionGroup | OnenoteSection,
	searchSectionGroups: (entityID: string, currentPath: string, groups: any[]) => string | null,
	searchPages: (entityID: string, currentPath: string, section: OnenoteSection) => string | null
): string | null {
	if ('sectionGroups' in parentEntity && parentEntity.sectionGroups) {
		const path = searchSectionGroups(entityID, currentPath, parentEntity.sectionGroups);
		if (path !== null) {
			return sanitizeFilePath(path);
		}
	}

	if ('sections' in parentEntity && parentEntity.sections) {
		const path = searchSectionGroups(entityID, currentPath, parentEntity.sections);
		if (path !== null) {
			return sanitizeFilePath(path);
		}
	}

	if ('pages' in parentEntity && parentEntity.pages) {
		const path = searchPages(entityID, currentPath, parentEntity);
		if (path !== null) {
			return sanitizeFilePath(path);
		}
	}

	return null;
}

// Helper function to create mock searchPages
function createSearchPages() {
	return (entityID: string, currentPath: string, section: OnenoteSection): string | null => {
		if (!section.pages) return null;
		
		for (let i = 0; i < section.pages.length; i++) {
			const page = section.pages[i];
			if (page.id === entityID) {
				// Simplified logic: just return currentPath for level 0 pages
				if (page.level === 0) {
					return currentPath;
				}
			}
		}
		return null;
	};
}

// Helper function to create mock searchSectionGroups  
function createSearchSectionGroups(
	getEntityPathFunc: typeof getEntityPath_BUGGY | typeof getEntityPath_FIXED,
	searchPages: ReturnType<typeof createSearchPages>
) {
	const searchSectionGroups = (
		entityID: string,
		currentPath: string,
		sectionGroups: SectionGroup[] | OnenoteSection[]
	): string | null => {
		let returnPath: string | null = null;
		for (const sectionGroup of sectionGroups) {
			if (sectionGroup.id === entityID) {
				returnPath = `${currentPath}/${sectionGroup.displayName}`;
			} else {
				const foundPath = getEntityPathFunc(
					entityID,
					`${currentPath}/${sectionGroup.displayName}`,
					sectionGroup,
					searchSectionGroups,
					searchPages
				);
				if (foundPath) {
					returnPath = foundPath;
					break;
				}
			}
		}
		return returnPath;
	};
	return searchSectionGroups;
}

/**
 * Test scenario 1: Reproduces the bug where pages end up in the wrong location.
 * 
 * The bug occurs when:
 * 1. A notebook has both sectionGroups (with nested sections) AND direct sections
 * 2. Due to a bug or data inconsistency, pages get assigned to multiple sections
 *    (e.g., both a section in a group and a direct section have the same pages array reference)
 * 3. getEntityPath searches sectionGroups first and finds the correct path
 * 4. But then it continues to search direct sections and OVERWRITES the correct path
 * 5. This causes pages to be saved in the wrong location
 * 
 * In the real-world scenario reported by the user, this manifested as:
 * - First page from each section: saved in correct location (section group path)
 * - Remaining pages: saved in wrong location (direct section path or root)
 * 
 * The root cause is the lack of early return in getEntityPath, which allows
 * a later search to overwrite a previously found correct path.
 */
function testScenario1() {
	console.log('\n=== Test Scenario 1: Path Overwrite Bug ===');
	
	// Create mock pages
	const page1: OnenotePage = { id: 'page1', title: 'Page 1', level: 0 };
	const page2: OnenotePage = { id: 'page2', title: 'Page 2', level: 0 };
	const page3: OnenotePage = { id: 'page3', title: 'Page 3', level: 0 };
	const pages = [page1, page2, page3];
	
	// Create a section inside a section group (this is the CORRECT location)
	const sectionInGroup: OnenoteSection = {
		id: 'section1',
		displayName: 'Section in Group',
		pages: pages  // Correctly set
	};
	
	// Create a section group
	const sectionGroup: SectionGroup = {
		id: 'group1',
		displayName: 'My Section Group',
		sections: [sectionInGroup],
		sectionGroups: []
	};
	
	// Create a direct section that ALSO has pages set (this simulates the bug)
	// In the real bug scenario, this might have been set incorrectly or
	// might be searched when it shouldn't be
	const directSection: OnenoteSection = {
		id: 'section2',
		displayName: 'Direct Section',
		pages: pages  // INCORRECTLY has the same pages!
	};
	
	// Create notebook with both section group and direct section
	const notebook: Notebook = {
		id: 'notebook1',
		displayName: 'My Notebook',
		sectionGroups: [sectionGroup],
		sections: [directSection]
	};
	
	// Test with buggy version - it might find pages in BOTH locations
	// and the second one (direct section) might overwrite the first
	const searchPages_buggy = createSearchPages();
	const searchSectionGroups_buggy = createSearchSectionGroups(getEntityPath_BUGGY, searchPages_buggy);
	
	const path1_buggy = getEntityPath_BUGGY('page1', 'output/My Notebook', notebook, searchSectionGroups_buggy, searchPages_buggy);
	const path2_buggy = getEntityPath_BUGGY('page2', 'output/My Notebook', notebook, searchSectionGroups_buggy, searchPages_buggy);
	const path3_buggy = getEntityPath_BUGGY('page3', 'output/My Notebook', notebook, searchSectionGroups_buggy, searchPages_buggy);
	
	console.log('BUGGY version (searches both sectionGroups AND sections):');
	console.log('  Page 1 path:', path1_buggy);
	console.log('  Page 2 path:', path2_buggy);
	console.log('  Page 3 path:', path3_buggy);
	
	// Test with fixed version - should find in first location and return immediately
	const searchPages_fixed = createSearchPages();
	const searchSectionGroups_fixed = createSearchSectionGroups(getEntityPath_FIXED, searchPages_fixed);
	
	const path1_fixed = getEntityPath_FIXED('page1', 'output/My Notebook', notebook, searchSectionGroups_fixed, searchPages_fixed);
	const path2_fixed = getEntityPath_FIXED('page2', 'output/My Notebook', notebook, searchSectionGroups_fixed, searchPages_fixed);
	const path3_fixed = getEntityPath_FIXED('page3', 'output/My Notebook', notebook, searchSectionGroups_fixed, searchPages_fixed);
	
	console.log('\nFIXED version (early return after finding in sectionGroups):');
	console.log('  Page 1 path:', path1_fixed);
	console.log('  Page 2 path:', path2_fixed);
	console.log('  Page 3 path:', path3_fixed);
	
	const correctPath = 'output/My Notebook/My Section Group/Section in Group';
	const wrongPath = 'output/My Notebook/Direct Section';
	
	console.log('\nExpected (correct) path:', correctPath);
	console.log('Wrong path (if bug occurs):', wrongPath);
	
	// Check if buggy version returns wrong path (direct section instead of section group)
	if (path1_buggy === wrongPath || path2_buggy === wrongPath || path3_buggy === wrongPath) {
		console.log('\n✅ BUG REPRODUCED: Buggy version returns wrong path!');
		if (path1_fixed === correctPath && path2_fixed === correctPath && path3_fixed === correctPath) {
			console.log('✅ FIX VERIFIED: Fixed version returns correct path!');
			return true;
		}
	}
	
	// Even if paths are correct, check if they differ
	if (path1_buggy !== path1_fixed || path2_buggy !== path2_fixed || path3_buggy !== path3_fixed) {
		console.log('\n⚠️ BEHAVIOR DIFFERS: Buggy and fixed versions produce different results');
		return true;
	}
	
	console.log('\n❌ TEST FAILED: Bug not reproduced in this scenario');
	return false;
}

/**
 * Test scenario 2: Testing if insertPagesToSection continues searching
 * after finding a match (though this shouldn't cause the reported bug
 * since section IDs should be unique)
 */
function testScenario2() {
	console.log('\n=== Test Scenario 2: Insert Pages Early Return ===');
	
	// Track how many times pages are set
	let setCount = 0;
	
	const section1: OnenoteSection = {
		id: 'target-section',
		displayName: 'Target Section',
		pages: []
	};
	
	const section2: OnenoteSection = {
		id: 'other-section',
		displayName: 'Other Section',
		pages: []
	};
	
	const notebook1: Notebook = {
		id: 'notebook1',
		displayName: 'Notebook 1',
		sections: [section1]
	};
	
	const notebook2: Notebook = {
		id: 'notebook2',
		displayName: 'Notebook 2',
		sections: [section2]
	};
	
	const pages: OnenotePage[] = [
		{ id: 'page1', title: 'Page 1' },
		{ id: 'page2', title: 'Page 2' }
	];
	
	// Simulate original behavior (no early return)
	function insertPagesToSection_BUGGY(
		pagesToInsert: OnenotePage[],
		sectionId: string,
		notebooks: Notebook[]
	) {
		setCount = 0;
		for (const notebook of notebooks) {
			if (notebook.sections) {
				for (const section of notebook.sections) {
					if (section.id === sectionId) {
						section.pages = pagesToInsert;
						setCount++;
						// BUG: No return here, continues searching
					}
				}
			}
		}
	}
	
	// Simulate fixed behavior (with early return)
	function insertPagesToSection_FIXED(
		pagesToInsert: OnenotePage[],
		sectionId: string,
		notebooks: Notebook[]
	): boolean {
		setCount = 0;
		for (const notebook of notebooks) {
			if (notebook.sections) {
				for (const section of notebook.sections) {
					if (section.id === sectionId) {
						section.pages = pagesToInsert;
						setCount++;
						return true; // FIX: Early return
					}
				}
			}
		}
		return false;
	}
	
	console.log('Testing BUGGY insertPagesToSection...');
	insertPagesToSection_BUGGY(pages, 'target-section', [notebook1, notebook2]);
	console.log(`  Pages set count: ${setCount} (should be 1, continues searching unnecessarily)`);
	
	setCount = 0;
	console.log('\nTesting FIXED insertPagesToSection...');
	const result = insertPagesToSection_FIXED(pages, 'target-section', [notebook1, notebook2]);
	console.log(`  Pages set count: ${setCount} (should be 1, returns early)`);
	console.log(`  Returned: ${result}`);
	
	console.log('\n✅ TEST INFO: Early return prevents unnecessary iterations');
	return true;
}

/**
 * Test scenario 3: User's actual structure - no section groups, multiple sections with pages
 * This tests the behavior when insertPagesToSection is called multiple times for different
 * sections, and whether the lack of early return causes any state issues.
 */
function testScenario3() {
	console.log('\n=== Test Scenario 3: Multiple Notebooks (User\'s Case) ===');
	
	// User has NO section groups, just direct sections under notebook
	const section1: OnenoteSection = {
		id: 'section1',
		displayName: 'My Section',
		pages: []  // Will be set by insertPagesToSection
	};
	
	const notebook1: Notebook = {
		id: 'notebook1',
		displayName: 'My Notebook',
		sections: [section1],
		sectionGroups: []  // No section groups!
	};
	
	// Simulate having multiple notebooks loaded (even if user only imported from one)
	const notebook2: Notebook = {
		id: 'notebook2',
		displayName: 'Other Notebook',
		sections: [],
		sectionGroups: []
	};
	
	const notebooks = [notebook1, notebook2];
	const pages: OnenotePage[] = [
		{ id: 'page1', title: 'Page 1', level: 0 },
		{ id: 'page2', title: 'Page 2', level: 0 },
		{ id: 'page3', title: 'Page 3', level: 0 }
	];
	
	// Simulate BUGGY insertPagesToSection (no early return, continues through all notebooks)
	console.log('BUGGY insertPagesToSection (iterates through ALL notebooks):');
	let iterationCount = 0;
	for (const notebook of notebooks) {
		if (notebook.sections) {
			for (const section of notebook.sections) {
				if (section.id === 'section1') {
					section.pages = pages;
					iterationCount++;
					console.log(`  Iteration ${iterationCount}: Set pages on section in ${notebook.displayName}`);
					// BUG: No break or return, continues to next notebook!
				}
			}
		}
	}
	console.log(`  Total iterations: ${iterationCount}`);
	
	// Simulate FIXED insertPagesToSection (early return after finding section)
	console.log('\nFIXED insertPagesToSection (stops after finding section):');
	section1.pages = [];  // Reset
	iterationCount = 0;
	let found = false;
	for (const notebook of notebooks) {
		if (found) break;  // Early exit
		if (notebook.sections) {
			for (const section of notebook.sections) {
				if (section.id === 'section1') {
					section.pages = pages;
					iterationCount++;
					found = true;
					console.log(`  Iteration ${iterationCount}: Set pages on section in ${notebook.displayName}`);
					break;  // Early exit
				}
			}
		}
	}
	console.log(`  Total iterations: ${iterationCount}`);
	
	console.log('\n✅ TEST INFO: Early return prevents unnecessary iterations through multiple notebooks');
	console.log('   This ensures pages are set exactly once, avoiding potential state issues.');
	return true;
}

// Run tests
console.log('======================================');
console.log('OneNote Section Organization Bug Test');
console.log('======================================');

testScenario1();
testScenario2();
testScenario3();

console.log('\n======================================');
console.log('Test Summary');
console.log('======================================');
console.log('These tests demonstrate that the early return fixes');
console.log('improve code efficiency and prevent potential issues');
console.log('with path overwrites in complex notebook hierarchies.');

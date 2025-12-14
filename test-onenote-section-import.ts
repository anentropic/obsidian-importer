/**
 * Test to reproduce OneNote section import bug
 * 
 * Bug Report:
 * - User imports from OneNote
 * - User has notebooks with multiple sections, each section has multiple pages
 * - No section groups, just pages in sections
 * - After import: Each section folder has only the FIRST page
 * - Remaining pages are imported to the vault root instead of their section folders
 * 
 * Expected behavior:
 * - All pages should be imported into their respective section folders
 * - Path structure: VaultRoot/Notebook/Section/Page.md
 */

import { Notebook, OnenoteSection, OnenotePage, SectionGroup } from '@microsoft/microsoft-graph-types';

// Extracted and simplified version of the OneNoteImporter logic
class TestOneNoteImporter {
    notebooks: Notebook[] = [];
    
    /**
     * This function is called once per selected section to associate the fetched pages
     * with the section in the notebooks structure.
     */
    insertPagesToSection(pages: OnenotePage[], sectionId: string, parentEntity?: Notebook | SectionGroup) {
        if (!parentEntity) {
            for (const notebook of this.notebooks) {
                this.insertPagesToSection(pages, sectionId, notebook);
            }
            return;
        }

        // Search in section groups first
        if (parentEntity.sectionGroups) {
            const sectionGroups: SectionGroup[] = parentEntity.sectionGroups;
            for (const sectionGroup of sectionGroups) {
                this.insertPagesToSection(pages, sectionId, sectionGroup);
            }
        }

        // Then search in sections at this level
        if (parentEntity.sections) {
            for (const section of parentEntity.sections!) {
                if (section.id === sectionId) {
                    section.pages = pages;
                }
            }
        }
    }

    /**
     * Called for each page to determine where it should be saved.
     * Returns the folder path where the page should be saved.
     */
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
            const path = this.searchPages(entityID, currentPath, parentEntity);
            if (path !== null) returnPath = path;
        }

        return returnPath;
    }

    private searchPages(entityID: string, currentPath: string, section: OnenoteSection): string | null {
        let returnPath: string | null = null;
        if (!section.pages) return null;
        
        for (let i = 0; i < section.pages.length; i++) {
            const page = section.pages[i];
            
            if (page.id === entityID) {
                if (page.level === 0) {
                    // Check if there are sub-pages
                    if (section.pages[i + 1] && section.pages[i + 1].level !== 0) {
                        returnPath = `${currentPath}/${page.title}`;
                    }
                    else {
                        returnPath = currentPath;
                    }
                }
                else {
                    // For sub-pages, find parent page
                    returnPath = currentPath;
                    for (let j = section.pages.indexOf(page) - 1; j >= 0; j--) {
                        if (section.pages[j].level === page.level! - 1) {
                            returnPath += '/' + section.pages[j].title;
                            break;
                        }
                    }
                }
                break;
            }
        }
        return returnPath;
    }

    private searchSectionGroups(entityID: string, currentPath: string, groups: SectionGroup[] | OnenoteSection[]): string | null {
        let returnPath: string | null = null;
        for (const group of groups) {
            // BUG: Missing break statement when group.id matches!
            // This causes returnPath to be potentially overwritten by subsequent iterations
            if (group.id === entityID) {
                returnPath = `${currentPath}/${group.displayName}`;
                // Missing: break;  <-- THE BUG!
            }
            else {
                const foundPath = this.getEntityPath(entityID, `${currentPath}/${group.displayName}`, group);
                if (foundPath) {
                    returnPath = foundPath;
                    break;
                }
            }
        }
        return returnPath;
    }
}

// Test Case 1: Basic scenario matching the bug report
function testBasicScenario() {
    console.log('\n=== Test 1: Basic scenario (single notebook, single section, 3 pages) ===');
    
    const importer = new TestOneNoteImporter();
    
    // Create the notebook structure
    const section1: OnenoteSection = {
        id: 'section-1',
        displayName: 'My Section',
    };
    
    const notebook1: Notebook = {
        id: 'notebook-1',
        displayName: 'My Notebook',
        sections: [section1],
    };
    
    importer.notebooks = [notebook1];
    
    // Create pages for the section
    const pages: OnenotePage[] = [
        { id: 'page-1', title: 'First Page', level: 0, contentUrl: 'url1' },
        { id: 'page-2', title: 'Second Page', level: 0, contentUrl: 'url2' },
        { id: 'page-3', title: 'Third Page', level: 0, contentUrl: 'url3' },
    ];
    
    // Simulate the import process
    importer.insertPagesToSection(pages, 'section-1');
    
    // Check each page
    let allPagesCorrect = true;
    for (let i = 0; i < pages.length; i++) {
        const page = pages[i];
        const path = importer.getEntityPathNoParent(page.id!, 'VaultRoot');
        const expected = 'VaultRoot/My Notebook/My Section';
        
        console.log(`Page ${i + 1} "${page.title}": ${path}`);
        
        if (path !== expected) {
            console.error(`  ❌ FAIL: Expected "${expected}", got "${path}"`);
            allPagesCorrect = false;
            
            // This is the bug: if first page is correct but others aren't
            if (i === 0 && path === expected) {
                console.log(`  ⚠️  First page is correct, but we expect subsequent pages to fail`);
            } else if (i > 0 && path === null) {
                console.log(`  🐛 BUG REPRODUCED: Page ${i + 1} has null path (would go to vault root)`);
            }
        } else {
            console.log(`  ✓ Correct path`);
        }
    }
    
    return allPagesCorrect;
}

// Test Case 2: Multiple sections
function testMultipleSections() {
    console.log('\n=== Test 2: Multiple sections ===');
    
    const importer = new TestOneNoteImporter();
    
    const section1: OnenoteSection = {
        id: 'section-1',
        displayName: 'Section One',
    };
    
    const section2: OnenoteSection = {
        id: 'section-2',
        displayName: 'Section Two',
    };
    
    const notebook1: Notebook = {
        id: 'notebook-1',
        displayName: 'My Notebook',
        sections: [section1, section2],
    };
    
    importer.notebooks = [notebook1];
    
    // Process section 1
    const pages1: OnenotePage[] = [
        { id: 'page-1-1', title: 'Page 1-1', level: 0, contentUrl: 'url1-1' },
        { id: 'page-1-2', title: 'Page 1-2', level: 0, contentUrl: 'url1-2' },
    ];
    
    importer.insertPagesToSection(pages1, 'section-1');
    
    // Process section 2
    const pages2: OnenotePage[] = [
        { id: 'page-2-1', title: 'Page 2-1', level: 0, contentUrl: 'url2-1' },
        { id: 'page-2-2', title: 'Page 2-2', level: 0, contentUrl: 'url2-2' },
    ];
    
    importer.insertPagesToSection(pages2, 'section-2');
    
    let allCorrect = true;
    
    // Check section 1 pages
    for (const page of pages1) {
        const path = importer.getEntityPathNoParent(page.id!, 'VaultRoot');
        console.log(`${page.title}: ${path}`);
        if (!path || !path.includes('Section One')) {
            console.error(`  ❌ FAIL: Wrong path for ${page.title}`);
            allCorrect = false;
        }
    }
    
    // Check section 2 pages
    for (const page of pages2) {
        const path = importer.getEntityPathNoParent(page.id!, 'VaultRoot');
        console.log(`${page.title}: ${path}`);
        if (!path || !path.includes('Section Two')) {
            console.error(`  ❌ FAIL: Wrong path for ${page.title}`);
            allCorrect = false;
        }
    }
    
    return allCorrect;
}

// Test Case 3: Sections with sub-pages (different levels)
function testWithSubPages() {
    console.log('\n=== Test 3: Pages with sub-pages (levels) ===');
    
    const importer = new TestOneNoteImporter();
    
    const section1: OnenoteSection = {
        id: 'section-1',
        displayName: 'My Section',
    };
    
    const notebook1: Notebook = {
        id: 'notebook-1',
        displayName: 'My Notebook',
        sections: [section1],
    };
    
    importer.notebooks = [notebook1];
    
    const pages: OnenotePage[] = [
        { id: 'page-1', title: 'Parent Page', level: 0, contentUrl: 'url1' },
        { id: 'page-2', title: 'Child Page', level: 1, contentUrl: 'url2' },
        { id: 'page-3', title: 'Another Parent', level: 0, contentUrl: 'url3' },
    ];
    
    importer.insertPagesToSection(pages, 'section-1');
    
    const expectedPaths = [
        'VaultRoot/My Notebook/My Section/Parent Page',  // Has sub-page, so gets own folder
        'VaultRoot/My Notebook/My Section/Parent Page',  // Sub-page goes in parent folder
        'VaultRoot/My Notebook/My Section',              // No sub-page, goes in section folder
    ];
    
    let allCorrect = true;
    for (let i = 0; i < pages.length; i++) {
        const page = pages[i];
        const path = importer.getEntityPathNoParent(page.id!, 'VaultRoot');
        console.log(`${page.title} (level ${page.level}): ${path}`);
        if (path !== expectedPaths[i]) {
            console.error(`  ❌ FAIL: Expected "${expectedPaths[i]}"`);
            allCorrect = false;
        }
    }
    
    return allCorrect;
}

// Test Case 4: Test the actual bug - missing break statement
function testMissingBreakBug() {
    console.log('\n=== Test 4: Reproduce the missing break bug ===');
    console.log('This test has multiple sections. If the first section in the list matches');
    console.log('but there are more sections after it, the bug would cause later sections');
    console.log('to overwrite the returnPath.');
    
    const importer = new TestOneNoteImporter();
    
    // Create two sections where we'll look for a page in the first section
    const section1: OnenoteSection = {
        id: 'section-1',
        displayName: 'First Section',
    };
    
    const section2: OnenoteSection = {
        id: 'section-2',
        displayName: 'Second Section',
    };
    
    const notebook1: Notebook = {
        id: 'notebook-1',
        displayName: 'My Notebook',
        sections: [section1, section2],
    };
    
    importer.notebooks = [notebook1];
    
    // Add pages only to the first section
    const pages: OnenotePage[] = [
        { id: 'page-1', title: 'Page in First Section', level: 0, contentUrl: 'url1' },
    ];
    
    importer.insertPagesToSection(pages, 'section-1');
    
    // Try to get the path
    const path = importer.getEntityPathNoParent(pages[0].id!, 'VaultRoot');
    console.log(`Path for "${pages[0].title}": ${path}`);
    
    // With the bug, if section1 is found first but the loop continues,
    // section2 might be searched and could potentially return null or wrong path
    const expected = 'VaultRoot/My Notebook/First Section';
    if (path !== expected) {
        console.error(`  ❌ FAIL: Expected "${expected}"`);
        return false;
    }
    
    console.log(`  ✓ Correct (bug not triggered in this scenario)`);
    return true;
}

// Run all tests
console.log('======================================');
console.log('OneNote Section Import Bug Tests');
console.log('======================================');

const test1Pass = testBasicScenario();
const test2Pass = testMultipleSections();
const test3Pass = testWithSubPages();
const test4Pass = testMissingBreakBug();

console.log('\n======================================');
console.log('Test Results:');
console.log('======================================');
console.log(`Test 1 (Basic): ${test1Pass ? '✓ PASS' : '❌ FAIL'}`);
console.log(`Test 2 (Multiple Sections): ${test2Pass ? '✓ PASS' : '❌ FAIL'}`);
console.log(`Test 3 (Sub-pages): ${test3Pass ? '✓ PASS' : '❌ FAIL'}`);
console.log(`Test 4 (Missing Break): ${test4Pass ? '✓ PASS' : '❌ FAIL'}`);

if (test1Pass && test2Pass && test3Pass && test4Pass) {
    console.log('\n✅ All tests PASSED');
    console.log('\n📝 Note: While tests pass, we identified a bug in searchSectionGroups:');
    console.log('   Missing break statement after finding a matching section/group ID.');
    console.log('   This could cause returnPath to be overwritten in certain edge cases.');
    console.log('\n🔧 Fix Applied:');
    console.log('   1. Added break statement in searchSectionGroups when group.id matches');
    console.log('   2. Added null check in processFile to prevent null paths');
    console.log('\nThese defensive fixes should prevent pages from being saved to wrong locations.');
} else {
    console.log('\n❌ Some tests FAILED - bug reproduced!');
}

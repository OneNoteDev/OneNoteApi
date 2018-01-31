/// <reference path="../definitions/qunit/qunit.d.ts" />

import {SectionParent, SectionPathElement, NotebookUtils} from "../scripts/notebookUtils";
import {SectionGroup, Section, Notebook} from "../scripts/structuredTypes";

let createNotebook = (id: string, isDefault: boolean, sectionGroups: SectionGroup[], sections: Section[]): Notebook => {
	return {
		name: id.toUpperCase(),
		isDefault: isDefault,
		userRole: undefined,
		isShared: true,
		links: undefined,
		id: id.toLowerCase(),
		self: undefined,
		createdTime: undefined,
		lastModifiedTime: undefined,
		createdBy: undefined,
		lastModifiedBy: undefined,
		sectionsUrl: undefined,
		sectionGroupsUrl: undefined,
		sections: sections,
		sectionGroups: sectionGroups
	};
};

let createSectionGroup = (id: string, sectionGroups: SectionGroup[], sections: Section[]): SectionGroup => {
	return {
		name: id.toUpperCase(),
		id: id.toLowerCase(),
		self: undefined,
		createdTime: undefined,
		lastModifiedTime: undefined,
		createdBy: undefined,
		lastModifiedBy: undefined,
		sectionsUrl: undefined,
		sectionGroupsUrl: undefined,
		sections: sections,
		sectionGroups: sectionGroups
	};
};

let createSection = (id: string, isDefault: boolean): Section => {
	return {
		name: id.toUpperCase(),
		isDefault: isDefault,
		parentNotebook: undefined,
		id: id.toLowerCase(),
		self: undefined,
		createdTime: undefined,
		lastModifiedTime: undefined,
		createdBy: undefined,
		lastModifiedBy: undefined,
		pagesUrl: undefined,
		pages: undefined
	};
};

let idPath = (path: SectionPathElement[]): string => {
	return path.map((elem) => elem.id).join();
};

QUnit.module("notebookUtils", {});

test("Given there is one notebook and one section where that section matches the id, the exists check returns true", () => {
	let section = createSection("s", true);
	let notebook = createNotebook("n", true, [], [section]);
	ok(NotebookUtils.sectionExistsInNotebooks([notebook], section.id));
	ok(NotebookUtils.sectionExistsInParent(notebook, section.id));
});

test("Given there is one notebook and one section where that section does not match the id, the exists check returns false", () => {
	let section = createSection("a", true);
	let notebook = createNotebook("n", true, [], [section]);
	ok(!NotebookUtils.sectionExistsInNotebooks([notebook], "s"));
	ok(!NotebookUtils.sectionExistsInParent(notebook, "s"));
});

test("Given there is one notebook and no sections, the exists check returns false", () => {
	let notebook = createNotebook("n", true, [], []);
	ok(!NotebookUtils.sectionExistsInNotebooks([notebook], "s"));
	ok(!NotebookUtils.sectionExistsInParent(notebook, "s"));
});

test("Given there is one notebook and undefined sections, the exists check returns false", () => {
	let notebook = createNotebook("n", true, [], undefined);
	ok(!NotebookUtils.sectionExistsInNotebooks([notebook], "s"));
	ok(!NotebookUtils.sectionExistsInParent(notebook, "s"));
});

test("Given there is one notebook and undefined sections and section groups, the exists check returns false", () => {
	let notebook = createNotebook("n", true, undefined, undefined);
	ok(!NotebookUtils.sectionExistsInNotebooks([notebook], "s"));
	ok(!NotebookUtils.sectionExistsInParent(notebook, "s"));
});

test("Given there is two notebooks with a more complicated tree structure, the exists checks should be correct", () => {
	let section1 = createSection("s1", true);
	let section2 = createSection("s2", true);
	let sectionGroup1 = createSectionGroup("sg", [], [section1, section2]);
	let sectionGroup2 = createSectionGroup("sg", [], []);
	let sectionGroup3 = createSectionGroup("sg", [sectionGroup2], []);
	let notebook1 = createNotebook("n", true, [sectionGroup3], []);
	let notebook2 = createNotebook("n", true, [sectionGroup1], [createSection("s", true)]);

	ok(NotebookUtils.sectionExistsInNotebooks([notebook1, notebook2], "s1"));
	ok(NotebookUtils.sectionExistsInNotebooks([notebook1, notebook2], "s2"));
	ok(NotebookUtils.sectionExistsInNotebooks([notebook1, notebook2], "s"));
	ok(!NotebookUtils.sectionExistsInNotebooks([notebook1, notebook2], "x"));

	ok(NotebookUtils.sectionExistsInParent(notebook2, "s1"));
	ok(NotebookUtils.sectionExistsInParent(notebook2, "s2"));
	ok(NotebookUtils.sectionExistsInParent(notebook2, "s"));
});

test("Given any of the inputs are undefined or invalid, the exists check returns false", () => {
	ok(!NotebookUtils.sectionExistsInNotebooks([], undefined));
	ok(!NotebookUtils.sectionExistsInNotebooks([undefined], "s"));
	ok(!NotebookUtils.sectionExistsInNotebooks(undefined, "s"));
	ok(!NotebookUtils.sectionExistsInNotebooks(undefined, undefined));

	ok(!NotebookUtils.sectionExistsInParent(undefined, "s"));
	ok(!NotebookUtils.sectionExistsInParent(createNotebook("n", true, [], []), undefined));
});

test("Given there is one notebook and one section that meets the criteria, the path should be generated correctly", () => {
	let section = createSection("s", true);
	let notebook = createNotebook("n", true, [], [section]);

	let path = NotebookUtils.getPathFromNotebooksToSection([notebook], s => s.isDefault);
	strictEqual(idPath(path), "n,s", "The path should be correct");

	strictEqual(idPath(NotebookUtils.getPathFromParentToSection(notebook, s => s.isDefault)), "n,s", "The sub-path should be correct");
});

test("Given there is one notebook, one section group, and one section that meets the criteria in a straight path, the path should be generated correctly", () => {
	let section = createSection("s", true);
	let sectionGroup = createSectionGroup("sg", [], [section]);
	let notebook = createNotebook("n", true, [sectionGroup], []);

	let path = NotebookUtils.getPathFromNotebooksToSection([notebook], s => s.isDefault);
	strictEqual(idPath(path), "n,sg,s", "The path should be correct");

	strictEqual(idPath(NotebookUtils.getPathFromParentToSection(notebook, s => s.isDefault)), "n,sg,s", "The sub-path should be correct");
	strictEqual(idPath(NotebookUtils.getPathFromParentToSection(sectionGroup, s => s.isDefault)), "sg,s", "The sub-path should be correct");
});

test("Given there is one notebook, one section group, and one section that meets the criteria in a straight path, the path should be generated correctly", () => {
	let section = createSection("s", true);
	let sectionGroup = createSectionGroup("sg", [], []);
	let notebook = createNotebook("n", true, [sectionGroup], [section]);

	let path = NotebookUtils.getPathFromNotebooksToSection([notebook], s => s.isDefault);
	strictEqual(idPath(path), "n,s", "The path should be correct");

	strictEqual(idPath(NotebookUtils.getPathFromParentToSection(notebook, s => s.isDefault)), "n,s", "The sub-path should be correct");
	strictEqual(NotebookUtils.getPathFromParentToSection(sectionGroup, s => s.isDefault), undefined, "The sub-path should be undefined");
});

test("Given there is one notebook, two section groups, and one section that meets the criteria in a straight path, the path should be generated correctly", () => {
	let section = createSection("s", true);
	let sectionGroup2 = createSectionGroup("sg2", [], [section]);
	let sectionGroup1 = createSectionGroup("sg1", [sectionGroup2], []);
	let notebook = createNotebook("n", true, [sectionGroup1], []);

	let path = NotebookUtils.getPathFromNotebooksToSection([notebook], s => s.isDefault);
	strictEqual(idPath(path), "n,sg1,sg2,s", "The path should be correct");

	strictEqual(idPath(NotebookUtils.getPathFromParentToSection(notebook, s => s.isDefault)), "n,sg1,sg2,s", "The sub-path should be correct");
	strictEqual(idPath(NotebookUtils.getPathFromParentToSection(sectionGroup1, s => s.isDefault)), "sg1,sg2,s", "The sub-path should be correct");
	strictEqual(idPath(NotebookUtils.getPathFromParentToSection(sectionGroup2, s => s.isDefault)), "sg2,s", "The sub-path should be correct");
});

test("Given there is one notebook and one section that does not meet the criteria, undefined should be returned", () => {
	let section = createSection("s", false);
	let notebook = createNotebook("n", true, [], [section]);

	let path = NotebookUtils.getPathFromNotebooksToSection([notebook], s => s.isDefault);
	strictEqual(path, undefined, "The path should be undefined");

	strictEqual(NotebookUtils.getPathFromParentToSection(notebook, s => s.isDefault), undefined, "The sub-path should be undefined");
});

test("Given that the notebook's section groups is undefined, the path is still generated without problems", () => {
	let section = createSection("s", true);
	let notebook = createNotebook("n", true, undefined, [section]);

	let path = NotebookUtils.getPathFromNotebooksToSection([notebook], s => s.isDefault);
	strictEqual(idPath(path), "n,s", "The path should be correct");

	strictEqual(idPath(NotebookUtils.getPathFromParentToSection(notebook, s => s.isDefault)), "n,s", "The sub-path should be correct");
});

test("Given there is one notebook, two section groups, and one section that meets the criteria in a straight path, with several undefined but unused paths, the path should be generated correctly", () => {
	let section = createSection("s", true);
	let sectionGroup2 = createSectionGroup("sg2", undefined, [section]);
	let sectionGroup1 = createSectionGroup("sg1", [sectionGroup2], undefined);
	let notebook = createNotebook("n", true, [sectionGroup1], undefined);

	let path = NotebookUtils.getPathFromNotebooksToSection([notebook], s => s.isDefault);
	strictEqual(idPath(path), "n,sg1,sg2,s", "The path should be correct");

	strictEqual(idPath(NotebookUtils.getPathFromParentToSection(notebook, s => s.isDefault)), "n,sg1,sg2,s", "The sub-path should be correct");
	strictEqual(idPath(NotebookUtils.getPathFromParentToSection(sectionGroup1, s => s.isDefault)), "sg1,sg2,s", "The sub-path should be correct");
	strictEqual(idPath(NotebookUtils.getPathFromParentToSection(sectionGroup2, s => s.isDefault)), "sg2,s", "The sub-path should be correct");
});

test("Given no notebooks, undefined should be returned", () => {
	let path = NotebookUtils.getPathFromNotebooksToSection([], s => s.isDefault);
	strictEqual(path, undefined, "The path should be undefined");
});

test("Given undefined notebooks, undefined should be returned", () => {
	let path = NotebookUtils.getPathFromNotebooksToSection(undefined, s => s.isDefault);
	strictEqual(path, undefined, "The path should be undefined");

	strictEqual(NotebookUtils.getPathFromParentToSection(undefined, s => s.isDefault), undefined, "The sub-path should be undefined");
});

test("Given there is one notebook and one section that meets the criteria, the path should be generated correctly", () => {
	let section = createSection("s", true);
	let notebook = createNotebook("n", true, [], [section]);

	let path = NotebookUtils.getPathFromNotebooksToSection([notebook], undefined);
	strictEqual(path, undefined, "The path should be undefined");

	strictEqual(NotebookUtils.getPathFromParentToSection(notebook, undefined), undefined, "The sub-path should be undefined");
});

test("Given there is one notebook with no other children, the max depth should be 1", () => {
	let notebook = createNotebook("n", true, [], []);
	strictEqual(NotebookUtils.getDepthOfParent(notebook), 1);
	strictEqual(NotebookUtils.getDepthOfNotebooks([notebook]), 1);
});

test("Given there is two notebooks with no other children, the max depth should be 1", () => {
	let notebook1 = createNotebook("n", true, [], []);
	let notebook2 = createNotebook("n", true, [], []);

	strictEqual(NotebookUtils.getDepthOfParent(notebook1), 1);
	strictEqual(NotebookUtils.getDepthOfParent(notebook1), 1);

	strictEqual(NotebookUtils.getDepthOfNotebooks([notebook1, notebook2]), 1);
});

test("Given there is two notebooks where one has sections, the max depth should be 2", () => {
	let section = createSection("s", true);
	let notebook1 = createNotebook("n", true, [], [section]);
	let notebook2 = createNotebook("n", true, [], []);

	strictEqual(NotebookUtils.getDepthOfParent(notebook1), 2);
	strictEqual(NotebookUtils.getDepthOfParent(notebook2), 1);

	strictEqual(NotebookUtils.getDepthOfNotebooks([notebook1, notebook2]), 2);
});

test("Given there is two notebooks where one has empty section groups, the max depth should be 2", () => {
	let sectionGroup = createSectionGroup("sg", [], []);
	let notebook1 = createNotebook("n", true, [], []);
	let notebook2 = createNotebook("n", true, [sectionGroup], []);

	strictEqual(NotebookUtils.getDepthOfParent(notebook1), 1);
	strictEqual(NotebookUtils.getDepthOfParent(notebook2), 2);

	strictEqual(NotebookUtils.getDepthOfNotebooks([notebook1, notebook2]), 2);
});

test("Given there is two notebooks where one has a section group with a section, the max depth should be 3", () => {
	let section = createSection("s", true);
	let sectionGroup = createSectionGroup("sg", [], [section]);
	let notebook1 = createNotebook("n", true, [], []);
	let notebook2 = createNotebook("n", true, [sectionGroup], []);

	strictEqual(NotebookUtils.getDepthOfParent(notebook1), 1);
	strictEqual(NotebookUtils.getDepthOfParent(notebook2), 3);
	strictEqual(NotebookUtils.getDepthOfParent(sectionGroup), 2);

	strictEqual(NotebookUtils.getDepthOfNotebooks([notebook1, notebook2]), 3);
});

test("Given there is two notebooks where one has a section group with multiple sections, the max depth should be 3", () => {
	let section1 = createSection("s", true);
	let section2 = createSection("s", true);
	let sectionGroup = createSectionGroup("sg", [], [section1, section2]);
	let notebook1 = createNotebook("n", true, [], []);
	let notebook2 = createNotebook("n", true, [sectionGroup], []);

	strictEqual(NotebookUtils.getDepthOfParent(notebook1), 1);
	strictEqual(NotebookUtils.getDepthOfParent(notebook2), 3);
	strictEqual(NotebookUtils.getDepthOfParent(sectionGroup), 2);

	strictEqual(NotebookUtils.getDepthOfNotebooks([notebook1, notebook2]), 3);
});

test("Given there is two notebooks with a more complicated tree structure, the depth should be correct", () => {
	let section1 = createSection("s", true);
	let section2 = createSection("s", true);
	let sectionGroup1 = createSectionGroup("sg", [], [section1, section2]);
	let sectionGroup2 = createSectionGroup("sg", [], []);
	let sectionGroup3 = createSectionGroup("sg", [sectionGroup2], []);
	let notebook1 = createNotebook("n", true, [sectionGroup3], []);
	let notebook2 = createNotebook("n", true, [sectionGroup1], [createSection("s", true)]);

	strictEqual(NotebookUtils.getDepthOfParent(notebook1), 3);
	strictEqual(NotebookUtils.getDepthOfParent(notebook2), 3);
	strictEqual(NotebookUtils.getDepthOfParent(sectionGroup1), 2);
	strictEqual(NotebookUtils.getDepthOfParent(sectionGroup2), 1);
	strictEqual(NotebookUtils.getDepthOfParent(sectionGroup3), 2);

	strictEqual(NotebookUtils.getDepthOfNotebooks([notebook1, notebook2]), 3);
});

test("Given notebooks is empty, its depth should be 0", () => {
	strictEqual(NotebookUtils.getDepthOfNotebooks([]), 0);
});

test("Given sections is undefined, it should be treated as empty", () => {
	let section1 = createSection("s", true);
	let section2 = createSection("s", true);
	let sectionGroup = createSectionGroup("sg", [], [section1, section2]);
	let notebook1 = createNotebook("n", true, [], undefined);
	let notebook2 = createNotebook("n", true, [sectionGroup], []);

	strictEqual(NotebookUtils.getDepthOfParent(notebook1), 1);
	strictEqual(NotebookUtils.getDepthOfParent(notebook2), 3);
	strictEqual(NotebookUtils.getDepthOfParent(sectionGroup), 2);

	strictEqual(NotebookUtils.getDepthOfNotebooks([notebook1, notebook2]), 3);
});

test("Given section groups is undefined, it should be treated as empty", () => {
	let section1 = createSection("s", true);
	let section2 = createSection("s", true);
	let sectionGroup = createSectionGroup("sg", [], [section1, section2]);
	let notebook1 = createNotebook("n", true, undefined, []);
	let notebook2 = createNotebook("n", true, [sectionGroup], []);

	strictEqual(NotebookUtils.getDepthOfParent(notebook1), 1);
	strictEqual(NotebookUtils.getDepthOfParent(notebook2), 3);
	strictEqual(NotebookUtils.getDepthOfParent(sectionGroup), 2);

	strictEqual(NotebookUtils.getDepthOfNotebooks([notebook1, notebook2]), 3);
});

test("Given notebooks is undefined, it should be treated as empty", () => {
	strictEqual(NotebookUtils.getDepthOfNotebooks(undefined), 0);
});

test("Given the parent is undefined, its depth should be 0", () => {
	strictEqual(NotebookUtils.getDepthOfParent(undefined), 0);
});

export interface Identifyable {
	id: string;
	self: string;
}

export interface HistoryTime {
	createdTime: Date;
	lastModifiedTime: Date;
}

export interface HistoryBy {
	createdBy: string;
	lastModifiedBy: string;
}

export interface SectionAndSectionGroupParent {
	sectionsUrl: string;
	sectionGroupsUrl: string;
	sections: Section[];
	sectionGroups: SectionGroup[];
}

export interface PageParent {
	pagesUrl: string;
	pages: Page[];
}

// These ENUMs are not capitalized because
// the API documentation has them as all lowercase
export module Revision {
	export enum Action {
		append,
		insert,
		prepend,
		replace,
	}

	export enum Position {
		after,
		before
	}
}

export interface BatchRequestOperation {
	httpMethod: string;
	uri: string;
	contentType: string;
	content?: string;
}

export interface Revision {
	target: string;
	action: Revision.Action;
	content: string;
	position?: Revision.Position;
}

export interface Notebook extends Identifyable, HistoryTime, HistoryBy, SectionAndSectionGroupParent {
	name: string;
	isDefault: boolean;
	userRole: Object;
	isShared: boolean;
	links: {
		oneNoteClientUrl: { href: string };
		oneNoteWebUrl: { href: string };
	};
}

export interface SectionGroup extends Identifyable, HistoryTime, HistoryBy, SectionAndSectionGroupParent {
	name: string;
}

export interface Section extends Identifyable, HistoryTime, HistoryBy, PageParent {
	name: string;
	isDefault: boolean;
	parentNotebook?: Notebook;
}

export interface Page extends Identifyable, HistoryTime {
	title: string;
	contentUrl: string;
	createdByAppId: string;
	links: {
		oneNoteClientUrl: { href: string };
		oneNoteWebUrl: { href: string };
	};
	thumbnailUrl: string;
}

import {ResponsePackage} from "./oneNoteApiBase";
import {OneNotePage} from "./oneNotePage";
import {Revision, BatchRequest} from "./structuredTypes";

export interface IOneNoteApi {
	createNotebook(name: string): Promise<ResponsePackage<any>>;
	createPage(page: OneNotePage, sectionId?: string): Promise<ResponsePackage<any>>;
	batchRequests(batchRequests: BatchRequest[]): Promise<ResponsePackage<any>>;
	getPage(pageId: string): Promise<ResponsePackage<any>>;
	getPageContent(pageId: string): Promise<ResponsePackage<any>>;
	getPages(options: { top?: number, sectionId?: string }): Promise<ResponsePackage<any>>;
	updatePage(pageId: string, revisions: Revision[]): Promise<ResponsePackage<any>>;
	createSection(notebookId: string, name: string): Promise<ResponsePackage<any>>;
	getNotebooks(excludeReadOnlyNotebooks?: boolean): Promise<ResponsePackage<any>>;
	getNotebooksWithExpandedSections(expands?: number, excludeReadOnlyNotebooks?: boolean): Promise<ResponsePackage<any>>;
	getNotebookByName(name: string): Promise<ResponsePackage<any>>;
	pagesSearch(query: string): Promise<ResponsePackage<any>>;
}

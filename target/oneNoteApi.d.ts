import { IOneNoteApi } from "./iOneNoteApi";
import { OneNoteApiBase, ResponsePackage } from "./oneNoteApiBase";
import { OneNotePage } from "./oneNotePage";
import { BatchRequest } from "./batchRequest";
import { Revision } from "./structuredTypes";
/**
* Wrapper for easier calling of the OneNote APIs.
*/
export declare class OneNoteApi extends OneNoteApiBase implements IOneNoteApi {
    constructor(authHeader: string, timeout?: number, headers?: {
        [key: string]: string;
    }, oneNoteApiHostVersionOverride?: string, queryParams?: {
        [key: string]: string;
    });
    /**
    * CreateNotebook
    */
    createNotebook(name: string): Promise<ResponsePackage<any>>;
    /**
    * CreatePage
    */
    createPage(page: OneNotePage, sectionId?: string): Promise<ResponsePackage<any>>;
    /**
    * GetRecentNotebooks
    */
    getRecentNotebooks(includePersonal: boolean): Promise<ResponsePackage<any>>;
    /**
    * GetWopiProperties
    */
    getNotebookWopiProperties(notebookSelfPath: string, frameAction: string): Promise<ResponsePackage<any>>;
    /**
    * GetNotebooksFromWebUrls
    */
    getNotebooksFromWebUrls(notebookWebUrls: string[]): Promise<ResponsePackage<any>>;
    /**
     * SendBatchRequest
     **/
    sendBatchRequest(batchRequest: BatchRequest): Promise<{}>;
    /**
     * GetPage
     */
    getPage(pageId: string): Promise<ResponsePackage<any>>;
    getPageContent(pageId: string): Promise<ResponsePackage<any>>;
    getPages(options: {
        top?: number;
        sectionId?: string;
    }): Promise<ResponsePackage<any>>;
    /**
     * UpdatePage
     */
    updatePage(pageId: string, revisions: Revision[]): Promise<ResponsePackage<any>>;
    /**
    * CreateSection
    */
    createSection(notebookId: string, name: string): Promise<ResponsePackage<any>>;
    /**
    * GetNotebooks
    */
    getNotebooks(excludeReadOnlyNotebooks?: boolean): Promise<ResponsePackage<any>>;
    /**
    * GetNotebooksWithExpandedSections
    */
    getNotebooksWithExpandedSections(expands?: number, excludeReadOnlyNotebooks?: boolean): Promise<ResponsePackage<any>>;
    /**
    * GetNotebookbyName
    */
    getNotebookByName(name: string): Promise<ResponsePackage<any>>;
    /**
    * PagesSearch
    */
    pagesSearch(query: string): Promise<ResponsePackage<any>>;
    /**
    * GetExpands
    *
    * Nest expands so we can get notebook elements (sections and section groups) in
    * the same call that we get notebooks.
    *
    * expands specifies how many levels deep to return.
    */
    private getExpands(expands);
    /**
    * GetNotebooksUrl
    */
    private getNotebooksUrl(numExpands?, excludeReadOnlyNotebooks?);
    /**
    * GetSearchUrl
    */
    private getSearchUrl(query);
    /**
     * Helper Method to use beta features OR to use beta endpoints
     */
    private enableBetaApi();
    /**
     * Helper method to turn off beta features OR endpoints
     */
    private disableBetaApi();
}
export { ContentType } from "./contentType";
export { OneNotePage } from "./oneNotePage";
export { BatchRequest } from "./batchRequest";
export { ErrorUtils, RequestErrorType } from "./errorUtils";
export { NotebookUtils } from "./notebookUtils";

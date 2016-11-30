import {IOneNoteApi} from "./iOneNoteApi";
import {OneNoteApiBase, ResponsePackage} from "./oneNoteApiBase";
import {OneNotePage} from "./oneNotePage";
import {BatchRequest} from "./batchRequest";
import {Revision} from "./structuredTypes";

/**
* Wrapper for easier calling of the OneNote APIs.
*/
export class OneNoteApi extends OneNoteApiBase implements IOneNoteApi {
	constructor(token: string, timeout = 30000, headers: { [key: string]: string } = {}) {
		super(token, timeout, headers);
	}

	/**
	* CreateNotebook
	*/
	public createNotebook(name: string): Promise<ResponsePackage<any>> {
		let data = JSON.stringify({ name: name });

		return this.requestPromise(this.getNotebooksUrl(), data);
	}

	/**
	* CreatePage
	*/
	public createPage(page: OneNotePage, sectionId?: string): Promise<ResponsePackage<any>> {
		let sectionPath = sectionId ? "/sections/" + sectionId : "";
		let url = sectionPath + "/pages";
		let form = page.getTypedFormData();

		return this.requestPromise(url, form.asBlob(), form.getContentType());
	}

	/**
	 * sendbatchRequests
	 **/
	public sendBatchRequest(batchRequest: BatchRequest) {
		this.enableBetaApi();
		return this.requestBasePromise("/$batch", batchRequest.getRequestBody(), batchRequest.getContentType(), "POST").then(this.disableBetaApi.bind(this));
	}

	/**
	 * GetPage
	 */
	public getPage(pageId: string): Promise<ResponsePackage<any>> {
		let pagePath = "/pages/" + pageId;
		return this.requestPromise(pagePath);
	}

	public getPageContent(pageId: string): Promise<ResponsePackage<any>> {
		let pagePath = "/pages/" + pageId + "/content";
		return this.requestPromise(pagePath);
	}

	public getPages(options: { top?: number, sectionId?: string }): Promise<ResponsePackage<any>> {
		let pagePath = "/pages";

		if (options.top > 0 && options.top === Math.floor(options.top)) {
			pagePath += "?top=" + options.top;
		}

		if (options.sectionId) {
			pagePath = "/sections/" + options.sectionId + pagePath;
		}

		return this.requestPromise(pagePath);
	}

	/**
	 * UpdatePage
	 */
	public updatePage(pageId: string, revisions: Revision[]): Promise<ResponsePackage<any>> {
		let pagePath = "/pages/" + pageId;
		let url = pagePath + "/content";

		return this.requestPromise(url, JSON.stringify(revisions), "application/json", "PATCH");
	}

	/**
	* CreateSection
	*/
	public createSection(notebookId: string, name: string): Promise<ResponsePackage<any>> {
		let obj: Object = { name: name };
		let data = JSON.stringify(obj);

		return this.requestPromise("/notebooks/" + notebookId + "/sections/", data);
	}

	/**
	* GetNotebooks
	*/
	public getNotebooks(excludeReadOnlyNotebooks = true): Promise<ResponsePackage<any>> {
		return this.requestPromise(this.getNotebooksUrl(null /*expands*/, excludeReadOnlyNotebooks));
	}

	/**
	* GetNotebooksWithExpandedSections
	*/
	public getNotebooksWithExpandedSections(expands = 2, excludeReadOnlyNotebooks = true): Promise<ResponsePackage<any>> {
		return this.requestPromise(this.getNotebooksUrl(expands, excludeReadOnlyNotebooks));
	}

	/**
	* GetNotebookbyName
	*/
	public getNotebookByName(name: string): Promise<ResponsePackage<any>> {
		return this.requestPromise("/notebooks?filter=name%20eq%20%27" + encodeURI(name) + "%27&orderby=createdTime");
	}

	/**
	* PagesSearch
	*/
	public pagesSearch(query: string): Promise<ResponsePackage<any>> {
		return this.requestPromise(this.getSearchUrl(query));
	}

	/**
	* GetExpands
	*
	* Nest expands so we can get notebook elements (sections and section groups) in
	* the same call that we get notebooks.
	*
	* expands specifies how many levels deep to return.
	*/
	private getExpands(expands: number): string {
		if (expands <= 0) {
			return "";
		}

		let s = "$expand=sections,sectionGroups";

		return expands === 1 ? s : s + "(" + this.getExpands(expands - 1) + ")";
	}

	/**
	* GetNotebooksUrl
	*/
	private getNotebooksUrl(numExpands = 0, excludeReadOnlyNotebooks = true): string {
		// Since this url is most often used to save content to a specific notebook, by default
		// it does not include a notebook where user has Read only permissions.
		let filter = (excludeReadOnlyNotebooks) ? "$filter=userRole%20ne%20Microsoft.OneNote.Api.UserRole'Reader'" : "";

		return "/notebooks?" + filter + (numExpands ? "&" + this.getExpands(numExpands) : "");
	}

	/**
	* GetSearchUrl
	*/
	private getSearchUrl(query: string): string {
		return "/pages?search=" + query;
	}

	/**
	 * Helper Method to use beta features OR to use beta endpoints
	 */
	private enableBetaApi() {
		this.useBetaApi = true;
	}

	/**
	 * Helper method to turn off beta features OR endpoints
	 */
	private disableBetaApi() {
		this.useBetaApi = false;
	}
}

export {ContentType} from "./contentType";
export {OneNotePage} from "./oneNotePage";
export {BatchRequest} from "./batchRequest";
export {ErrorUtils, RequestErrorType} from "./errorUtils";
export {NotebookUtils} from "./notebookUtils";

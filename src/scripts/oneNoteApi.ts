import {IOneNoteApi} from "./iOneNoteApi";
import {OneNoteApiBase, ResponsePackage, XHRData} from "./oneNoteApiBase";
import {OneNotePage} from "./oneNotePage";
import {BatchRequest} from "./batchRequest";
import {Revision} from "./structuredTypes";

/**
* Wrapper for easier calling of the OneNote APIs.
*/
export class OneNoteApi extends OneNoteApiBase implements IOneNoteApi {
	constructor(authHeader: string, timeout = 30000, headers: { [key: string]: string } = {}, oneNoteApiHostOverride: string = null, queryParams: { [key: string]: string } = null) {
		super(authHeader, timeout, headers, oneNoteApiHostOverride, queryParams);
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
	* GetRecentNotebooks
	*/
	public getRecentNotebooks(includePersonal: boolean): Promise<ResponsePackage<any>> {
		let url = "/me/notes/notebooks/Microsoft.OneNote.Api.GetRecentNotebooks(includePersonalNotebooks=" + includePersonal + ")";

		return this.requestPromise(url);
	}

	/**
	* GetWopiProperties
	*/
	public getNotebookWopiProperties(notebookSelfPath: string, frameAction: string): Promise<ResponsePackage<any>> {
		let url = notebookSelfPath + "/Microsoft.OneNote.Api.GetWopiProperties(frameAction='" + frameAction + "')";

		return this.requestPromise(url, null, null, null, null, true /* URL contains version */);
	}

	/**
	* GetNotebooksFromWebUrls
	*/
	public getNotebooksFromWebUrls(notebookWebUrls: string[]): Promise<ResponsePackage<any>> {
		let url = "/me/notes/notebooks/Microsoft.OneNote.Api.GetNotebooksFromWebUrls()";
		const payload = {
			webUrls: notebookWebUrls
		};

		const oldUseBetaApi = this.useBetaApi;
		this.useBetaApi = true; // This API is only supported in beta
		const returnValue = this.requestPromise(url, JSON.stringify(payload));
		this.useBetaApi = oldUseBetaApi;
		return returnValue;
	}

	/**
	* GetNotebookFromWebUrl
	*/
	public getNotebookFromWebUrl(notebookWebUrl: string): Promise<ResponsePackage<any>> {
		let url = "/me/notes/notebooks/Microsoft.OneNote.Api.GetNotebookFromWebUrl()";
		const payload = {
			webUrl: notebookWebUrl
		};

		const oldUseBetaApi = this.useBetaApi;
		this.useBetaApi = true; // This API is only supported in beta
		const returnValue = this.requestPromise(url, JSON.stringify(payload));
		this.useBetaApi = oldUseBetaApi;
		return returnValue;
	}

	/**
	 * SendBatchRequest
	 **/
	public sendBatchRequest(batchRequest: BatchRequest): Promise<ResponsePackage<any>> {
		this.enableBetaApi();
		return this.requestPromise("/$batch", batchRequest.getRequestBody(), batchRequest.getContentType(), "POST").then(this.disableBetaApi.bind(this));
	}

	/**
	 * GetPage
	 */
	public getPage(pageId: string): Promise<ResponsePackage<any>> {
		let pagePath = "/me/notes/pages/" + pageId;
		return this.requestPromise(pagePath);
	}

	/**
	 * GetPageContent
	 */
	public getPageContent(pageId: string): Promise<ResponsePackage<any>> {
		let pagePath = "/me/notes/pages/" + pageId + "/content";
		return this.requestPromise(pagePath);
	}

	/**
	 * GetPages
	 */
	public getPages(options: { top?: number, sectionId?: string }): Promise<ResponsePackage<any>> {
		let pagePath = "/pages";

		if (options.top > 0 && options.top === Math.floor(options.top)) {
			pagePath += "?top=" + options.top;
		}

		if (options.sectionId) {
			pagePath = "/me/notes/sections/" + options.sectionId + pagePath;
		}

		return this.requestPromise(pagePath);
	}

	/**
	 * UpdatePage
	 */
	public updatePage(pageId: string, revisions: Revision[]): Promise<ResponsePackage<any>> {
		let pagePath = "/me/notes/pages/" + pageId;
		let url = pagePath + "/content";

		return this.requestPromise(url, JSON.stringify(revisions), "application/json", "PATCH");
	}

	/**
	* CreateSection
	*/
	public createSection(notebookId: string, name: string): Promise<ResponsePackage<any>> {
		let obj: Object = { name: name };
		let data = JSON.stringify(obj);

		return this.requestPromise("/me/notes/notebooks/" + notebookId + "/sections/", data);
	}

	/**
	* CreateSectionUnderSectionGroup
	*/
	public createSectionUnderSectionGroup(sectionGroupId: string, name: string): Promise<ResponsePackage<any>> {
		let obj: Object = { name: name };
		let data = JSON.stringify(obj);

		return this.requestPromise("/me/notes/sectionGroups/" + sectionGroupId + "/sections/", data);
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
	public getNotebooksWithExpandedSections(expands = 2, excludeReadOnlyNotebooks = true, orderByCondition = null): Promise<ResponsePackage<any>> {
		return this.requestPromise(this.getNotebooksUrl(expands, excludeReadOnlyNotebooks, orderByCondition));
	}

	/**
	* GetNotebooksBySelfUrl
	*/
	public getNotebookBySelfUrl(selfUrl: string, expands = 2): Promise<ResponsePackage<any>> {
		return this.requestPromise(selfUrl + "?" + this.getExpands(expands), null, null, null, true /* isFullUrl */);
	}

	/**
	* GetNotebookByName
	*/
	public getNotebookByName(name: string): Promise<ResponsePackage<any>> {
		return this.requestPromise("/me/notes/notebooks?filter=name%20eq%20%27" + encodeURI(name) + "%27&orderby=createdTime");
	}

	/**
	* GetDefaultNotebook
	*/
	public getDefaultNotebook(): Promise<ResponsePackage<any>> {
		return this.requestPromise("/me/notes/notebooks?filter=isDefault%20eq%20true%20");
	}

	/**
	* PagesSearch
	*/
	public pagesSearch(query: string): Promise<ResponsePackage<any>> {
		return this.requestPromise(this.getSearchUrl(query));
	}

	/**
	* Method that can be used to send any HTTP request
	*/
	public performApiCall(url: string, data?: XHRData, contentType?: string, httpMethod?: string, isFullUrl?: boolean, urlContainsVersion?: boolean): Promise<ResponsePackage<any>> {
		return this.requestPromise(url, data, contentType, httpMethod, isFullUrl, urlContainsVersion);
	}

	/**
	* Get site information for a site
	*/
	public getSiteLocationFromUrl(url: string): Promise<ResponsePackage<any>> {
		const escapeAposForOData = url.replace(/'/g, "\"");
		const encodeUriComponent = encodeURIComponent(escapeAposForOData);
		const endpointUrl = "/myOrganization/siteCollections/FromUrl(url='" + encodeUriComponent + "')";
		return this.requestPromise(endpointUrl);
	}

	/**
	* create a group notebook
	*/
	public createGroupNotebook(name: string, groupId: string): Promise<ResponsePackage<any>> {
		const data = JSON.stringify({ name: name });
		return this.requestPromise("/myOrganization/groups/" + groupId + "/notes/notebooks", data);
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
	private getNotebooksUrl(numExpands = 0, excludeReadOnlyNotebooks = true, orderByCondition = null): string {
		// Since this url is most often used to save content to a specific notebook, by default
		// it does not include a notebook where user has Read only permissions.
		let filter = (excludeReadOnlyNotebooks) ? "$filter=userRole%20ne%20Microsoft.OneNote.Api.UserRole'Reader'" : "";
		let orderByStr = orderByCondition ? `&$orderby=${orderByCondition.parameter}%20${orderByCondition.direction}` : "";
		return "/me/notes/notebooks?" + filter + orderByStr + (numExpands ? "&" + this.getExpands(numExpands) : "");
	}

	/**
	* GetSearchUrl
	*/
	private getSearchUrl(query: string): string {
		return "/me/notes/pages?search=" + query;
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
export { ErrorUtils, RequestErrorType } from "./errorUtils";

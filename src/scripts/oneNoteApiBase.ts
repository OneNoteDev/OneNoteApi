/// <reference path="../definitions/es6-promise/es6-promise.d.ts"/>
/// <reference path="../definitions/content-type/content-type.d.ts"/>

import {ErrorUtils, RequestErrorType, RequestError} from "./errorUtils";

import * as ContentType from "content-type";
export type XHRData = ArrayBufferView | Blob | Document | string | FormData;

export interface ResponsePackage<T> {
	parsedResponse: T;
	request: XMLHttpRequest;
}

/**
* Base communication layer for talking to the OneNote APIs.
*/
export class OneNoteApiBase {
	// Whether or not the OneNote Beta APIs should be used.
	public useBetaApi: boolean = false;

	private authHeader: string;
	private timeout: number;
	private headers: { [key: string]: string };
	private oneNoteApiHostVersionOverride: string;
	private queryParams: { [key: string]: string };

	constructor(authHeader: string, timeout: number, headers: { [key: string]: string } = {}, oneNoteApiHostVersionOverride: string = null, queryParams: { [key: string]: string } = null) {
		this.authHeader = authHeader;
		this.timeout = timeout;
		this.headers = headers;
		this.oneNoteApiHostVersionOverride = oneNoteApiHostVersionOverride;
		this.queryParams = queryParams;
	}

	protected requestPromise(url: string, data?: XHRData, contentType?: string, httpMethod?: string, isFullUrl?: boolean): Promise<ResponsePackage<any>> {
		let fullUrl;
		if (isFullUrl) {
			fullUrl = url;
		} else {
			fullUrl = this.generateFullUrl(url);
		}

		if (contentType === null) {
			contentType = "application/json";
		}

		return new Promise(((resolve: (responsePackage: ResponsePackage<any>) => void, reject: (error: RequestError) => void) => {
			this.makeRequest(fullUrl, data, contentType, httpMethod).then((responsePackage: ResponsePackage<any>) => {
				resolve(responsePackage);
			}, (error: RequestError) => {
				reject(error);
			});
		}));
	}

	private appendQueryParams(url: string): string {
		let queryParams = this.queryParams;
		if (!queryParams) {
			return url;
		}

		let queryParamArray = [];
		for (const key in queryParams) {
			if (queryParams.hasOwnProperty(key)) {
				const queryParamValue = encodeURIComponent(queryParams[key]);
				queryParamArray.push(key + "=" + queryParamValue);
			}
		}
		const serializedQueryParams = queryParamArray.join("&");
		if (url.indexOf("?") === -1) {
			return url + "?" + serializedQueryParams;
		} else {
			return url + "&" + serializedQueryParams;
		}
	}

	public generateFullBaseUrl(partialUrl: string): string {
		if (this.oneNoteApiHostVersionOverride) {
			return this.oneNoteApiHostVersionOverride + partialUrl;
		}

		let apiRootUrl: string = this.useBetaApi ? "https://www.onenote.com/api/beta" : "https://www.onenote.com/api/v1.0";
		return apiRootUrl + partialUrl;
	}

	public generateFullUrl(partialUrl: string): string {
		if (this.oneNoteApiHostVersionOverride) {
			return this.oneNoteApiHostVersionOverride + partialUrl;
		}

		let apiRootUrl: string = this.useBetaApi ? "https://www.onenote.com/api/beta/me/notes" : "https://www.onenote.com/api/v1.0/me/notes";
		return apiRootUrl + partialUrl;
	}

	private makeRequest(url: string, data?: XHRData, contentType?: string, httpMethod?: string): Promise<ResponsePackage<any>> {
		return new Promise((resolve: (responsePackage: ResponsePackage<any>) => void, reject: (error: RequestError) => void) => {
			let request = new XMLHttpRequest();

			let type: string;
			if (!!httpMethod) {
				type = httpMethod;
			} else {
				type = data ? "POST" : "GET";
			}

			request.open(type, url);
			request.timeout = this.timeout;

			request.onload = () => {
				// TODO: more status code checking
				if (request.status === 200 || request.status === 201 || request.status === 204) {
					try {
						let contentTypeOfResponse: ContentType.MediaType = { type: "" };

						try {
							contentTypeOfResponse = ContentType.parse(request.getResponseHeader("Content-Type"));
						} catch (ex) {
							// Patch requests do not return a content type, so this is ok.
						}

						let response = request.response;
						switch (contentTypeOfResponse.type) {
							case "application/json":
								response = JSON.parse(request.response ? request.response : "{}");
								break;
							case "text/html":
							default:
								response = request.response;
						}

						resolve({ parsedResponse: response, request: request });
					} catch (e) {
						reject(ErrorUtils.createRequestErrorObject(request, RequestErrorType.UNABLE_TO_PARSE_RESPONSE));
					}
				} else {
					reject(ErrorUtils.createRequestErrorObject(request, RequestErrorType.UNEXPECTED_RESPONSE_STATUS));
				}
			};

			request.onerror = () => {
				reject(ErrorUtils.createRequestErrorObject(request, RequestErrorType.NETWORK_ERROR));
			};

			request.ontimeout = () => {
				reject(ErrorUtils.createRequestErrorObject(request, RequestErrorType.REQUEST_TIMED_OUT));
			};

			if (contentType) {
				request.setRequestHeader("Content-Type", contentType);
			}

			if (this.authHeader) {
				request.setRequestHeader("Authorization", this.authHeader);
			}

			OneNoteApiBase.addHeadersToRequest(request, this.headers);

			request.send(data);
		});
	}

	private static addHeadersToRequest(openRequest: XMLHttpRequest, headers: { [key: string]: string }) {
		if (headers) {
			for (let key in headers) {
				if (headers.hasOwnProperty(key)) {
					openRequest.setRequestHeader(key, headers[key]);
				}
			}
		}
	}
}

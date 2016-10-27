/// <reference path="../definitions/es6-promise/es6-promise.d.ts"/>
/// <reference path="../oneNoteApi.d.ts"/>
/// <reference path="../definitions/content-type/content-type.d.ts"/>

import {ErrorUtils, RequestErrorType} from "./errorUtils";

import * as ContentType from "content-type";

type XHRData = ArrayBufferView | Blob | Document | string | FormData;

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

	private token: string;
	private timeout: number;
	private headers: { [key: string]: string };

	constructor(token: string, timeout: number, headers: { [key: string]: string } = {}) {
		this.token = token;
		this.timeout = timeout;
		this.headers = headers;
	}

	public requestPromise(partialUrl: string, data?: XHRData, contentType?: string, httpMethod?: string): Promise<ResponsePackage<any> | OneNoteApi.RequestError> {
		let fullUrl = this.generateFullUrl(partialUrl);

		if (contentType === null) {
			contentType = "application/json";
		}

		return new Promise(((resolve: (responsePackage: ResponsePackage<any>) => void, reject: (error: OneNoteApi.RequestError) => void) => {
			this.makeRequest(fullUrl, data, contentType, httpMethod).then((responsePackage: ResponsePackage<any>) => {
				resolve(responsePackage);
			}, (error: OneNoteApi.RequestError) => {
				reject(error);
			});
		}));
	}

	private generateFullUrl(partialUrl: string): string {
		let apiRootUrl: string = this.useBetaApi ? "https://www.onenote.com/api/beta/me/notes" : "https://www.onenote.com/api/v1.0/me/notes";
		return apiRootUrl + partialUrl;
	}

	private makeRequest(url: string, data?: XHRData, contentType?: string, httpMethod?: string): Promise<ResponsePackage<any> | OneNoteApi.RequestError> {
		return new Promise((resolve: (responsePackage: ResponsePackage<any>) => void, reject: (error: OneNoteApi.RequestError) => void) => {
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

			request.setRequestHeader("Authorization", "Bearer " + this.token);
			OneNoteApiBase.addHeadersToRequest(request, this.headers);

			request.send(data);
		});
	}

	public static addHeadersToRequest(openRequest: XMLHttpRequest, headers: { [key: string]: string }) {
		if (headers) {
			for (let key in headers) {
				if (headers.hasOwnProperty(key)) {
					openRequest.setRequestHeader(key, headers[key]);
				}
			}
		}
	}
}

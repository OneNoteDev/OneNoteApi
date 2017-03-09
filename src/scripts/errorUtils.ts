/// <reference path="../oneNoteApi.d.ts"/>

export enum RequestErrorType {
	NETWORK_ERROR,
	UNEXPECTED_RESPONSE_STATUS,
	REQUEST_TIMED_OUT,
	UNABLE_TO_PARSE_RESPONSE
}

export class ErrorUtils {
	public static createRequestErrorObject(request: XMLHttpRequest, errorType: RequestErrorType): OneNoteApi.RequestError {
		if (request === undefined) {
			return;
		}

		return ErrorUtils.createRequestErrorObjectInternal(request.status, request.readyState, request.response, request.getAllResponseHeaders(), request.timeout, errorType);
	}

	/**
	 * Split out for unit testing purposes.
	 * Meant only to be called by ErrorUtils.createRequestErrorObject and UTs.
	 */
	public static createRequestErrorObjectInternal(status: number, readyState: number, response: any, responseHeaders: string, timeout: number, errorType: RequestErrorType): OneNoteApi.RequestError {
		let errorMessage: string = ErrorUtils.formatRequestErrorTypeAsString(errorType);
		if (errorType === RequestErrorType.NETWORK_ERROR) {
			errorMessage += ErrorUtils.getAdditionalNetworkErrorInfo(readyState);
		}
		if (errorType === RequestErrorType.REQUEST_TIMED_OUT) {
			status = 408;
		}

		let requestErrorObject: OneNoteApi.RequestError = {
			error: errorMessage,
			statusCode: status,
			responseHeaders: ErrorUtils.convertResponseHeadersToJsonInternal(responseHeaders),
			response: response
		};

		// add the timeout property iff
		// 1) timeout is greater than 0ms
		// 2) status code is not in the 2XX range
		if (timeout > 0 && !(status >= 200 && status < 300)) {
			requestErrorObject.timeout = timeout;
		}

		return requestErrorObject;
	}

	public static convertResponseHeadersToJson(request: XMLHttpRequest): { [key: string]: string } {
		if (request === undefined) {
			return;
		}

		let responseHeaders: string = request.getAllResponseHeaders();
		return ErrorUtils.convertResponseHeadersToJsonInternal(responseHeaders);
	}

	/**
	 * Split out for unit testing purposes.
	 * Meant only to be called by ErrorUtils.convertResponseHeadersToJson, ErrorUtils.createRequestErrorObject, and UTs.
	 */
	public static convertResponseHeadersToJsonInternal(responseHeaders: string): { [key: string]: string } {
		// matches a string of form [key]:[value] or [key]: [value]
		// g = global modifier, returns all matches found in string and not just the first
		let responseHeadersRegex = /([^:]+):\s?(.*)/g;
		let responseHeadersJson: { [key: string]: string } = {};

		let m: RegExpExecArray;
		while (m = responseHeadersRegex.exec(responseHeaders)) {
			if (m.index === responseHeadersRegex.lastIndex) {
				responseHeadersRegex.lastIndex++;
			}

			let headerKey = m[1].trim();
			let headerValue = m[2].trim();

			responseHeadersJson[headerKey] = headerValue;
		}

		return responseHeadersJson;
	}

	private static getAdditionalNetworkErrorInfo(readyState: number): string {
		return ": " + JSON.stringify({ readyState: readyState });
	}

	private static formatRequestErrorTypeAsString(errorType: RequestErrorType): string {
		let errorTypeString = RequestErrorType[errorType];
		return errorTypeString.charAt(0).toUpperCase() + errorTypeString.replace(/_/g, " ").toLowerCase().slice(1);
	}
}

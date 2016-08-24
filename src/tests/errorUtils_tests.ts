/// <reference path="../definitions/qunit/qunit.d.ts" />

import {ErrorUtils, RequestErrorType} from "../scripts/errorUtils";

QUnit.module("requestError", {});

// convertResponseHeadersToJsonInternal

let defaultResponseHeadersJson: { [key: string]: string } = { "X-Powered-By": "ASP.NET", "X-OfficeFE": "OneNoteServiceFrontEndApi_IN_0", "X-CorrelationId": "1f7990e5-8118-408d-a2ea-42a1bd728ac4" };
let defaultResponseHeadersString = "X-Powered-By: ASP.NET\r\n\r\nX-OfficeFE:OneNoteServiceFrontEndApi_IN_0\r\n\r\nX-CorrelationId: 1f7990e5-8118-408d-a2ea-42a1bd728ac4";

test("convertResponseHeadersToJsonInternal should handle a correctly-formatted response headers string containing CRLF", () => {
	let responseHeadersJson: { [key: string]: string } = ErrorUtils.convertResponseHeadersToJsonInternal(defaultResponseHeadersString);
	deepEqual(responseHeadersJson, defaultResponseHeadersJson, "Response headers JSON object is incorrect. Expected: " + JSON.stringify(defaultResponseHeadersJson) + ". Returned: " + JSON.stringify(responseHeadersJson));
});

test("convertResponseHeadersToJsonInternal should handle an incorrectly-formatted string", () => {
	let expectedResponseHeadersJson = {};
	let testResponseHeadersString = "random string";

	let responseHeadersJson: { [key: string]: string } = ErrorUtils.convertResponseHeadersToJsonInternal(testResponseHeadersString);
	deepEqual(responseHeadersJson, expectedResponseHeadersJson, "Response headers JSON object is incorrect. Expected: " + JSON.stringify(expectedResponseHeadersJson) + ". Returned: " + JSON.stringify(responseHeadersJson));
});

// createRequestErrorObjectInternal

let defaultStatus = 42; // non-sensical code to prove that it shouldn't affect method functionality
let defaultReadyState = 0;
let defaultTimeout = 30000;
let defaultResponse = "responded";

test("createRequestErrorObjectInternal returns RequestError of type NETWORK_ERROR", () => {
	let expectedRequestError: OneNoteApi.RequestError = {
		error: "Network error: {\"readyState\":0}",
		statusCode: defaultStatus,
		responseHeaders: defaultResponseHeadersJson,
		response: defaultResponse,
		timeout: defaultTimeout
	};
	let actualRequestError = ErrorUtils.createRequestErrorObjectInternal(defaultStatus, defaultReadyState, defaultResponse, defaultResponseHeadersString, defaultTimeout, RequestErrorType.NETWORK_ERROR);

	deepEqual(expectedRequestError, actualRequestError, "RequestError object is incorrect");
});

test("createRequestErrorObjectInternal returns RequestError of type NETWORK_ERROR", () => {
	let expectedRequestError: OneNoteApi.RequestError = {
		error: "Network error: {\"readyState\":0}",
		statusCode: defaultStatus,
		responseHeaders: defaultResponseHeadersJson,
		response: defaultResponse,
		timeout: defaultTimeout
	};
	let actualRequestError = ErrorUtils.createRequestErrorObjectInternal(defaultStatus, defaultReadyState, defaultResponse, defaultResponseHeadersString, defaultTimeout, RequestErrorType.NETWORK_ERROR);

	deepEqual(expectedRequestError, actualRequestError, "RequestError object is incorrect");
});

test("createRequestErrorObjectInternal returns RequestError of type REQUEST_TIMED_OUT", () => {
	let expectedRequestError: OneNoteApi.RequestError = {
		error: "Request timed out",
		statusCode: 408,
		responseHeaders: defaultResponseHeadersJson,
		response: defaultResponse,
		timeout: defaultTimeout
	};
	let actualRequestError = ErrorUtils.createRequestErrorObjectInternal(defaultStatus, defaultReadyState, defaultResponse, defaultResponseHeadersString, defaultTimeout, RequestErrorType.REQUEST_TIMED_OUT);

	deepEqual(expectedRequestError, actualRequestError, "RequestError object is incorrect");
});

test("createRequestErrorObjectInternal returns RequestError of type UNABLE_TO_PARSE_RESPONSE", () => {
	let expectedRequestError: OneNoteApi.RequestError = {
		error: "Unable to parse response",
		statusCode: defaultStatus,
		responseHeaders: defaultResponseHeadersJson,
		response: defaultResponse,
		timeout: defaultTimeout
	};
	let actualRequestError = ErrorUtils.createRequestErrorObjectInternal(defaultStatus, defaultReadyState, defaultResponse, defaultResponseHeadersString, defaultTimeout, RequestErrorType.UNABLE_TO_PARSE_RESPONSE);

	deepEqual(expectedRequestError, actualRequestError, "RequestError object is incorrect");
});

test("createRequestErrorObjectInternal returns RequestError of type UNEXPECTED_RESPONSE_STATUS", () => {
	let expectedRequestError: OneNoteApi.RequestError = {
		error: "Unexpected response status",
		statusCode: defaultStatus,
		responseHeaders: defaultResponseHeadersJson,
		response: defaultResponse,
		timeout: defaultTimeout
	};
	let actualRequestError = ErrorUtils.createRequestErrorObjectInternal(defaultStatus, defaultReadyState, defaultResponse, defaultResponseHeadersString, defaultTimeout, RequestErrorType.UNEXPECTED_RESPONSE_STATUS);

	deepEqual(expectedRequestError, actualRequestError, "RequestError object is incorrect");
});

test("createRequestErrorObjectInternal returns RequestError without timeout property if undefined", () => {
	let expectedRequestError: OneNoteApi.RequestError = {
		error: "Request timed out",
		statusCode: 408,
		responseHeaders: defaultResponseHeadersJson,
		response: defaultResponse
	};
	let actualRequestError = ErrorUtils.createRequestErrorObjectInternal(defaultStatus, defaultReadyState, defaultResponse, defaultResponseHeadersString, undefined, RequestErrorType.REQUEST_TIMED_OUT);

	deepEqual(expectedRequestError, actualRequestError, "RequestError object is incorrect");
});

test("createRequestErrorObjectInternal returns RequestError without timeout property if 0ms", () => {
	let expectedRequestError: OneNoteApi.RequestError = {
		error: "Request timed out",
		statusCode: 408,
		responseHeaders: defaultResponseHeadersJson,
		response: defaultResponse
	};
	let actualRequestError = ErrorUtils.createRequestErrorObjectInternal(defaultStatus, defaultReadyState, defaultResponse, defaultResponseHeadersString, 0, RequestErrorType.REQUEST_TIMED_OUT);

	deepEqual(expectedRequestError, actualRequestError, "RequestError object is incorrect");
});

test("createRequestErrorObjectInternal returns RequestError without timeout property if status code is in 200-range", () => {
	let statusCodeArg = 200;

	let expectedRequestError: OneNoteApi.RequestError = {
		error: "Unable to parse response",
		statusCode: statusCodeArg,
		responseHeaders: defaultResponseHeadersJson,
		response: defaultResponse
	};
	let actualRequestError = ErrorUtils.createRequestErrorObjectInternal(statusCodeArg, defaultReadyState, defaultResponse, defaultResponseHeadersString, defaultTimeout, RequestErrorType.UNABLE_TO_PARSE_RESPONSE);

	deepEqual(expectedRequestError, actualRequestError, "RequestError object is incorrect for status code " + statusCodeArg);

	statusCodeArg = 299;

	expectedRequestError = {
		error: "Unable to parse response",
		statusCode: statusCodeArg,
		responseHeaders: defaultResponseHeadersJson,
		response: defaultResponse
	};
	actualRequestError = ErrorUtils.createRequestErrorObjectInternal(statusCodeArg, defaultReadyState, defaultResponse, defaultResponseHeadersString, defaultTimeout, RequestErrorType.UNABLE_TO_PARSE_RESPONSE);

	deepEqual(expectedRequestError, actualRequestError, "RequestError object is incorrect for status code " + statusCodeArg);
});

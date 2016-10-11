/*
The MIT License (MIT)

Copyright (c) 2015 OneNoteDev

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/(function(f){if(typeof exports==="object"&&typeof module!=="undefined"){module.exports=f()}else if(typeof define==="function"&&define.amd){define([],f)}else{var g;if(typeof window!=="undefined"){g=window}else if(typeof global!=="undefined"){g=global}else if(typeof self!=="undefined"){g=self}else{g=this}g.OneNoteApi = f()}})(function(){var define,module,exports;return (function e(t,n,r){function s(o,u){if(!n[o]){if(!t[o]){var a=typeof require=="function"&&require;if(!u&&a)return a(o,!0);if(i)return i(o,!0);var f=new Error("Cannot find module '"+o+"'");throw f.code="MODULE_NOT_FOUND",f}var l=n[o]={exports:{}};t[o][0].call(l.exports,function(e){var n=t[o][1][e];return s(n?n:e)},l,l.exports,e,t,n,r)}return n[o].exports}var i=typeof require=="function"&&require;for(var o=0;o<r.length;o++)s(r[o]);return s})({1:[function(require,module,exports){
/*--------------------------------------------------------------------------
   Definitions of the supported ContentTypes.
 -------------------------------------------------------------------------*/
"use strict";
(function (ContentType) {
    ContentType[ContentType["Html"] = 0] = "Html";
    ContentType[ContentType["Image"] = 1] = "Image";
    ContentType[ContentType["EnhancedUrl"] = 2] = "EnhancedUrl";
    ContentType[ContentType["Url"] = 3] = "Url";
    ContentType[ContentType["Onml"] = 4] = "Onml";
})(exports.ContentType || (exports.ContentType = {}));
var ContentType = exports.ContentType;

},{}],2:[function(require,module,exports){
/// <reference path="../oneNoteApi.d.ts"/>
"use strict";
(function (RequestErrorType) {
    RequestErrorType[RequestErrorType["NETWORK_ERROR"] = 0] = "NETWORK_ERROR";
    RequestErrorType[RequestErrorType["UNEXPECTED_RESPONSE_STATUS"] = 1] = "UNEXPECTED_RESPONSE_STATUS";
    RequestErrorType[RequestErrorType["REQUEST_TIMED_OUT"] = 2] = "REQUEST_TIMED_OUT";
    RequestErrorType[RequestErrorType["UNABLE_TO_PARSE_RESPONSE"] = 3] = "UNABLE_TO_PARSE_RESPONSE";
})(exports.RequestErrorType || (exports.RequestErrorType = {}));
var RequestErrorType = exports.RequestErrorType;
var ErrorUtils = (function () {
    function ErrorUtils() {
    }
    ErrorUtils.createRequestErrorObject = function (request, errorType) {
        if (request === undefined || request === null) {
            return;
        }
        return ErrorUtils.createRequestErrorObjectInternal(request.status, request.readyState, request.response, request.getAllResponseHeaders(), request.timeout, errorType);
    };
    /**
     * Split out for unit testing purposes.
     * Meant only to be called by ErrorUtils.createRequestErrorObject and UTs.
     */
    ErrorUtils.createRequestErrorObjectInternal = function (status, readyState, response, responseHeaders, timeout, errorType) {
        var errorMessage = ErrorUtils.formatRequestErrorTypeAsString(errorType);
        if (errorType === RequestErrorType.NETWORK_ERROR) {
            errorMessage += ErrorUtils.getAdditionalNetworkErrorInfo(readyState);
        }
        if (errorType === RequestErrorType.REQUEST_TIMED_OUT) {
            status = 408;
        }
        var requestErrorObject = {
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
    };
    ErrorUtils.convertResponseHeadersToJson = function (request) {
        if (request === undefined || request === null) {
            return;
        }
        var responseHeaders = request.getAllResponseHeaders();
        return ErrorUtils.convertResponseHeadersToJsonInternal(responseHeaders);
    };
    /**
     * Split out for unit testing purposes.
     * Meant only to be called by ErrorUtils.convertResponseHeadersToJson, ErrorUtils.createRequestErrorObject, and UTs.
     */
    ErrorUtils.convertResponseHeadersToJsonInternal = function (responseHeaders) {
        // matches a string of form [key]:[value] or [key]: [value]
        // g = global modifier, returns all matches found in string and not just the first
        var responseHeadersRegex = /([^:]+):\s?(.*)/g;
        var responseHeadersJson = {};
        var m;
        while (m = responseHeadersRegex.exec(responseHeaders)) {
            if (m.index === responseHeadersRegex.lastIndex) {
                responseHeadersRegex.lastIndex++;
            }
            var headerKey = m[1].trim();
            var headerValue = m[2].trim();
            responseHeadersJson[headerKey] = headerValue;
        }
        return responseHeadersJson;
    };
    ErrorUtils.getAdditionalNetworkErrorInfo = function (readyState) {
        return ": " + JSON.stringify({ readyState: readyState });
    };
    ErrorUtils.formatRequestErrorTypeAsString = function (errorType) {
        var errorTypeString = RequestErrorType[errorType];
        return errorTypeString.charAt(0).toUpperCase() + errorTypeString.replace(/_/g, " ").toLowerCase().slice(1);
    };
    return ErrorUtils;
}());
exports.ErrorUtils = ErrorUtils;

},{}],3:[function(require,module,exports){
/// <reference path="../oneNoteApi.d.ts"/>
"use strict";
var NotebookUtils = (function () {
    function NotebookUtils() {
    }
    /**
     * Checks to see if the section exists in the notebook list.
     *
     * @param notebooks List of notebooks to search
     * @param sectionId Section id to check the existence of
     * @return true if the section exists in the notebooks; false otherwise
     */
    NotebookUtils.sectionExistsInNotebooks = function (notebooks, sectionId) {
        if (!notebooks || !sectionId) {
            return false;
        }
        for (var i = 0; i < notebooks.length; i++) {
            if (NotebookUtils.sectionExistsInParent(notebooks[i], sectionId)) {
                return true;
            }
        }
        return false;
    };
    /**
     * Checks to see if the section exists in the notebook or section group.
     *
     * @param parent Notebook or section group to search
     * @param sectionId Section id to check the existence of
     * @return true if the section exists in the parent; false otherwise
     */
    NotebookUtils.sectionExistsInParent = function (parent, sectionId) {
        if (!parent || !sectionId) {
            return false;
        }
        if (parent.sections) {
            for (var i = 0; i < parent.sections.length; i++) {
                var section = parent.sections[i];
                if (section && section.id === sectionId) {
                    return true;
                }
            }
        }
        if (parent.sectionGroups) {
            for (var i = 0; i < parent.sectionGroups.length; i++) {
                if (NotebookUtils.sectionExistsInParent(parent.sectionGroups[i], sectionId)) {
                    return true;
                }
            }
        }
        return false;
    };
    /**
     * Retrieves the path starting from the notebook to the first ancestor section found that
     * meets a given criteria.
     *
     * @param notebooks List of notebooks to search
     * @return section path (e.g., [notebook, sectionGroup, section]); undefined if there is none
     */
    NotebookUtils.getPathFromNotebooksToSection = function (notebooks, filter) {
        if (!notebooks || !filter) {
            return undefined;
        }
        for (var i = 0; i < notebooks.length; i++) {
            var notebook = notebooks[i];
            var notebookSearchResult = NotebookUtils.getPathFromParentToSection(notebook, filter);
            if (notebookSearchResult) {
                return notebookSearchResult;
            }
        }
        return undefined;
    };
    /**
     * Recursively retrieves the path starting from the specified parent to the first ancestor
     * section found that meets a given criteria.
     *
     * @param parent The notebook or section group to search
     * @return section path (e.g., [parent, sectionGroup, sectionGroup, section]); undefined if there is none
     */
    NotebookUtils.getPathFromParentToSection = function (parent, filter) {
        if (!parent || !filter) {
            return undefined;
        }
        if (parent.sections) {
            for (var i = 0; i < parent.sections.length; i++) {
                var section = parent.sections[i];
                if (filter(section)) {
                    return [parent, section];
                }
            }
        }
        if (parent.sectionGroups) {
            for (var i = 0; i < parent.sectionGroups.length; i++) {
                var sectionGroup = parent.sectionGroups[i];
                var sectionGroupSearchResult = NotebookUtils.getPathFromParentToSection(sectionGroup, filter);
                if (sectionGroupSearchResult) {
                    sectionGroupSearchResult.unshift(parent);
                    return sectionGroupSearchResult;
                }
            }
        }
        return undefined;
    };
    /**
     * Computes the maximum depth of the notebooks list, including sections.
     *
     * @param notebooks List of notebooks
     * @return Maximum depth
     */
    NotebookUtils.getDepthOfNotebooks = function (notebooks) {
        if (!notebooks || notebooks.length === 0) {
            return 0;
        }
        return notebooks.map(function (notebook) { return NotebookUtils.getDepthOfParent(notebook); }).reduce(function (d1, d2) { return Math.max(d1, d2); });
    };
    /**
     * Computes the maximum depth of the non-section parent entity, including sections.
     *
     * @param notebooks Non-section parent entity
     * @return Maximum depth
     */
    NotebookUtils.getDepthOfParent = function (parent) {
        if (!parent) {
            return 0;
        }
        var containsAtLeastOneSection = parent.sections && parent.sections.length > 0;
        var maxDepth = containsAtLeastOneSection ? 1 : 0;
        if (parent.sectionGroups) {
            for (var i = 0; i < parent.sectionGroups.length; i++) {
                maxDepth = Math.max(NotebookUtils.getDepthOfParent(parent.sectionGroups[i]), maxDepth);
            }
        }
        // Include the parent itself
        return maxDepth + 1;
    };
    return NotebookUtils;
}());
exports.NotebookUtils = NotebookUtils;

},{}],4:[function(require,module,exports){
"use strict";
var __extends = (this && this.__extends) || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
};
var oneNoteApiBase_1 = require("./oneNoteApiBase");
/**
* Wrapper for easier calling of the OneNote APIs.
*/
var OneNoteApi = (function (_super) {
    __extends(OneNoteApi, _super);
    function OneNoteApi(token, timeout, headers) {
        if (timeout === void 0) { timeout = 30000; }
        if (headers === void 0) { headers = {}; }
        _super.call(this, token, timeout, headers);
    }
    /**
    * CreateNotebook
    */
    OneNoteApi.prototype.createNotebook = function (name) {
        var data = JSON.stringify({ name: name });
        return this.requestPromise(this.getNotebooksUrl(), data);
    };
    /**
    * CreatePage
    */
    OneNoteApi.prototype.createPage = function (page, sectionId) {
        var sectionPath = sectionId ? "/sections/" + sectionId : "";
        var url = sectionPath + "/pages";
        var form = page.getTypedFormData();
        return this.requestPromise(url, form.asBlob(), form.getContentType());
    };
    /**
     * UpdatePage
     */
    OneNoteApi.prototype.updatePage = function (pageId, revisions) {
        var pagePath = "/pages/" + pageId;
        var url = pagePath + "/content";
        return this.requestPromise(url, revisions, "application/json", "PATCH");
    };
    /**
    * CreateSection
    */
    OneNoteApi.prototype.createSection = function (notebookId, name) {
        var obj = { name: name };
        var data = JSON.stringify(obj);
        return this.requestPromise("/notebooks/" + notebookId + "/sections/", data);
    };
    /**
    * GetNotebooks
    */
    OneNoteApi.prototype.getNotebooks = function (excludeReadOnlyNotebooks) {
        if (excludeReadOnlyNotebooks === void 0) { excludeReadOnlyNotebooks = true; }
        return this.requestPromise(this.getNotebooksUrl(null /*expands*/, excludeReadOnlyNotebooks));
    };
    /**
    * GetNotebooksWithExpandedSections
    */
    OneNoteApi.prototype.getNotebooksWithExpandedSections = function (expands, excludeReadOnlyNotebooks) {
        if (expands === void 0) { expands = 2; }
        if (excludeReadOnlyNotebooks === void 0) { excludeReadOnlyNotebooks = true; }
        return this.requestPromise(this.getNotebooksUrl(expands, excludeReadOnlyNotebooks));
    };
    /**
    * GetNotebookbyName
    */
    OneNoteApi.prototype.getNotebookByName = function (name) {
        return this.requestPromise("/notebooks?filter=name%20eq%20%27" + encodeURI(name) + "%27&orderby=createdTime");
    };
    /**
    * PagesSearch
    */
    OneNoteApi.prototype.pagesSearch = function (query) {
        return this.requestPromise(this.getSearchUrl(query));
    };
    /**
    * GetExpands
    *
    * Nest expands so we can get notebook elements (sections and section groups) in
    * the same call that we get notebooks.
    *
    * expands specifies how many levels deep to return.
    */
    OneNoteApi.prototype.getExpands = function (expands) {
        if (expands <= 0) {
            return "";
        }
        var s = "$expand=sections,sectionGroups";
        return expands === 1 ? s : s + "(" + this.getExpands(expands - 1) + ")";
    };
    /**
    * GetNotebooksUrl
    */
    OneNoteApi.prototype.getNotebooksUrl = function (numExpands, excludeReadOnlyNotebooks) {
        if (numExpands === void 0) { numExpands = 0; }
        if (excludeReadOnlyNotebooks === void 0) { excludeReadOnlyNotebooks = true; }
        // Since this url is most often used to save content to a specific notebook, by default
        // it does not include a notebook where user has Read only permissions.
        var filter = (excludeReadOnlyNotebooks) ? "$filter=userRole%20ne%20Microsoft.OneNote.Api.UserRole'Reader'" : "";
        return "/notebooks?" + filter + (numExpands ? "&" + this.getExpands(numExpands) : "");
    };
    /**
    * GetSearchUrl
    */
    OneNoteApi.prototype.getSearchUrl = function (query) {
        return "/pages?search=" + query;
    };
    return OneNoteApi;
}(oneNoteApiBase_1.OneNoteApiBase));
exports.OneNoteApi = OneNoteApi;
var contentType_1 = require("./contentType");
exports.ContentType = contentType_1.ContentType;
var oneNotePage_1 = require("./oneNotePage");
exports.OneNotePage = oneNotePage_1.OneNotePage;
var errorUtils_1 = require("./errorUtils");
exports.ErrorUtils = errorUtils_1.ErrorUtils;
exports.RequestErrorType = errorUtils_1.RequestErrorType;
var notebookUtils_1 = require("./notebookUtils");
exports.NotebookUtils = notebookUtils_1.NotebookUtils;

},{"./contentType":1,"./errorUtils":2,"./notebookUtils":3,"./oneNoteApiBase":5,"./oneNotePage":6}],5:[function(require,module,exports){
/// <reference path="../definitions/es6-promise/es6-promise.d.ts"/>
/// <reference path="../oneNoteApi.d.ts"/>
"use strict";
var errorUtils_1 = require("./errorUtils");
/**
* Base communication layer for talking to the OneNote APIs.
*/
var OneNoteApiBase = (function () {
    function OneNoteApiBase(token, timeout, headers) {
        if (headers === void 0) { headers = {}; }
        // Whether or not the OneNote Beta APIs should be used.
        this.useBetaApi = false;
        this.token = token;
        this.timeout = timeout;
        this.headers = headers;
    }
    OneNoteApiBase.prototype.requestPromise = function (partialUrl, data, contentType, verb) {
        var _this = this;
        var fullUrl = this.generateFullUrl(partialUrl);
        if (contentType === null) {
            contentType = "application/json";
        }
        return new Promise((function (resolve, reject) {
            _this.makeRequest(fullUrl, data, contentType, verb).then(function (responsePackage) {
                resolve(responsePackage);
            }, function (error) {
                reject(error);
            });
        }));
    };
    OneNoteApiBase.prototype.generateFullUrl = function (partialUrl) {
        var apiRootUrl = this.useBetaApi ? "https://www.onenote.com/api/beta/me/notes" : "https://www.onenote.com/api/v1.0/me/notes";
        return apiRootUrl + partialUrl;
    };
    OneNoteApiBase.prototype.makeRequest = function (url, data, contentType, verb) {
        var _this = this;
        return new Promise(function (resolve, reject) {
            var request = new XMLHttpRequest();
            var type = verb ? verb : data ? "POST" : "GET";
            request.open(type, url);
            request.timeout = _this.timeout;
            request.onload = function () {
                // TODO: more status code checking
                if (request.status === 200 || request.status === 201) {
                    try {
                        var parsedResponse = JSON.parse(request.response);
                        var responsePackage = { parsedResponse: parsedResponse, request: request };
                        resolve(responsePackage);
                    }
                    catch (e) {
                        reject(errorUtils_1.ErrorUtils.createRequestErrorObject(request, errorUtils_1.RequestErrorType.UNABLE_TO_PARSE_RESPONSE));
                    }
                }
                else {
                    reject(errorUtils_1.ErrorUtils.createRequestErrorObject(request, errorUtils_1.RequestErrorType.UNEXPECTED_RESPONSE_STATUS));
                }
            };
            request.onerror = function () {
                reject(errorUtils_1.ErrorUtils.createRequestErrorObject(request, errorUtils_1.RequestErrorType.NETWORK_ERROR));
            };
            request.ontimeout = function () {
                reject(errorUtils_1.ErrorUtils.createRequestErrorObject(request, errorUtils_1.RequestErrorType.REQUEST_TIMED_OUT));
            };
            if (contentType) {
                request.setRequestHeader("Content-Type", contentType);
            }
            request.setRequestHeader("Authorization", "Bearer " + _this.token);
            OneNoteApiBase.addHeadersToRequest(request, _this.headers);
            request.send(data);
        });
    };
    OneNoteApiBase.addHeadersToRequest = function (openRequest, headers) {
        if (headers) {
            for (var key in headers) {
                if (headers.hasOwnProperty(key)) {
                    openRequest.setRequestHeader(key, headers[key]);
                }
            }
        }
    };
    return OneNoteApiBase;
}());
exports.OneNoteApiBase = OneNoteApiBase;

},{"./errorUtils":2}],6:[function(require,module,exports){
"use strict";
var typedFormData_1 = require("./typedFormData");
/**
 * The page payload consists of multiple data parts. The first data part is the 'Presentation'
 * that describes the elements that render on the page itself as ONML. Subsequent data parts are referenced
 * by the Presentation, e.g. binary data.
 */
var OneNotePage = (function () {
    function OneNotePage(title, presentationBody, locale, pageMetadata) {
        if (title === void 0) { title = ""; }
        if (presentationBody === void 0) { presentationBody = ""; }
        if (locale === void 0) { locale = "en-us"; }
        if (pageMetadata === void 0) { pageMetadata = undefined; }
        // Non-'Presentation' data parts
        this.dataParts = [];
        this.title = title;
        this.presentationBody = presentationBody;
        this.locale = locale;
        this.pageMetadata = pageMetadata;
    }
    /**
     * Includes everything inside and including the <html> tags
     */
    OneNotePage.prototype.getEntireOnml = function () {
        return '<html xmlns="http://www.w3.org/1999/xhtml" lang=' + this.locale + ">" +
            this.getHeadAsHtml() +
            "<body>" + this.presentationBody + "</body></html>";
    };
    OneNotePage.prototype.getHeadAsHtml = function () {
        var pageMetadataHtml = this.getPageMetadataAsHtml();
        var createTime = this.formUtcOffsetString(new Date());
        return "<head><title>" + this.escapeHtmlEntities(this.title) + "</title>" +
            '<meta name="created" content="' + createTime + ' ">' + pageMetadataHtml + "</head>";
    };
    OneNotePage.prototype.getPageMetadataAsHtml = function () {
        var pageMetadataHtml = "";
        if (this.pageMetadata) {
            for (var key in this.pageMetadata) {
                pageMetadataHtml += '<meta name="' + key + '" content="' + this.escapeHtmlEntities(this.pageMetadata[key]) + '" />';
            }
        }
        return pageMetadataHtml;
    };
    OneNotePage.prototype.formUtcOffsetString = function (date) {
        var offset = date.getTimezoneOffset();
        var sign = offset >= 0 ? "-" : "+";
        offset = Math.abs(offset);
        var hours = Math.floor(offset / 60) + "";
        var mins = Math.round(offset % 60) + "";
        if (parseInt(hours, 10) < 10) {
            hours = "0" + hours;
        }
        if (parseInt(mins, 10) < 10) {
            mins = "0" + mins;
        }
        return sign + hours + ":" + mins;
    };
    OneNotePage.prototype.generateMimePartName = function (prefix) {
        return prefix + Math.floor(Math.random() * 10000).toString();
    };
    OneNotePage.prototype.escapeHtmlEntities = function (value) {
        var divElement = document.createElement("div");
        divElement.innerText = value;
        return divElement.innerHTML;
    };
    /**
     * Converts the Presentation and subsequent data parts entirely into data parts
     */
    OneNotePage.prototype.getTypedFormData = function () {
        var tfd = new typedFormData_1.TypedFormData();
        tfd.append("Presentation", this.getEntireOnml(), "application/xhtml+xml");
        for (var i = 0; i < this.dataParts.length; i++) {
            var part = this.dataParts[i];
            tfd.append(part.name, part.content, part.type);
        }
        return tfd;
    };
    OneNotePage.prototype.addOnml = function (onml) {
        this.presentationBody += onml;
    };
    OneNotePage.prototype.addHtml = function (html) {
        var mimeName = this.generateMimePartName("Html");
        this.dataParts.push({
            content: html,
            name: mimeName,
            type: "text/HTML"
        });
        this.addOnml('<img data-render-src="name:' + mimeName + '"/>');
        return mimeName;
    };
    OneNotePage.prototype.addImage = function (imgUrl) {
        this.addOnml('<img src="' + imgUrl + '"/>');
    };
    /**
     * The input can either be a url, or a reference to a MIME part containing binary
     * e.g., "name:REFERENCE"
     */
    OneNotePage.prototype.addObjectUrlAsImage = function (url) {
        this.addOnml('<img data-render-src="' + url + '"/>');
    };
    OneNotePage.prototype.addAttachment = function (binary, name) {
        // We hardcode the name as "Attachment" as each page is only allowed to have one
        var mimeName = this.generateMimePartName("Attachment");
        this.dataParts.push({
            content: binary,
            name: mimeName,
            type: "application/pdf"
        });
        this.addOnml('<object data-attachment="' + name + '" data="name:' + mimeName + '" type="application/pdf" />');
        return mimeName;
    };
    OneNotePage.prototype.addUrl = function (url) {
        this.addOnml('<div data-render-src="' + url + '" data-render-method="extract" data-render-fallback="none"></div>');
    };
    OneNotePage.prototype.addCitation = function (format, urlToDisplay, rawUrl) {
        this.addOnml(format.replace("{0}", '<a href="' + (rawUrl ? rawUrl : urlToDisplay) + '">' + urlToDisplay + "</a>"));
    };
    return OneNotePage;
}());
exports.OneNotePage = OneNotePage;

},{"./typedFormData":7}],7:[function(require,module,exports){
"use strict";
// A substitute for FormData that allows the ability to set content-type per part
var TypedFormData = (function () {
    function TypedFormData() {
        /*const*/
        this.contentTypeMultipart = "multipart/form-data; boundary=";
        this.dataParts = [];
        this.boundaryName = "OneNoteTypedDataBoundary" + Math.floor((Math.random() * 1000));
    }
    TypedFormData.prototype.getContentType = function () {
        return this.contentTypeMultipart + this.boundaryName;
    };
    TypedFormData.prototype.append = function (name, content, type) {
        this.dataParts.push({ content: content, name: name, type: type });
    };
    TypedFormData.prototype.asBlob = function () {
        var boundary = "--" + this.boundaryName;
        var payload = [boundary];
        for (var i = 0; i < this.dataParts.length; i++) {
            var curr = this.dataParts[i];
            var header = "\r\n" +
                "Content-Type: " + curr.type + "\r\n" +
                "Content-Disposition: form-data; name=\"" + curr.name + "\"\r\n\r\n";
            payload.push(header);
            payload.push(curr.content);
            payload.push("\r\n" + boundary);
        }
        payload.push("--\r\n");
        return new Blob(payload);
    };
    return TypedFormData;
}());
exports.TypedFormData = TypedFormData;

},{}]},{},[4])(4)
});
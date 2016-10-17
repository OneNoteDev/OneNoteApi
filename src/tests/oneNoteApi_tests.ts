/// <reference path="../definitions/qunit/qunit.d.ts" />
/// <reference path="../definitions/sinon/sinon.d.ts" />

import {OneNoteApi} from "../scripts/oneNoteApi";
import {ResponsePackage} from "../scripts/oneNoteApiBase";

let xhr: Sinon.SinonFakeXMLHttpRequest;
let server: Sinon.SinonFakeServer;

QUnit.module("oneNoteApi", {
	beforeEach: () => {
		xhr = sinon.useFakeXMLHttpRequest();
		let requests = this.requests = [];
		xhr.onCreate = req => {
			requests.push(req);
		};

		server = sinon.fakeServer.create();
		server.respondImmediately = true;
	},
	afterEach: () => {
		xhr.restore();
		server.restore();
	}
});

/**
 * POST Notebooks
 */

test("createNotebook should bubble up the notebook's metadeta in the response", (assert: QUnitAssert) => {
	let done = assert.async();

	let notebookName = "Hello world";
	let responseJson = {
		"@odata.context": "https://www.onenote.com/api/v1.0/$metadata#me/notes/notebooks/$entity",
		"isDefault": false,
		"userRole": "Contributor",
		"isShared": false,
		"sectionsUrl": "https://www.onenote.com/api/v1.0/notebooks/notebook ID/sections",
		"sectionGroupsUrl": "https://www.onenote.com/api/v1.0/notebooks/notebook ID/sectionGroups",
		"links": {
			"oneNoteClientUrl": {
				"href": "https:{client URL}"
			},
			"oneNoteWebUrl": {
				"href": "https://{web URL}"
			}
		},
		"id": "notebook ID",
		"name": notebookName,
		"self": "https://www.onenote.com/api/v1.0/notebooks/notebook ID",
		"createdBy": "user name",
		"lastModifiedBy": "user name",
		"createdTime": "2013-10-05T10:57:00.683Z",
		"lastModifiedTime": "2014-01-28T18:49:00.47Z"
	};
	server.respondWith([200, {}, JSON.stringify(responseJson)]);

	let api = new OneNoteApi("token");
	api.createNotebook(notebookName).then((responsePackage: ResponsePackage<any>) => {
		deepEqual(responsePackage.parsedResponse, responseJson,
			"The parsed response should be the server response as a JSON object");
		ok(responsePackage.request, "The request object should be non-undefined");
	}, (error) => {
		ok(false, "reject should not be called");
	}).then(() => {
		done();
	});
});

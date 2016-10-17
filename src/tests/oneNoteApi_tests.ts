/// <reference path="../definitions/qunit/qunit.d.ts" />
/// <reference path="../definitions/sinon/sinon.d.ts" />

import {OneNoteApi} from "../scripts/oneNoteApi";

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

test("this test should succeed", () => {
	ok(true);
});

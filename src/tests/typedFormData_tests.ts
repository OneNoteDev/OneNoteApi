/// <reference path="../definitions/qunit/qunit.d.ts" />

import {TypedFormData} from "../scripts/typedFormData";

QUnit.module("typedFormData", {});

test("The blob payload should generate just the boundary by default", (assert: QUnitAssert) => {
	let done = assert.async();

	let data: TypedFormData = new TypedFormData();

	let reader = new FileReader();
	reader.onloadend = () => {
		let blobAsText = reader.result;
		ok(blobAsText.match(/--OneNoteTypedDataBoundary\d{0,4}--\r\n/g),
			"The default payload should just contain the boundary");
		done();
	};
	reader.readAsText(data.asBlob());
});

test("The blob payload should generate the data part when it is added", (assert: QUnitAssert) => {
	let done = assert.async();

	let data: TypedFormData = new TypedFormData();
	data.append("name", "content", "type");

	let reader = new FileReader();
	reader.onloadend = () => {
		let blobAsText = reader.result;
		ok(blobAsText.match(/--OneNoteTypedDataBoundary\d{0,4}\r\nContent-Type: type\r\nContent-Disposition: form-data; name="name"\r\n\r\ncontent\r\n--OneNoteTypedDataBoundary\d{0,4}--\r\n/g),
			"The payload should contain one data part");
		done();
	};
	reader.readAsText(data.asBlob());
});

test("The blob payload should generate several data parts when they are added", (assert: QUnitAssert) => {
	let done = assert.async();

	let data: TypedFormData = new TypedFormData();
	data.append("name", "content", "type");
	data.append("name1", "content1", "type1");

	let reader = new FileReader();
	reader.onloadend = () => {
		let blobAsText = reader.result;
		ok(blobAsText.match(/--OneNoteTypedDataBoundary\d{0,4}\r\nContent-Type: type\r\nContent-Disposition: form-data; name="name"\r\n\r\ncontent\r\n--OneNoteTypedDataBoundary\d{0,4}\r\nContent-Type: type1\r\nContent-Disposition: form-data; name="name1"\r\n\r\ncontent1\r\n/g),
			"The payload should contain two data parts");
		done();
	};
	reader.readAsText(data.asBlob());
});

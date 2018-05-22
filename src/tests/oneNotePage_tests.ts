import {OneNotePage} from "../scripts/oneNotePage";
import {QUnit, test, QUnitAssert, ok, equal} from "qunitjs";

let title = "TITLE";
let body = "BODY";

QUnit.module("oneNotePage", {});

test("Page data is present in the blob payload", (assert: QUnitAssert) => {
	let done = assert.async();

	let page = new OneNotePage(title, body);
	let blobPayload = page.getTypedFormData().asBlob();

	let reader = new FileReader();
	reader.onloadend = () => {
		let blobAsText = reader.result;
		ok(blobAsText.indexOf(body) >= 0, "The body is present");
		ok(blobAsText.indexOf(title) >= 0, "title present");
		ok(blobAsText.indexOf("en-us") >= 0, "en-us defaulted");
		ok(blobAsText.indexOf("ja-jp") === -1, "ja-jp not defaulted");
		done();
	};
	reader.readAsText(blobPayload);
});

test("The blob payload is formatted according to what OneNote API's POST docs expect", (assert: QUnitAssert) => {
	let done = assert.async();

	let page = new OneNotePage(title, body);
	let blobPayload = page.getTypedFormData().asBlob();

	let reader = new FileReader();
	reader.onloadend = () => {
		let blobAsText = reader.result;
		ok(blobAsText.match('--OneNoteTypedDataBoundary\\d{0,4}\\r\\nContent-Type: application\\\/xhtml\\+xml\\r\\nContent-Disposition: form-data; name=\"Presentation\"\\r\\n\\r\\n<html xmlns="http:\\\/\\\/www.w3.org\\\/1999\\\/xhtml" lang=en-us><head><title>TITLE<\\\/title><meta name="created" content="[\\-\\+]\\d{1,2}:\\d{1,2} "><\\\/head><body>BODY<\\\/body><\\\/html>\\r\\n--OneNoteTypedDataBoundary\\d{0,4}--\\r\\n'),
			"Title and body inserted correctly");
		done();
	};
	reader.readAsText(blobPayload);
});

test("The blob payload is formatted according to what OneNote API's POST docs expect after AddOnml", (assert: QUnitAssert) => {
	let done = assert.async();

	let page = new OneNotePage(title, body);
	page.addOnml("foo");
	let blobPayload = page.getTypedFormData().asBlob();

	let reader = new FileReader();
	reader.onloadend = () => {
		let blobAsText = reader.result;
		ok(blobAsText.match('--OneNoteTypedDataBoundary\\d{0,4}\\r\\nContent-Type: application\\\/xhtml\\+xml\\r\\nContent-Disposition: form-data; name=\"Presentation\"\\r\\n\\r\\n<html xmlns="http:\\\/\\\/www.w3.org\\\/1999\\\/xhtml" lang=en-us><head><title>TITLE<\\\/title><meta name="created" content="[\\-\\+]\\d{1,2}:\\d{1,2} "><\\\/head><body>BODYfoo<\\\/body><\\\/html>\\r\\n--OneNoteTypedDataBoundary\\d{0,4}--\\r\\n'),
			"AddOnml should simply append the given html into the body");
		done();
	};
	reader.readAsText(blobPayload);
});

test("The blob payload is formatted according to what OneNote API's POST docs expect after AddHtml", (assert: QUnitAssert) => {
	let done = assert.async();

	let page = new OneNotePage(title, body);
	page.addHtml("<div></div>");
	let blobPayload = page.getTypedFormData().asBlob();

	let reader = new FileReader();
	reader.onloadend = () => {
		let blobAsText = reader.result;
		ok(blobAsText.match('--OneNoteTypedDataBoundary\\d{0,4}\\r\\nContent-Type: application\\\/xhtml\\+xml\\r\\nContent-Disposition: form-data; name="Presentation"\\r\\n\\r\\n<html xmlns="http:\\\/\\\/www.w3.org\\\/1999\\\/xhtml" lang=en-us><head><title>TITLE<\\\/title><meta name="created" content="[\\-\\+]\\d{1,2}:\\d{1,2} "><\\\/head><body>BODY<img data-render-src="name:Html\\d{0,4}"\\\/><\\\/body><\\\/html>\\r\\n--OneNoteTypedDataBoundary\\d{0,4}\\r\\nContent-Type: text\\\/HTML\\r\\nContent-Disposition: form-data; name="Html\\d{0,4}"\\r\\n\\r\\n<div><\\\/div>\\r\\n--OneNoteTypedDataBoundary\\d{0,4}--\\r\\n'),
			"AddHtml should add the html as an img attachment which references the data part containing the html");
		done();
	};
	reader.readAsText(blobPayload);
});

test("The blob payload is formatted according to what OneNote API's POST docs expect after AddImage", (assert: QUnitAssert) => {
	let done = assert.async();

	let page = new OneNotePage(title, body);
	page.addImage("url");
	let blobPayload = page.getTypedFormData().asBlob();

	let reader = new FileReader();
	reader.onloadend = () => {
		let blobAsText = reader.result;
		ok(blobAsText.match('--OneNoteTypedDataBoundary\\d{0,4}\\r\\nContent-Type: application\\\/xhtml\\+xml\\r\\nContent-Disposition: form-data; name=\"Presentation\"\\r\\n\\r\\n<html xmlns="http:\\\/\\\/www.w3.org\\\/1999\\\/xhtml" lang=en-us><head><title>TITLE<\\\/title><meta name="created" content="[\\-\\+]\\d{1,2}:\\d{1,2} "><\\\/head><body>BODY<img src="url"\\\/><\\\/body><\\\/html>\\r\\n--OneNoteTypedDataBoundary\\d{0,4}--\\r\\n'),
			"AddUrl should simply append img tags to the body");
		done();
	};
	reader.readAsText(blobPayload);
});

test("The blob payload is formatted according to what OneNote API's POST docs expect after AddObjectUrlAsImage", (assert: QUnitAssert) => {
	let done = assert.async();

	let page = new OneNotePage(title, body);
	page.addObjectUrlAsImage("url");
	let blobPayload = page.getTypedFormData().asBlob();

	let reader = new FileReader();
	reader.onloadend = () => {
		let blobAsText = reader.result;
		ok(blobAsText.match('--OneNoteTypedDataBoundary\\d{0,4}\\r\\nContent-Type: application\\\/xhtml\\+xml\\r\\nContent-Disposition: form-data; name=\"Presentation\"\\r\\n\\r\\n<html xmlns="http:\\\/\\\/www.w3.org\\\/1999\\\/xhtml" lang=en-us><head><title>TITLE<\\\/title><meta name="created" content="[\\-\\+]\\d{1,2}:\\d{1,2} "><\\\/head><body>BODY<img data-render-src="url"\\\/><\\\/body><\\\/html>\\r\\n--OneNoteTypedDataBoundary\\d{0,4}--\\r\\n'),
			"AddObjectUrlAsImage should simply append img tags to the body");
		done();
	};
	reader.readAsText(blobPayload);
});

test("The blob payload is formatted according to what OneNote API's POST docs expect after AddAttachment", (assert: QUnitAssert) => {
	let done = assert.async();

	let page = new OneNotePage(title, body);
	page.addAttachment(new ArrayBuffer(8), "MYATTACHMENT");
	let blobPayload = page.getTypedFormData().asBlob();

	let reader = new FileReader();
	reader.onloadend = () => {
		let blobAsText = reader.result;
		ok(blobAsText.match('--OneNoteTypedDataBoundary\\d{0,4}\\r\\nContent-Type: application\\\/xhtml\\+xml\\r\\nContent-Disposition: form-data; name="Presentation"\\r\\n\\r\\n<html xmlns="http:\\\/\\\/www.w3.org\\\/1999\\\/xhtml" lang=en-us><head><title>TITLE<\\\/title><meta name="created" content="[\\-\\+]\\d{1,2}:\\d{1,2} "><\\\/head><body>BODY<object data-attachment="MYATTACHMENT" data="name:Attachment\\d{0,4}" type="application\\\/pdf" \\\/><\\\/body><\\\/html>\\r\\n--OneNoteTypedDataBoundary\\d{0,4}\\r\\nContent-Type: application\\\/pdf\\r\\nContent-Disposition: form-data; name="Attachment\\d{0,4}"\\r\\n\\r\\n.*\\r\\n--OneNoteTypedDataBoundary\\d{0,4}--\\r\\n'),
			"AddAttachment should add the attachment as an object element which references the data part containing the binary");
		done();
	};
	reader.readAsText(blobPayload);
});

test("The blob payload is formatted according to what OneNote API's POST docs expect after AddUrl", (assert: QUnitAssert) => {
	let done = assert.async();

	let page = new OneNotePage(title, body);
	page.addUrl("url");
	let blobPayload = page.getTypedFormData().asBlob();

	let reader = new FileReader();
	reader.onloadend = () => {
		let blobAsText = reader.result;
		ok(blobAsText.match('--OneNoteTypedDataBoundary\\d{0,4}\\r\\nContent-Type: application\\\/xhtml\\+xml\\r\\nContent-Disposition: form-data; name=\"Presentation\"\\r\\n\\r\\n<html xmlns="http:\\\/\\\/www.w3.org\\\/1999\\\/xhtml" lang=en-us><head><title>TITLE<\\\/title><meta name="created" content="[\\-\\+]\\d{1,2}:\\d{1,2} "><\\\/head><body>BODY<div data-render-src="url" data-render-method="extract" data-render-fallback="none"><\\\/div><\\\/body><\\\/html>\\r\\n--OneNoteTypedDataBoundary\\d{0,4}--\\r\\n'),
			"AddUrl should simply append div tags to the body referencing the url, specifying the extract method with no fallback");
		done();
	};
	reader.readAsText(blobPayload);
});

test("The blob payload is formatted according to what OneNote API's POST docs expect after AddCitation when the rawUrl is specified", (assert: QUnitAssert) => {
	let done = assert.async();

	let page = new OneNotePage(title, body);
	page.addCitation("START{0}FINISH", "URLTODISPLAY", "RAWURL");
	let blobPayload = page.getTypedFormData().asBlob();

	let reader = new FileReader();
	reader.onloadend = () => {
		let blobAsText = reader.result;
		ok(blobAsText.match('--OneNoteTypedDataBoundary\\d{0,4}\\r\\nContent-Type: application\\\/xhtml\\+xml\\r\\nContent-Disposition: form-data; name=\"Presentation\"\\r\\n\\r\\n<html xmlns="http:\\\/\\\/www.w3.org\\\/1999\\\/xhtml" lang=en-us><head><title>TITLE<\\\/title><meta name="created" content="[\\-\\+]\\d{1,2}:\\d{1,2} "><\\\/head><body>BODYSTART<a href="RAWURL">URLTODISPLAY<\\\/a>FINISH<\\\/body><\\\/html>\\r\\n--OneNoteTypedDataBoundary\\d{0,4}--\\r\\n'),
			"AddCitation should add the url displayed as the display url to the body");
		done();
	};
	reader.readAsText(blobPayload);
});

test("The blob payload is formatted according to what OneNote API's POST docs expect after AddCitation when the rawUrl is not specified", (assert: QUnitAssert) => {
	let done = assert.async();

	let page = new OneNotePage(title, body);
	page.addCitation("START{0}FINISH", "URLTODISPLAY");
	let blobPayload = page.getTypedFormData().asBlob();

	let reader = new FileReader();
	reader.onloadend = () => {
		let blobAsText = reader.result;
		ok(blobAsText.match('--OneNoteTypedDataBoundary\\d{0,4}\\r\\nContent-Type: application\\\/xhtml\\+xml\\r\\nContent-Disposition: form-data; name=\"Presentation\"\\r\\n\\r\\n<html xmlns="http:\\\/\\\/www.w3.org\\\/1999\\\/xhtml" lang=en-us><head><title>TITLE<\\\/title><meta name="created" content="[\\-\\+]\\d{1,2}:\\d{1,2} "><\\\/head><body>BODYSTART<a href="URLTODISPLAY">URLTODISPLAY<\\\/a>FINISH<\\\/body><\\\/html>\\r\\n--OneNoteTypedDataBoundary\\d{0,4}--\\r\\n'),
			"AddCitation should add the string formatted with the specified url to the body");
		done();
	};
	reader.readAsText(blobPayload);
});

test("EscapeHtmlEntities_EscapedCorrectly", () => {
	let page = new OneNotePage();

	let testString = 'This < is <b>the</b> "title" & other title <';
	equal(page.escapeHtmlEntities(testString), 'This &lt; is &lt;b&gt;the&lt;/b&gt; "title" &amp; other title &lt;');
});

test("EscapeHtmlEntities_NothingToEscape", () => {
	let page = new OneNotePage();

	let testString = "This is the title";
	equal(page.escapeHtmlEntities(testString), "This is the title");
});

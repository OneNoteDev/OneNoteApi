type Blobbable = string | ArrayBuffer;

export interface DataPart {
	name: string;
	content: Blobbable;
	type: string;
}

// A substitute for FormData that allows the ability to set content-type per part
export class TypedFormData {
	/*const*/
	private contentTypeMultipart: string = "multipart/form-data; boundary=";
	private boundaryName: string;
	private dataParts: DataPart[] = [];

	constructor() {
		this.boundaryName = "OneNoteTypedDataBoundary" + Math.floor((Math.random() * 1000));
	}

	public getContentType() {
		return this.contentTypeMultipart + this.boundaryName;
	}

	public append(name: string, content: Blobbable, type: string) {
		this.dataParts.push({ content: content, name: name, type: type });
	}

	public asBlob(): Blob {
		let boundary: string = "--" + this.boundaryName;
		let payload: Blobbable[] = [boundary];
		for (let i = 0; i < this.dataParts.length; i++) {
			let curr = this.dataParts[i];
			let header =
				"\r\n" +
				"Content-Type: " + curr.type + "\r\n" +
				"Content-Disposition: form-data; name=\"" + curr.name + "\"\r\n\r\n";
			payload.push(header);
			payload.push(curr.content);
			payload.push("\r\n" + boundary);
		}
		payload.push("--\r\n");
		return new Blob(payload);
	}
}

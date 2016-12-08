import {BatchRequestOperation} from "./structuredTypes";

/**
 * The BATCH API allows a user to execute multiple OneNoteApi actions in a single HTTP request.
 * For example, sending two PATCHES in the same HTTP request
 * To use, construct a new BatchRequest and then pass in an object that adheres to the BatchRequestOperation interface into
 * 	BatchRequest::addOperation(...). Once the request is built, send it using OneNoteApi::sendBatchRequest(...)
 */
export class BatchRequest {
	private operations: BatchRequestOperation[];
	private boundaryName: string;
	private requestBody: string;

	constructor() {
		this.operations = [];
		this.boundaryName = "batch_" + Math.floor(Math.random() * 1000);
	}

	public addOperation(op: BatchRequestOperation): void {
		this.operations.push(op);
	}

	public getOperation(index: number): BatchRequestOperation {
		return this.operations[index];
	}

	public getNumOperations(): number {
		return this.operations.length;
	}

	public getRequestBody(): string {
		let data = "";
		this.operations.forEach((operation) => {
			let req = "";
			req += "--" + this.boundaryName + "\r\n";
			req += "Content-Type: application/http" + "\r\n";
			req += "Content-Transfer-Encoding: binary" + "\r\n";

			req += "\r\n";

			req += operation.httpMethod + " " + operation.uri + " " + "HTTP/1.1" + "\r\n";
			req += "Content-Type: " + operation.contentType + "\r\n";

			req += "\r\n";

			req += (operation.content ? operation.content : "") + "\r\n";

			req += "\r\n";

			data += req;
		});
		data += "--" + this.boundaryName + "--\r\n";

		return data;
	}

	public getContentType(): string {
		return 'multipart/mixed; boundary="' + this.boundaryName + '"';
	}
}

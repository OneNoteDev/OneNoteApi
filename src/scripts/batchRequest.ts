import {BatchRequestOperation} from "./structuredTypes";

/**
 * The BatchRequest class consists of multiple operations that the user wants to group together.
 * Each operation is a separate request, described in the Operation interface. The BATCH API allows
 * a user to send many operations in one request, and the backend can sometimes optimize these
 */
export class BatchRequest {
	private operations: BatchRequestOperation[];
	private boundaryName: string;
	private requestBody: string;

	constructor() {
		this.operations = [];
		this.boundaryName = "batch_" + Math.floor(Math.random() * 1000);
	}

	public addOperation(op: BatchRequestOperation) {
		this.operations.push(op);
	}

	public getRequestBody(): string {
		// There are separate functions for creating the request body and returning it to the caller
		// It is possible to cache this result, but a user could add on operation, and we would have to recompute it
		// So for now, we take the simple road and compute it on demand
		return this.convertOperationsToHttpRequestBody();
	}

	private convertOperationsToHttpRequestBody(): string {
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

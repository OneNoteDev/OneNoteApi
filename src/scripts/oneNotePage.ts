import {ContentType} from "./contentType";
import {TypedFormData, DataPart} from "./typedFormData";

/**
 * The page payload consists of multiple data parts. The first data part is the 'Presentation'
 * that describes the elements that render on the page itself as ONML. Subsequent data parts are referenced
 * by the Presentation, e.g. binary data.
 */
export class OneNotePage {
	// Non-'Presentation' data parts
	private dataParts: DataPart[] = [];

	private title: string;
	private presentationBody: string;
	private locale: string;
	private pageMetadata: { [key: string]: string };

	constructor(title = "", presentationBody = "", locale = "en-us", pageMetadata: { [key: string]: string } = undefined) {
		this.title = title;
		this.presentationBody = presentationBody;
		this.locale = locale;
		this.pageMetadata = pageMetadata;
	}

	/**
	 * Includes everything inside and including the <html> tags
	 */
	private getEntireOnml(): string {
		return '<html xmlns="http://www.w3.org/1999/xhtml" lang=' + this.locale + ">" +
			this.getHeadAsHtml() +
			"<body>" + this.presentationBody + "</body></html>";
	}

	private getHeadAsHtml(): string {
		let pageMetadataHtml = this.getPageMetadataAsHtml();
		let createTime = this.formUtcOffsetString(new Date());
		return "<head><title>" + this.escapeHtmlEntities(this.title) + "</title>" +
			'<meta name="created" content="' + createTime + ' ">' + pageMetadataHtml + "</head>";
	}

	private getPageMetadataAsHtml(): string {
		let pageMetadataHtml = "";
		if (this.pageMetadata) {
			for (let key in this.pageMetadata) {
				pageMetadataHtml += '<meta name="' + key + '" content="' + this.escapeHtmlEntities(this.pageMetadata[key]) + '" />';
			}
		}
		return pageMetadataHtml;
	}

	private formUtcOffsetString(date: Date): string {
		let offset: number = date.getTimezoneOffset();

		let sign = offset >= 0 ? "-" : "+";
		offset = Math.abs(offset);

		let hours: string = Math.floor(offset / 60) + "";
		let mins: string = Math.round(offset % 60) + "";

		if (parseInt(hours, 10) < 10) {
			hours = "0" + hours;
		}

		if (parseInt(mins, 10) < 10) {
			mins = "0" + mins;
		}

		return sign + hours + ":" + mins;
	}

	private generateMimePartName(prefix: string): string {
		return prefix + Math.floor(Math.random() * 10000).toString();
	}

	public escapeHtmlEntities(value: string): string {
		let divElement: HTMLDivElement = document.createElement("div");
		divElement.innerText = value;
		return divElement.innerHTML;
	}

	/**
	 * Converts the Presentation and subsequent data parts entirely into data parts
	 */
	public getTypedFormData(): TypedFormData {
		let tfd = new TypedFormData();
		tfd.append("Presentation", this.getEntireOnml(), "application/xhtml+xml");

		for (let i = 0; i < this.dataParts.length; i++) {
			let part = this.dataParts[i];
			tfd.append(part.name, part.content, part.type);
		}
		return tfd;
	}

	public addOnml(onml: string) {
		this.presentationBody += onml;
	}

	public addHtml(html: string) {
		let mimeName = this.generateMimePartName("Html");
		this.dataParts.push({
			content: html,
			name: mimeName,
			type: "text/HTML"
		});

		this.addOnml('<img data-render-src="name:' + mimeName + '"/>');
		return mimeName;
	}

	public addImage(imgUrl: string) {
		this.addOnml('<img src="' + imgUrl + '"/>');
	}

	/**
	 * The input can either be a url, or a reference to a MIME part containing binary
	 * e.g., "name:REFERENCE"
	 */
	public addObjectUrlAsImage(url: string) {
		this.addOnml('<img data-render-src="' + url + '"/>');
	}

	public addAttachment(binary: ArrayBuffer, name: string) {
		// We hardcode the name as "Attachment" as each page is only allowed to have one
		let mimeName = this.generateMimePartName("Attachment");
		this.dataParts.push({
			content: binary,
			name: mimeName,
			type: "application/pdf"
		});
		this.addOnml('<object data-attachment="' + name + '" data="name:' + mimeName + '" type="application/pdf" />');
		return mimeName;
	}

	public addUrl(url: string) {
		this.addOnml('<div data-render-src="' + url + '" data-render-method="extract" data-render-fallback="none"></div>');
	}

	public addCitation(format: string, urlToDisplay: string, rawUrl?: string) {
		this.addOnml(format.replace("{0}", '<a href="' + (rawUrl ? rawUrl : urlToDisplay) + '">' + urlToDisplay + "</a>"));
	}
}

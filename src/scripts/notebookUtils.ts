import {Notebook, SectionGroup, Section} from "./structuredTypes";

export type SectionParent = Notebook | SectionGroup;
export type SectionPathElement = SectionParent | Section;

export class NotebookUtils {
	/**
	 * Checks to see if the section exists in the notebook list.
	 *
	 * @param notebooks List of notebooks to search
	 * @param sectionId Section id to check the existence of
	 * @return true if the section exists in the notebooks; false otherwise
	 */
	public static sectionExistsInNotebooks(notebooks: Notebook[], sectionId: string): boolean {
		if (!notebooks || !sectionId) {
			return false;
		}

		for (let i = 0; i < notebooks.length; i++) {
			if (NotebookUtils.sectionExistsInParent(notebooks[i], sectionId)) {
				return true;
			}
		}
		return false;
	}

	/**
	 * Checks to see if the section exists in the notebook or section group.
	 *
	 * @param parent Notebook or section group to search
	 * @param sectionId Section id to check the existence of
	 * @return true if the section exists in the parent; false otherwise
	 */
	public static sectionExistsInParent(parent: SectionParent, sectionId: string): boolean {
		if (!parent || !sectionId) {
			return false;
		}

		if (parent.sections) {
			for (let i = 0; i < parent.sections.length; i++) {
				let section = parent.sections[i];
				if (section && section.id === sectionId) {
					return true;
				}
			}
		}

		if (parent.sectionGroups) {
			for (let i = 0; i < parent.sectionGroups.length; i++) {
				if (NotebookUtils.sectionExistsInParent(parent.sectionGroups[i], sectionId)) {
					return true;
				}
			}
		}

		return false;
	}

	/**
	 * Retrieves the path starting from the notebook to the first ancestor section found that
	 * meets a given criteria.
	 *
	 * @param notebooks List of notebooks to search
	 * @return section path (e.g., [notebook, sectionGroup, section]); undefined if there is none
	 */
	public static getPathFromNotebooksToSection(notebooks: Notebook[], filter: (s: Section) => boolean): SectionPathElement[] {
		if (!notebooks || !filter) {
			return undefined;
		}

		for (let i = 0; i < notebooks.length; i++) {
			let notebook = notebooks[i];
			let notebookSearchResult = NotebookUtils.getPathFromParentToSection(notebook, filter);
			if (notebookSearchResult) {
				return notebookSearchResult;
			}
		}

		return undefined;
	}

	/**
	 * Recursively retrieves the path starting from the specified parent to the first ancestor
	 * section found that meets a given criteria.
	 *
	 * @param parent The notebook or section group to search
	 * @return section path (e.g., [parent, sectionGroup, sectionGroup, section]); undefined if there is none
	 */
	public static getPathFromParentToSection(parent: SectionParent, filter: (s: Section) => boolean): SectionPathElement[] {
		if (!parent || !filter) {
			return undefined;
		}

		if (parent.sections) {
			for (let i = 0; i < parent.sections.length; i++) {
				let section = parent.sections[i];
				if (filter(section)) {
					return [parent, section];
				}
			}
		}

		if (parent.sectionGroups) {
			for (let i = 0; i < parent.sectionGroups.length; i++) {
				let sectionGroup = parent.sectionGroups[i];
				let sectionGroupSearchResult = NotebookUtils.getPathFromParentToSection(sectionGroup, filter);
				if (sectionGroupSearchResult) {
					sectionGroupSearchResult.unshift(parent);
					return sectionGroupSearchResult;
				}
			}
		}

		return undefined;
	}

	/**
	 * Computes the maximum depth of the notebooks list, including sections.
	 *
	 * @param notebooks List of notebooks
	 * @return Maximum depth
	 */
	public static getDepthOfNotebooks(notebooks: Notebook[]): number {
		if (!notebooks || notebooks.length === 0) {
			return 0;
		}
		return notebooks.map((notebook) => NotebookUtils.getDepthOfParent(notebook)).reduce((d1, d2) => Math.max(d1, d2));
	}

	/**
	 * Computes the maximum depth of the non-section parent entity, including sections.
	 *
	 * @param notebooks Non-section parent entity
	 * @return Maximum depth
	 */
	public static getDepthOfParent(parent: SectionParent): number {
		if (!parent) {
			return 0;
		}

		let containsAtLeastOneSection = parent.sections && parent.sections.length > 0;
		let maxDepth = containsAtLeastOneSection ? 1 : 0;

		if (parent.sectionGroups) {
			for (let i = 0; i < parent.sectionGroups.length; i++) {
				maxDepth = Math.max(NotebookUtils.getDepthOfParent(parent.sectionGroups[i]), maxDepth);
			}
		}

		// Include the parent itself
		return maxDepth + 1;
	}
}

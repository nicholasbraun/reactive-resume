// @vitest-environment happy-dom

import type { ResumeData } from "@reactive-resume/schema/resume/data";
import { describe, expect, it } from "vitest";
import { sampleResumeData } from "@reactive-resume/schema/resume/sample";
import { buildDocument } from "./builder";

type Attrs = Record<string, unknown>;

/** Collects the attributes of every element called `name` in a prepForXml tree. */
function attrsOf(tree: unknown, name: string): Attrs[] {
	const found: Attrs[] = [];
	const walk = (node: unknown): void => {
		if (Array.isArray(node)) {
			for (const item of node) walk(item);
			return;
		}
		if (!node || typeof node !== "object") return;
		for (const [key, value] of Object.entries(node)) {
			if (key === name) {
				const attrs = Array.isArray(value)
					? value.find((item) => item && typeof item === "object" && "_attr" in item)?._attr
					: (value as { _attr?: Attrs } | undefined)?._attr;
				found.push(attrs ?? {});
			}
			walk(value);
		}
	};
	walk(tree);
	return found;
}

function render(data: ResumeData) {
	const file = buildDocument(data);
	const context = { file, viewWrapper: file.Document, stack: [] };
	return {
		styles: file.Styles.prepForXml(context),
		body: file.Document.View.prepForXml(context),
	};
}

/** Width of the page's text area in twips, read from the section properties. */
function textAreaWidth(body: unknown): number {
	const [size] = attrsOf(body, "w:pgSz");
	const [margin] = attrsOf(body, "w:pgMar");
	return Number(size?.["w:w"]) - Number(margin?.["w:left"]) - Number(margin?.["w:right"]);
}

function fullWidthSample(): ResumeData {
	const data = structuredClone(sampleResumeData);
	const [page] = data.metadata.layout.pages;
	data.metadata.layout.pages = [
		{ fullWidth: true, main: [...(page?.main ?? []), ...(page?.sidebar ?? [])], sidebar: [] },
	];
	return data;
}

function twoColumnSample(): ResumeData {
	const data = structuredClone(sampleResumeData);
	data.metadata.template = "lapras";
	data.metadata.layout.sidebarWidth = 35;
	data.metadata.layout.pages = [
		{ fullWidth: false, main: ["summary", "experience"], sidebar: ["skills", "education"] },
	];
	return data;
}

describe("DOCX page layout", () => {
	it("declares Normal as the default paragraph style", () => {
		// Without a default paragraph style, Apple Pages gives every unstyled paragraph the style of the
		// paragraph before it — so everything after a section heading renders as a heading.
		const { styles } = render(fullWidthSample());
		const normal = attrsOf(styles, "w:style").find((attrs) => attrs["w:styleId"] === "Normal");

		expect(normal?.["w:type"]).toBe("paragraph");
		expect(String(normal?.["w:default"])).toMatch(/^(1|true)$/);
	});

	it("sizes the two-column grid in twips and fixes the table layout", () => {
		// Pages and LibreOffice lay columns out from <w:tblGrid>; the docx default is 100 twips per column.
		const { body } = render(twoColumnSample());
		const columns = attrsOf(body, "w:gridCol").map((attrs) => Number(attrs["w:w"]));
		const [main = 0, sidebar = 0] = columns;

		expect(columns).toHaveLength(2);
		expect(Math.abs(main + sidebar - textAreaWidth(body))).toBeLessThanOrEqual(1);
		expect(sidebar / (main + sidebar)).toBeCloseTo(0.35, 2);
		expect(attrsOf(body, "w:tblLayout").map((attrs) => attrs["w:type"])).toEqual(["fixed"]);
	});
});

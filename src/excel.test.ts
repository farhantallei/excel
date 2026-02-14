import { expect, test } from "bun:test"
import XLSX from "xlsx"

import { buildSheet } from "./builder"
import { applyColumnConfig } from "./formatter"
import { exportToExcel } from "./index"
import { writeToBlob } from "./writer"

test("buildSheet uses headerRows, footer, and merges", () => {
	const { ws, totalRows, headerRowCount } = buildSheet({
		sheetName: "Any",
		headerRows: [
			["Title", "Subtitle"],
			["Header A", "Header B"],
		],
		footerRows: [["Total", 2]],
		columns: [],
		data: [{ value: "x" }, { value: "y" }],
		mapRow: (item) => [item.value, item.value.toUpperCase()],
		merges: [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }],
	})

	const rows = XLSX.utils.sheet_to_json(ws, { header: 1 })

	expect(rows).toEqual([
		["Title", "Subtitle"],
		["Header A", "Header B"],
		["x", "X"],
		["y", "Y"],
		["Total", 2],
	])
	expect(totalRows).toBe(5)
	expect(headerRowCount).toBe(2)
	expect(ws["!merges"]).toEqual([{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }])
})

test("buildSheet falls back to header and honors headerRowCount override", () => {
	const { ws, totalRows, headerRowCount } = buildSheet({
		sheetName: "Any",
		header: ["Col1", "Col2"],
		headerRowCount: 3,
		columns: [],
		data: [],
		mapRow: () => [],
	})

	expect(XLSX.utils.sheet_to_json(ws, { header: 1 })).toEqual([
		["Col1", "Col2"],
	])
	expect(totalRows).toBe(1)
	expect(headerRowCount).toBe(3)
})

test("applyColumnConfig sets widths, types, and formats", () => {
	const ws = XLSX.utils.aoa_to_sheet([
		["H1", "H2", "H3", "H4"],
		[1234.5, "text", "2020-01-02", 99],
		[2, "other", "invalid-date", 50],
	])

	applyColumnConfig({
		ws,
		columns: [
			{ width: 80, type: "number", format: "0.00" },
			{ width: 90, type: "string", format: "@" },
			{ width: 100, type: "date", format: "yyyy-mm-dd" },
			{ width: 110 },
		],
		totalRows: 3,
		headerRowCount: 1,
	})

	expect(ws["!cols"]).toEqual([
		{ wpx: 80 },
		{ wpx: 90 },
		{ wpx: 100 },
		{ wpx: 110 },
	])

	const numCell = ws.A2
	expect(numCell.t).toBe("n")
	expect(numCell.z).toBe("0.00")

	const strCell = ws.B2
	expect(strCell.t).toBe("s")
	expect(strCell.z).toBe("@")

	const dateCell = ws.C2
	expect(dateCell.t).toBe("d")
	expect(dateCell.v).toBeInstanceOf(Date)
	expect(dateCell.z).toBe("yyyy-mm-dd")

	const invalidDateCell = ws.C3
	expect(invalidDateCell.v).toBe("invalid-date")
	expect(invalidDateCell.t).toBe("s")
	expect(invalidDateCell.z).toBe("yyyy-mm-dd")
})

test("writeToBlob creates readable workbook", async () => {
	const ws = XLSX.utils.aoa_to_sheet([["hello", 42]])
	const blob = writeToBlob(ws, "SheetOne")
	const wb = XLSX.read(await blob.arrayBuffer())
	expect(wb.SheetNames).toEqual(["SheetOne"])
	const sheet = wb.Sheets.SheetOne
	if (!sheet) {
		throw new Error("Expected worksheet 'SheetOne' to exist")
	}
	expect(sheet.A1.v).toBe("hello")
	expect(sheet.B1.v).toBe(42)
})

test("exportToExcel builds worksheet and applies config", async () => {
	const blob = exportToExcel({
		sheetName: "Report",
		header: ["No", "Name"],
		columns: [
			{ width: 50, type: "number" },
			{ width: 120, type: "string" },
		],
		data: [
			{ n: 1, name: "Alpha" },
			{ n: 2, name: "Beta" },
		],
		mapRow: (item, idx) => [item.n + idx, item.name],
		merges: [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }],
	})

	const wb = XLSX.read(await blob.arrayBuffer())
	const sheet = wb.Sheets.Report
	if (!sheet) {
		throw new Error("Expected worksheet 'Report' to exist")
	}

	expect(wb.SheetNames).toEqual(["Report"])
	expect(sheet["!merges"]).toEqual([{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }])
	expect(XLSX.utils.sheet_to_json(sheet, { header: 1 })).toEqual([
		["No", "Name"],
		[1, "Alpha"],
		[3, "Beta"],
	])
})

import XLSX from "xlsx"

import { buildSheet } from "./builder"
import { applyColumnConfig } from "./formatter"
import type { ExcelExportConfig } from "./types"
import { writeToBlob } from "./writer"

/**
 * Generate an Excel file (XLSX) from a declarative configuration.
 *
 * This function builds a worksheet using the provided header, data, and row mapping
 * defined in the configuration object. Each column can define metadata such as
 * data type (string/number/date), column width, and optional custom formatting.
 * Once the worksheet is constructed, the function returns a downloadable XLSX `Blob`.
 *
 * ! Environment Constraint:
 * - This function must be executed in a server-side environment
 *   as `xlsx` is not browser-compatible.
 * - The output is a `Blob` that can safely be sent to the client for download.
 *
 * Guarantees:
 * - No global state is modified.
 * - No external I/O is performed beyond internal workbook operations.
 * - All XLSX logic is fully encapsulated; the caller only provides configuration.
 *
 * @typeParam T - The type of each item in the `data` array.
 *
 * @param config - Excel export configuration.
 *  - `sheetName`: The name of the worksheet.
 *  - `header`: Column order for the first row of the Excel file.
 *  - `columns`: Column metadata (width/type/format).
 *  - `data`: Raw data to be mapped into Excel rows.
 *  - `mapRow`: A function that maps each data item into an array of cell values (aligned with the header).
 *
 * @returns A downloadable Excel `Blob`.
 *
 * @example
 * const blob = await exportToExcel({
 *   sheetName: "Master Banner",
 *   header: ["No", "Title", "Status"],
 *   columns: [
 *     { width: 35, type: "number" },
 *     { width: 300, type: "string" },
 *     { width: 100, type: "string" },
 *   ],
 *   data: banners,
 *   mapRow: (item, idx) => [idx + 1, item.title, item.active ? "Active" : "Inactive"],
 * })
 *
 * // on the client
 * const url = URL.createObjectURL(blob)
 * downloadFile(url, "master-banner.xlsx")
 */
export function exportToExcel<T = unknown>(config: ExcelExportConfig<T>) {
	const { ws, totalRows, headerRowCount } = buildSheet(config)

	applyColumnConfig({
		ws,
		columns: config.columns,
		totalRows,
		headerRowCount,
	})

	return writeToBlob(ws, config.sheetName)
}

export async function readExcelToJson<T>(file: File): Promise<T[]> {
	const arrayBuffer = await file.arrayBuffer()
	const workbook = XLSX.read(arrayBuffer, { type: "array" })
	const sheetName = workbook.SheetNames[0]
	if (!sheetName) {
		throw new Error("The Excel file does not contain any worksheets.")
	}
	const worksheet = workbook.Sheets[sheetName]
	if (!worksheet) {
		throw new Error(`Worksheet "${sheetName}" could not be found.`)
	}
	return XLSX.utils.sheet_to_json(worksheet, { defval: "" })
}

export type { ExcelExportConfig }

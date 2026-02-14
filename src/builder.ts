import XLSX from "xlsx"

import type { ExcelExportConfig } from "./types"

export function buildSheet<T>(config: ExcelExportConfig<T>) {
	const { header, headerRows, data, mapRow, footerRows } = config

	const rows: (string | number | Date | null)[][] = []

	if (headerRows && headerRows.length > 0) {
		for (let i = 0; i < headerRows.length; i++) {
			const row = headerRows[i]
			if (row) rows.push(row)
		}
	} else if (header && header.length > 0) {
		rows.push(header)
	}

	for (let i = 0; i < data.length; i++) {
		const item = data[i]
		if (item) rows.push(mapRow(item, i))
	}

	if (footerRows && footerRows.length > 0) {
		for (let i = 0; i < footerRows.length; i++) {
			const footerRow = footerRows[i]
			if (footerRow) rows.push(footerRow)
		}
	}

	const ws = XLSX.utils.aoa_to_sheet(rows)

	if (config.merges?.length) {
		ws["!merges"] = config.merges
	}

	const totalRows = rows.length

	const headerRowCount =
		config.headerRowCount ?? (config.headerRows ? config.headerRows.length : 1)

	return { ws, totalRows, headerRowCount }
}

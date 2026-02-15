import * as XLSX from "xlsx"

import type { ColumnConfig } from "./types"

interface ApplyProps {
	ws: XLSX.WorkSheet
	columns: ColumnConfig[]
	totalRows: number
	headerRowCount: number
}

export function applyColumnConfig({
	ws,
	columns,
	totalRows,
	headerRowCount,
}: ApplyProps) {
	ws["!cols"] = columns.map((col) => ({ wpx: col.width ?? 100 }))

	for (let colIndex = 0; colIndex < columns.length; colIndex++) {
		const col = columns[colIndex]
		if (!col) continue
		if (!col.type && !col.format) continue

		const colLetter = XLSX.utils.encode_col(colIndex)

		for (let row = headerRowCount; row < totalRows; row++) {
			const cellAddress = `${colLetter}${row + 1}`
			const cell = ws[cellAddress]
			if (!cell) continue

			switch (col.type) {
				case "number":
					if (typeof cell.v === "number") {
						cell.t = "n"
						cell.z = col.format || "#,##0"
					}
					break

				case "date": {
					const dt = new Date(cell.v)
					if (!Number.isNaN(dt.getTime())) {
						cell.t = "d"
						cell.v = dt
					}
					break
				}

				default:
					cell.t = "s"
			}

			if (col.format) {
				cell.z = col.format
			}
		}
	}
}

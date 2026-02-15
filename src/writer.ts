import * as XLSX from "xlsx"

export function writeToBlob(ws: XLSX.WorkSheet, sheetName: string) {
	const wb = XLSX.utils.book_new()
	XLSX.utils.book_append_sheet(wb, ws, sheetName)

	const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" })

	return new Blob([wbout], {
		type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
	})
}

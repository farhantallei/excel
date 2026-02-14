export interface CellMergeRange {
	s: { r: number; c: number } // start row/col (0-based)
	e: { r: number; c: number } // end row/col (0-based)
}

export interface ColumnConfig {
	width?: number
	type?: "string" | "number" | "date"
	format?: string
}

export interface ExcelExportConfig<T = unknown> {
	sheetName: string
	header?: string[]
	headerRows?: (string | null)[][]
	footerRows?: (string | number | Date | null)[][]
	headerRowCount?: number
	columns: ColumnConfig[]
	data: T[]
	mapRow: (item: T, index: number) => (string | number | Date | null)[]
	merges?: CellMergeRange[]
}

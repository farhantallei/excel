# @farhantallei/excel

> Type-safe, declarative Excel (XLSX) export for Node.js server environments.

Build Excel files using a clean configuration object — with typed columns, formatting, and row mapping. Designed for server-side execution only.

---

## Features

- ✅ Declarative worksheet configuration
- ✅ Fully type-safe (TypeScript-first)
- ✅ Column metadata (width, type, format)
- ✅ Generic row mapping
- ✅ Server-only safe (no browser XLSX usage)
- ✅ Zero global state
- ✅ Clean separation of concerns
- ✅ JSON import from Excel  

---

## Installation

```bash
npm install @farhantallei/excel
# or
yarn add @farhantallei/excel
```

---

## Usage

### Quick Start (Export)

```ts
import { exportToExcel } from "@farhantallei/excel"

const blob = await exportToExcel({
  sheetName: "Users",
  header: ["No", "Name", "Active"],
  columns: [
    { width: 50, type: "number" },
    { width: 200, type: "string" },
    { width: 100, type: "string" },
  ],
  data: users,
  mapRow: (user, index) => [
    index + 1,
    user.name,
    user.active ? "Yes" : "No",
  ],
})
```

#### Send to Client (Next.js Example)

```ts
// Server Action / Route Handler
return new Response(blob, {
  headers: {
    "Content-Type":
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "Content-Disposition": 'attachment; filename="users.xlsx"',
  },
})
```

### Import Excel → JSON

```ts
import { readExcelToJson } from "@farhantallei/excel"

const data = await readExcelToJson<User>(file)
```

#### Optional sheet selection:

```ts
await readExcelToJson<User>(file, {
  sheetName: "Users",
})
```

---

## API

### `exportToExcel<T>(config)`

Generate an Excel `Blob` from a declarative configuration.

#### Config

```ts
{
  sheetName: string
  header: string[]
  columns: {
    width?: number
    type?: "string" | "number" | "date"
    format?: string
  }[]
  data: T[]
  mapRow: (item: T, index: number) => unknown[]
}
```

#### Returns

```ts
Promise<Blob>
```

### `readExcelToJson<T>(file, options?)`

Read an Excel file and convert the first (or specified) worksheet into JSON.

#### Options

```ts
{
  sheetName?: string
}
```

#### Returns

```ts
Promise<T[]>
```

### Error Handling

* `APIError` – Standardized error class for HTTP requests

```ts
try {
  await fetcher("/users")({ method: "GET" })
} catch (err) {
  if (err instanceof APIError) {
    console.error(err.status, err.message)
  }
}
```

---

## ⚠️ Environment Constraint

This library is server-side only.

It must be executed in:

- Node.js
- Next.js Server Actions
- Next.js Route Handlers
- Any server runtime

It **must not** be used in the browser, as `xlsx` is not browser-compatible in this setup.

---

## 🏗 Design Principles

- No global state mutation
- No external I/O
- XLSX logic fully encapsulated
- Caller provides only configuration
- Strict TypeScript compatibility

---

## Contributing

PRs and issues are welcome. please fork the repository, create a feature branch, ensure all tests and type checks pass, and submit a pull request.

---

## License

MIT

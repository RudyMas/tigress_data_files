# Tigress Data Files — Programmer's Manual

**Version:** 2025.12.09  
**Package:** `tigress/data-files`  
**License:** GPL-3.0-or-later  
**Requires:** PHP >= 8.5, `shuchkin/simplexlsxgen` >= 1.4

---

## Overview

`Tigress\DataFiles` is a PHP class that manages tabular data in memory and exports it to CSV, Excel (XLSX), or JSON. You populate rows via an array-based API, optionally attach header/footer rows and an index row, and then write the result to a file.

---

## Installation

```bash
composer require tigress/data-files
```

---

## Quick Start

```php
use Tigress\DataFiles;

$df = new DataFiles();

// Populate data
$df->addLine(['Alice', 30, 'Engineer']);
$df->addLine(['Bob',   25, 'Designer']);

// Optional: header, footer, index
$df->setHeader([['Name', 'Age', 'Role']]);
$df->setFooter([['Report generated ' . date('Y-m-d')]]);

// Export
$path = $df->createCsvFile('report', 'output/');
echo "Written to: $path\n";
```

---

## Data Management

### `addLine(array $data): void`
Appends one row.

### `addLines(array $data): void`
Merges multiple rows at once. Internally uses `array_merge`, so numeric keys are **re-indexed** sequentially.

```php
$df->addLines([
    ['Alice', 30],
    ['Bob',   25],
]);
```

### `changeLine(int $index, array $data): void`
Replaces the row at a given index.

```php
$df->changeLine(0, ['Charlie', 35]);
```

### `changeLines(array $indexes, array $data): void`
Replaces **every** index in `$indexes` with the **same** `$data` array.

```php
$df->changeLines([0, 2], ['X', 0]);  // rows 0 and 2 both become ['X', 0]
```

### `removeLine(int $index): void`
Removes the row at `$index` using `unset`. This **leaves a gap** in the array — the array is not re-indexed.

### `removeLines(array $indexes): void`
Removes multiple rows at once.

### `reset(): void`
Empties all stored data (does NOT clear header, footer, or indexList).

```php
$df->reset();
```

### `getData(): array`
Returns the internal data array.

---

## Creating Data from a Database

### `createArray(array $data): void`
Converts an array of associative/hybrid rows (e.g. from a PDO fetch) into plain indexed arrays. **Overwrites** any existing data.

```php
$rows = $db->fetchAll('SELECT name, age FROM users');
$df->createArray($rows);
// Each row is now [name, age] — column keys are dropped.
```

Equivalent pseudo-code: `array_map('array_values', $rows)`.

---

## Header, Footer & Index

### `setHeader(array $header): void`
One or more rows prepended to every export. Must be an array of arrays — each sub-array is one row.

```php
$df->setHeader([
    ['Product Inventory'],
    ['Item', 'Qty', 'Price'],
]);
```

### `setFooter(array $footer): void`
One or more rows appended to every export. Same structure as header.

```php
$df->setFooter([
    ['Total:', 150],
]);
```

HTML entities in header/footer cells are decoded automatically (`html_entity_decode` with `ENT_QUOTES | UTF-8`).

### `setIndexList(array $indexList): void`
Sets an **index row** — a single row that can optionally appear at the top and/or bottom of the data section (controlled by `$addIndexStart` / `$addIndexEnd` in export methods).

```php
$df->setIndexList(['#', 'Product', 'Count', 'Price']);
// This row will be inserted between header and data, and/or after data.
```

### `getIndexList(): array`
Returns the current index list.

---

## Exporting Files

### `createCsvFile(...): string`
Export to CSV. Returns the full file path of the written file.

```php
$df->createCsvFile(
    filename: 'report',
    filepath: 'exports/',        // '' = current dir (default)
    delimiter: ',',              // default
    enclosure: '"',              // default
    escape: '\\',                // default
    addIndexStart: true,         // insert index row before data
    addIndexEnd: false,          // insert index row after data
    addBOM: false                // prepend UTF-8 BOM for Excel compat
);
```

The `.csv` extension is appended automatically if missing.

### `createExcel(...): string`
Export to XLSX using SimpleXLSXGen. Returns the full file path.

```php
$df->createExcel(
    filename: 'report',
    filepath: 'exports/',
    fontSize: 13,                // default
    addIndexStart: true,
    addIndexEnd: false
);
```

The `.xlsx` extension is appended automatically if missing.

**Merge cells** (Excel only) — see below.

### `createJsonFile(...): string`
Export to JSON (`JSON_PRETTY_PRINT`). Returns the full file path.

```php
$df->createJsonFile(
    filename: 'report',
    filepath: 'exports/',
    addIndexStart: true,
    addIndexEnd: false
);
```

The `.json` extension is appended automatically if missing.

---

## Excel Merge Cells

### `setMergeCells(array $mergeCells): void`
Sets cell merge ranges for Excel output. Each element is a string passed directly to `SimpleXLSXGen::mergeCells()`.

```php
$df->setMergeCells(['A1:C1']);  // merge cells A1 through C1
```

Only takes effect when calling `createExcel()`.

---

## Utility

### `version(): string` (static)
Returns the library version string.

```php
echo DataFiles::version();  // "2025.12.09"
```

---

## Output Structure Diagram

For `addIndexStart = true`, `addIndexEnd = false`:

```
[ header row 0 ]      ← setHeader()
[ header row 1 ]
...
[ index row ]          ← setIndexList()
[ data row 0 ]
[ data row 1 ]
...
[ footer row 0 ]      ← setFooter()
[ footer row 1 ]
...
```

For `addIndexEnd = true`, the index row is also appended after all data rows (before footer).

---

## Notes & Caveats

1. **Index gaps after removal:** `removeLine()` / `removeLines()` use `unset`, so indices become sparse. `addLines()` re-indexes via `array_merge`. Be aware when referencing indices after removal.
2. **`changeLines()` applies the same data:** All indices listed get the same row array — there is no one-to-one mapping. For individual changes, call `changeLine()` in a loop.
3. **`createArray()` overwrites:** It always starts from an empty data array.
4. **Header/footer are arrays of rows:** Each call to `setHeader()`/`setFooter()` expects an array of arrays, not a flat array. For a single row: `setHeader([['a', 'b', 'c']])`.
5. **HTML decoding:** Only applied to header and footer cells, not to data rows.
6. **File extension:** Automatically appended only if the filename does not already end with the correct extension.
7. **Directory creation:** If `filepath` is non-empty and does not exist, it is created recursively with `0777` permissions.

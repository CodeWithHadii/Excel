# Excel

This extension allows you to interact with Excel files on Android using App Inventor. You can perform operations such as adding rows, deleting rows, summing columns, finding cell data, and clearing the sheet. It uses the **JXL library** to handle Excel file manipulations.

## Features

- **Add Row Data**: Add a new row to the Excel sheet with provided data.
- **Delete Row Data**: Delete a specific row by its row number.
- **Sum Column Data**: Sum all numeric values in a specified column.
- **Find Cell Data**: Find the first occurrence of a value in a column and return the corresponding row number.
- **Clear Sheet**: Clear all data in the sheet.
- **Events**: Each operation triggers events to notify success or failure.

## Requirements

- **JXL Library**: The extension relies on the JXL library for Excel file handling.

## Installation

1. Download the extension `.aix` file and import it into your App Inventor project.
2. Use the provided blocks to call functions and handle the corresponding events.

## Properties

- **ExcelFilePath**: Set the file path for the Excel file.
- **SheetName**: Set the sheet name (default is `Sheet1`).

## Functions

### `AddRowData(List<String> rowData)`
Adds a new row with the given data.

**Arguments**:
- `rowData`: A list of strings to add as a new row in the Excel sheet.

### `DeleteRowData(int rowNumber)`
Deletes a specific row by its number.

**Arguments**:
- `rowNumber`: The row number to delete (starting from 0).

### `SumColumnData(String columnName)`
Sums all numeric values in the specified column.

**Arguments**:
- `columnName`: The name of the column to sum (case-insensitive).

### `FindCellData(String columnName, String valueToFind)`
Finds the first occurrence of a value in the specified column and returns the row number.

**Arguments**:
- `columnName`: The name of the column to search in.
- `valueToFind`: The value to search for.

### `ClearSheet()`
Clears all data in the sheet.

## Events

### `AfterAddRowData(String message, String status)`
Triggered after adding a row. `status` will indicate whether the operation succeeded or failed.

### `AfterDeleteRowData(String message, String status)`
Triggered after deleting a row. `status` will indicate whether the operation succeeded or failed.

### `AfterSumColumnData(double sum, String status)`
Triggered after summing a column. `status` will indicate success or failure.

### `AfterFindCellData(int rowNumber, String status)`
Triggered after searching for a value in a column. Returns the row number where the value was found or -1 if not found.

### `AfterClearSheet(String message, String status)`
Triggered after clearing the sheet. `status` will indicate success or failure.

## Error Handling

The extension provides event callbacks to notify users about success or failure. You can use these events to handle any exceptions or errors.

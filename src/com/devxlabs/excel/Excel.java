package com.devxlabs.excel;

import java.io.File;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import jxl.Cell;
import jxl.Workbook;
import jxl.Sheet;
import jxl.write.Label;
import jxl.write.WritableWorkbook;
import jxl.write.WritableSheet;
import com.google.appinventor.components.annotations.SimpleFunction;
import com.google.appinventor.components.annotations.SimpleEvent;
import com.google.appinventor.components.annotations.SimpleProperty;
import com.google.appinventor.components.runtime.AndroidNonvisibleComponent;
import com.google.appinventor.components.runtime.ComponentContainer;
import com.google.appinventor.components.runtime.EventDispatcher;

public class Excel extends AndroidNonvisibleComponent {

    private String excelFilePath;
    private String sheetName = "Sheet1";

    public Excel(ComponentContainer container) {
        super(container.$form());
    }

    @SimpleProperty(description = "Sets the file path of the Excel file.")
    public void ExcelFilePath(String path) {
        this.excelFilePath = path;
    }

    @SimpleProperty(description = "Sets the name of the sheet within the Excel file.")
    public void SheetName(String name) {
        this.sheetName = name;
    }

    @SimpleFunction(description = "Adds a new row with the provided data.")
    public void AddRowData(List<String> rowData) {
        try {
            Workbook workbook = Workbook.getWorkbook(new File(excelFilePath));
            WritableWorkbook writableWorkbook = Workbook.createWorkbook(new File(excelFilePath), workbook);
            WritableSheet sheet = writableWorkbook.getSheet(sheetName);
            int lastRow = sheet.getRows();
            for (int i = 0; i < rowData.size(); i++) {
                sheet.addCell(new Label(i, lastRow, rowData.get(i)));
            }
            writableWorkbook.write();
            writableWorkbook.close();
            workbook.close();
            AfterAddRowData("Row added successfully.", "Success");
        } catch (Exception e) {
            AfterAddRowData("Error: " + e.getMessage(), "Failure");
        }
    }

    @SimpleEvent(description = "Event triggered after a row is added successfully or fails.")
    public void AfterAddRowData(String message, String status) {
        EventDispatcher.dispatchEvent(this, "AfterAddRowData", message, status);
    }

    @SimpleFunction(description = "Deletes a specific row identified by its row number.")
    public void DeleteRowData(int rowNumber) {
        try {
            Workbook workbook = Workbook.getWorkbook(new File(excelFilePath));
            WritableWorkbook writableWorkbook = Workbook.createWorkbook(new File(excelFilePath), workbook);
            WritableSheet sheet = writableWorkbook.getSheet(sheetName);
            int rows = sheet.getRows();
            WritableSheet newSheet = writableWorkbook.createSheet(sheetName, 0);
            
            for (int i = 0; i < rows; i++) {
                if (i != rowNumber) {
                    for (int j = 0; j < sheet.getColumns(); j++) {
                        newSheet.addCell(new Label(j, i < rowNumber ? i : i - 1, sheet.getCell(j, i).getContents()));
                    }
                }
            }
            writableWorkbook.write();
            writableWorkbook.close();
            workbook.close();
            AfterDeleteRowData("Row deleted successfully.", "Success");
        } catch (Exception e) {
            AfterDeleteRowData("Error: " + e.getMessage(), "Failure");
        }
    }

    @SimpleEvent(description = "Event triggered after a row is deleted successfully or fails.")
    public void AfterDeleteRowData(String message, String status) {
        EventDispatcher.dispatchEvent(this, "AfterDeleteRowData", message, status);
    }

    @SimpleFunction(description = "Sums all numeric values in a specific column.")
    public void SumColumnData(String columnName) {
        try {
            Workbook workbook = Workbook.getWorkbook(new File(excelFilePath));
            Sheet sheet = workbook.getSheet(sheetName);
            int columnIndex = -1;
            for (int i = 0; i < sheet.getColumns(); i++) {
                if (sheet.getCell(i, 0).getContents().equalsIgnoreCase(columnName)) {
                    columnIndex = i;
                    break;
                }
            }
            if (columnIndex != -1) {
                double sum = 0;
                for (int i = 1; i < sheet.getRows(); i++) {
                    try {
                        double value = Double.parseDouble(sheet.getCell(columnIndex, i).getContents());
                        sum += value;
                    } catch (NumberFormatException e) {
                    }
                }
                AfterSumColumnData(sum, "Success");
            } else {
                AfterSumColumnData(0, "Failure: Column not found.");
            }
            workbook.close();
        } catch (Exception e) {
            AfterSumColumnData(0, "Failure: " + e.getMessage());
        }
    }

    @SimpleEvent(description = "Event triggered after the sum of column data is calculated.")
    public void AfterSumColumnData(double sum, String status) {
        EventDispatcher.dispatchEvent(this, "AfterSumColumnData", sum, status);
    }

    @SimpleFunction(description = "Finds the first occurrence of a value in a column and returns the corresponding row number.")
    public void FindCellData(String columnName, String valueToFind) {
        try {
            Workbook workbook = Workbook.getWorkbook(new File(excelFilePath));
            Sheet sheet = workbook.getSheet(sheetName);
            int columnIndex = -1;
            for (int i = 0; i < sheet.getColumns(); i++) {
                if (sheet.getCell(i, 0).getContents().equalsIgnoreCase(columnName)) {
                    columnIndex = i;
                    break;
                }
            }
            if (columnIndex != -1) {
                for (int i = 1; i < sheet.getRows(); i++) {
                    if (sheet.getCell(columnIndex, i).getContents().equalsIgnoreCase(valueToFind)) {
                        AfterFindCellData(i, "Success");
                        return;
                    }
                }
                AfterFindCellData(-1, "Failure: Value not found.");
            } else {
                AfterFindCellData(-1, "Failure: Column not found.");
            }
            workbook.close();
        } catch (Exception e) {
            AfterFindCellData(-1, "Failure: " + e.getMessage());
        }
    }

    @SimpleEvent(description = "Event triggered after searching for a value in a column.")
    public void AfterFindCellData(int rowNumber, String status) {
        EventDispatcher.dispatchEvent(this, "AfterFindCellData", rowNumber, status);
    }

    @SimpleFunction(description = "Clears all data in the sheet.")
    public void ClearSheet() {
        try {
            Workbook workbook = Workbook.getWorkbook(new File(excelFilePath));
            WritableWorkbook writableWorkbook = Workbook.createWorkbook(new File(excelFilePath), workbook);
            WritableSheet sheet = writableWorkbook.getSheet(sheetName);
            int rows = sheet.getRows();
            int columns = sheet.getColumns();

            for (int i = 0; i < rows; i++) {
                for (int j = 0; j < columns; j++) {
                    sheet.addCell(new Label(j, i, ""));
                }
            }

            writableWorkbook.write();
            writableWorkbook.close();
            workbook.close();
            AfterClearSheet("Sheet cleared successfully.", "Success");
        } catch (Exception e) {
            AfterClearSheet("Error: " + e.getMessage(), "Failure");
        }
    }

    @SimpleEvent(description = "Event triggered after clearing the sheet.")
    public void AfterClearSheet(String message, String status) {
        EventDispatcher.dispatchEvent(this, "AfterClearSheet", message, status);
    }
}

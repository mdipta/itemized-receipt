package com.excel.receipt.util;

import com.excel.receipt.model.Item;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ExcelUtils {

    public static void createExcelReceipt(List<Item> items, String fileName) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Receipt");

        // Create the header row
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Item");
        headerRow.createCell(1).setCellValue("Quantity");
        headerRow.createCell(2).setCellValue("Price");
        headerRow.createCell(3).setCellValue("Total");

        // Fill in the item rows
        int rowIndex = 1;
        double totalAmount = 0;

        for (Item item : items) {
            Row row = sheet.createRow(rowIndex++);
            row.createCell(0).setCellValue(item.getName());
            row.createCell(1).setCellValue(item.getQuantity());
            row.createCell(2).setCellValue(item.getPrice());
            row.createCell(3).setCellValue(item.getTotalPrice());
            totalAmount += item.getTotalPrice();
        }

        // Create a total row
        Row totalRow = sheet.createRow(rowIndex);
        totalRow.createCell(2).setCellValue("Total");
        totalRow.createCell(3).setCellValue(totalAmount);

        // Auto-size columns
        for (int i = 0; i < 4; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
            workbook.write(fileOut);
            System.out.println("Receipt has been written to " + fileName);
        } catch (IOException e) {
            System.out.println("Error writing to Excel file: " + e.getMessage());
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                System.out.println("Error closing workbook: " + e.getMessage());
            }
        }
    }
}

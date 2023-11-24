package org.example;

import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.List;

public class CSVtoExcelConverter {

    public static void main(String[] args) {
        String excelFilePath = "D:\\CSVToExel\\Ak.xlsx";
        String csvFilePath = "D:\\CSVToExel\\Ak.csv";



        try {
            List<String[]> csvData = readCSV(csvFilePath);
            writeExcel(csvData, excelFilePath);
            System.out.println("CSV converted to Excel successfully.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static List<String[]> readCSV(String filePath) throws IOException {
        try (CSVReader csvReader = new CSVReader(new FileReader(filePath))) {
            return csvReader.readAll();
        } catch (CsvException e) {
            throw new RuntimeException(e);
        }
    }

    private static void writeExcel(List<String[]> data, String filePath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");

            // Create the header row
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Name");
            headerRow.createCell(1).setCellValue("Age");

            int rowCount = 1; // Start writing data from the second row
            for (String[] row : data) {
                Row excelRow = sheet.createRow(rowCount++);

                int columnCount = 0;
                for (String value : row) {
                    Cell cell = excelRow.createCell(columnCount++);
                    cell.setCellValue(value);
                }
            }

            try (FileOutputStream fileOutputStream = new FileOutputStream(filePath)) {
                workbook.write(fileOutputStream);
            }
        }
    }
}

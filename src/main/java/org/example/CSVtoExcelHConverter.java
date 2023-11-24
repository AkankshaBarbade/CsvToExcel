package org.example;

import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.List;

public class CSVtoExcelHConverter {

    public static void main(String[] args) {

        String excelFilePath = "D:\\CSVToExel\\Ak.xlsx";
        String csvFilePath = "D:\\CSVToExel\\Ak.csv";
       String headerFilePath = "D:\\CSVToExel\\headerfile.txt";


        try {
            List<String[]> csvData = readCSV(csvFilePath);
            String[] header = readHeader(headerFilePath);

            if (header != null && header.length > 0) {
                writeExcel(csvData, header, excelFilePath);
                System.out.println("CSV converted to Excel successfully.");
            } else {
                System.out.println("Error: Header is missing or empty.");
            }

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

    private static String[] readHeader(String filePath) throws IOException {
        try (BufferedReader br = new BufferedReader(new FileReader(filePath))) {
            String headerLine = br.readLine();
            if (headerLine != null && !headerLine.isEmpty()) {
                return headerLine.split(",");
            }
        }
        return null;
    }

    private static void writeExcel(List<String[]> data, String[] header, String filePath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");

            // Write header
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell headerCell = headerRow.createCell(i);
                headerCell.setCellValue(header[i]);
            }

            // Write data
            int rowCount = 1;
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

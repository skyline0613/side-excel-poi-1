package com.bran.app;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class OrderPrinter {
    public static void main(String[] args) {
        try (FileInputStream fileIn = new FileInputStream("OrderTemplate.xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheetAt(0);

            // Prepare master order data
            Map<String, String> masterData = new HashMap<>();
            masterData.put("${orderId}", "ORD12345");
            masterData.put("${customerName}", "John Doe");
            masterData.put("${orderDate}", "2023-10-01");

            // Fill master order data
            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue();
                        if (masterData.containsKey(cellValue)) {
                            cell.setCellValue(masterData.get(cellValue));
                        }
                    }
                }
            }

            // Prepare order details data
//            List<Map<String, String>> orderDetails = List.of(
//                Map.of("item", "Item A", "quantity", "2", "price", "10.00"),
//                Map.of("item", "Item B", "quantity", "1", "price", "20.00")
//            );

            List<Map<String, String>> orderDetails = new ArrayList<Map<String, String>>();

            Map<String, String> item1 = new HashMap<String, String>();
            item1.put("item", "Item A");
            item1.put("quantity", "2");
            item1.put("price", "10.00");
            orderDetails.add(item1);

            Map<String, String> item2 = new HashMap<String, String>();
            item2.put("item", "Item B");
            item2.put("quantity", "1");
            item2.put("price", "20.00");
            orderDetails.add(item2);            
            
            
            // Fill order details data starting from row 6
            int rowIndex = 5;
            for (Map<String, String> detail : orderDetails) {
                Row row = sheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(detail.get("item"));
                row.createCell(1).setCellValue(detail.get("quantity"));
                row.createCell(2).setCellValue(detail.get("price"));
            }

            // Write filled template to a new file
            try (FileOutputStream fileOut = new FileOutputStream("FilledOrder.xlsx")) {
                workbook.write(fileOut);
            }

            // Close workbook
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
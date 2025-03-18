package com.bran.app;

import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

@SpringBootTest
class ExcelTemplateImplApplicationTests {

	@Test
	void contextLoads() {
	}
	
	
	//@Test
	void run1() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Policy Proposal");

        // 創建標題行
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Customer Name");
        headerRow.createCell(1).setCellValue("Policy Number");
        // 創建變數行
        Row dataRow = sheet.createRow(1);
        dataRow.createCell(0).setCellValue("${customerName}");
        dataRow.createCell(1).setCellValue("${policyNumber}");
        dataRow.createCell(2).setCellValue("${insuranceAmount}");

        // 調整列寬
        for (int i = 0; i < 3; i++) {
            sheet.autoSizeColumn(i);
        }

        // 將工作簿寫入文件
        try (FileOutputStream fileOut = new FileOutputStream("PolicyTemplate.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }

        // 關閉工作簿
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }			
		
		
	}
	
	
	@Test
	void runExcelFiller(){
        try (FileInputStream fileIn = new FileInputStream("PolicyTemplate.xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheetAt(0);

            // 準備數據
            Map<String, String> data = new HashMap<>();
            data.put("${customerName}", "John Doe");
            data.put("${policyNumber}", "123456789");
            data.put("${insuranceAmount}", "1000000");

            // 填充數據
            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue();
                        if (data.containsKey(cellValue)) {
                            cell.setCellValue(data.get(cellValue));
                        }
                    }
                }
            }

            // 將填充後的工作簿寫入文件
            try (FileOutputStream fileOut = new FileOutputStream("FilledPolicy.xlsx")) {
                workbook.write(fileOut);
            }

            // 關閉工作簿
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }		
	}

}

package com.bran.app;

import org.apache.poi.xwpf.usermodel.*;
import java.io.*;

public class POIWordTableExample {
    public static void main(String[] args) {
        try (FileInputStream fis = new FileInputStream("template.docx");
             XWPFDocument document = new XWPFDocument(fis);
             FileOutputStream fos = new FileOutputStream("output.docx")) {

            // 取得 Word 文件中的第一個表格
            XWPFTable table = document.getTables().get(0);
            
            // 在特定的行和列填入數據
            table.getRow(1).getCell(1).setText("John Doe");
            table.getRow(2).getCell(1).setText("Developer");
            table.getRow(3).getCell(1).setText("john.doe@example.com");
            
            // 保存文件
            document.write(fos);
            System.out.println("Word 文件已生成: output.docx");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
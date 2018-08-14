package com.hand.selector;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CreationHelper;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

public class ExcelWrite {
    public static void main(String[] args) {
        //创建一个工作簿，即Excel文件，再该文件中创建一个Sheet
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheets = wb.createSheet("sheet1");

        //在sheet1中创建一行
        /**
         *
         */
        HSSFRow hssfRow = sheets.createRow(0);

        //在该行中插入各种类型的数
        hssfRow.createCell(0).setCellValue(true);
        hssfRow.createCell(1).setCellValue("Jack");
        hssfRow.createCell(2).setCellValue(23);

        //设置保留两位小数
        HSSFCell cell = hssfRow.createCell(3);
        cell.setCellValue(6000);
        HSSFCellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
        cell.setCellStyle(cellStyle);

        //在写入日期格式的 数据需要进行特殊处理（这是一种，简单的处理方式）
        CreationHelper creationHelper = wb.getCreationHelper();
        HSSFCellStyle style = wb.createCellStyle();
        style.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd"));

        cell = hssfRow.createCell(4);
        cell.setCellValue(new Date());
        cell.setCellStyle(style);

        //最后写回磁盘
        FileOutputStream out = null;
        try {
            out = new FileOutputStream("excel write.xls");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            wb.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Over!");
    }
}

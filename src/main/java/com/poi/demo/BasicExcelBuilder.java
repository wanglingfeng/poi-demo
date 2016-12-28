package com.poi.demo;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

/**
 * Created by lfwang on 2016/12/27.
 */
public class BasicExcelBuilder {

    public void generate(String filename) {
        // 创建Excel的工作书册 Workbook,对应到一个excel文档
        HSSFWorkbook workbook = new HSSFWorkbook();

        // 创建Excel的工作sheet,对应到一个excel文档的sheet
        HSSFSheet sheet = workbook.createSheet("first");

        // 设置excel每列宽度
        sheet.setColumnWidth(0, 4000);
        sheet.setColumnWidth(1, 3500);

        // 创建字体样式
        HSSFFont font = workbook.createFont();
        font.setFontName("Verdana");
        font.setBold(true);
        font.setFontHeight((short) 300);
        font.setColor(HSSFColor.BLUE.index);

        // 创建单元格样式
        HSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(HSSFColor.LIGHT_TURQUOISE.index);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // 设置字体
        style.setFont(font);

        // 创建Excel的sheet的一行
        HSSFRow row = sheet.createRow(0);
        row.setHeight((short) 500); // 设定行的高度

        // 创建一个Excel的单元格
        HSSFCell cell = row.createCell(0);

        // 合并单元格(startRow，endRow，startColumn，endColumn)
        sheet.addMergedRegion(new CellRangeAddress(0, 0 , 0, 2));

        // 给Excel的单元格设置样式和赋值
        cell.setCellStyle(style);
        cell.setCellValue("Hello World");

        // 设置单元格内容格式
        HSSFCellStyle dataStyle = workbook.createCellStyle();
        dataStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("h:mm:ss"));
        dataStyle.setWrapText(true); // 自动换行

        // 设置边框
        dataStyle.setBottomBorderColor(HSSFColor.RED.index); // 设置底边宽颜色
        dataStyle.setBorderBottom(BorderStyle.THIN);
        dataStyle.setBorderLeft(BorderStyle.THIN);
        dataStyle.setBorderRight(BorderStyle.THIN);
        dataStyle.setBorderTop(BorderStyle.THIN);

        row = sheet.createRow(1);

        // 设置单元格的样式格式
        cell = row.createCell(0);
        cell.setCellStyle(dataStyle);
        cell.setCellValue(new Date());

        // 创建超链接
        HSSFHyperlink hyperlink = workbook.getCreationHelper().createHyperlink(HyperlinkType.URL);
        hyperlink.setAddress("http://josephwlf.com");

        cell = row.createCell(1);
        cell.setCellValue("个人主页");
        cell.setHyperlink(hyperlink); // 设定单元格的链接

        try {
            FileOutputStream outputStream = new FileOutputStream(filename);
            workbook.write(outputStream);

            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

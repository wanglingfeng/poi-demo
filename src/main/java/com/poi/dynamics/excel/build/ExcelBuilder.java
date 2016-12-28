package com.poi.dynamics.excel.build;

import com.poi.dynamics.excel.build.annotation.ExcelCellHeader;
import org.apache.poi.hssf.usermodel.*;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Date;
import java.util.List;

/**
 * Created by lfwang on 2016/12/27.
 */
public class ExcelBuilder<T> {

    public void generate(List<T> dataList, OutputStream outputStream)
            throws NoSuchMethodException, InvocationTargetException, IllegalAccessException, IOException {
        if (null == dataList || dataList.isEmpty()) return;

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();

        // 设置标题行
        HSSFRow headerRow = sheet.createRow(0);

        for (int rowIndex = 0; rowIndex < dataList.size(); rowIndex++) {
            HSSFRow row = sheet.createRow(rowIndex + 1);

            T t = dataList.get(rowIndex);

            Field[] fields = t.getClass().getDeclaredFields();

            for (int cellIndex = 0; cellIndex < fields.length; cellIndex++) {
                Field field = fields[cellIndex];

                if (field.toString().contains("static")) continue;

                // 填充标题行数据
                if (0 == rowIndex) {
                    ExcelCellHeader header = field.getAnnotation(ExcelCellHeader.class);

                    // 利用反射，根据javabean属性的先后顺序，动态调用getXxx()方法得到属性值
                    HSSFCell headerCell = headerRow.createCell(cellIndex);
                    headerCell.setCellValue(header.value());
                }

                String fileName = field.getName();
                String getMethodName = "get" + fileName.substring(0, 1).toUpperCase() + fileName.substring(1);

                Method getMethod = t.getClass().getMethod(getMethodName);
                Object value = getMethod.invoke(t);

                // 通过i创建cell
                HSSFCell cell = row.createCell(cellIndex);

                if (value instanceof Double) {
                    HSSFCellStyle numberStyle = workbook.createCellStyle();
                    numberStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00")); // 设置小数格式

                    cell.setCellStyle(numberStyle);
                    cell.setCellValue((Double) value);
                } else if (value instanceof Date) {
                    HSSFCellStyle dataStyle = workbook.createCellStyle();
                    dataStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-MM-dd hh:mm:ss")); // 设置日期格式

                    cell.setCellStyle(dataStyle);
                    cell.setCellValue((Date) value);
                } else {
                    cell.setCellValue(value.toString());
                }

                sheet.autoSizeColumn(cellIndex); // 设置列自动宽度
            }
        }

        workbook.write(outputStream);
    }
}

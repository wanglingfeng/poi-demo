package com.poi;

import com.poi.dynamics.excel.build.ExcelBuilder;
import com.poi.model.User;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Created by lfwang on 2016/12/27.
 */
public class MainRunner {

    public static void main(String... args) {
        /*BasicExcelBuilder builder = new BasicExcelBuilder();
        builder.generate("D:\\excel\\result.xls");*/

        ExcelBuilder<User> builder = new ExcelBuilder<>();
        List<User> list = new ArrayList<>();
        list.add(new User("Test", 123, new Date()));
        list.add(new User("Wang", 128, new Date()));

        try {
            FileOutputStream outputStream = new FileOutputStream("D:\\excel\\result.xls");
            builder.generate(list, outputStream);
        } catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException | IOException e) {
            e.printStackTrace();
        }
    }
}

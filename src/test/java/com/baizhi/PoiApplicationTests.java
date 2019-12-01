package com.baizhi;


import com.baizhi.entity.User;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@SpringBootTest(classes = PoiApplication.class)
@RunWith(SpringRunner.class)
public class PoiApplicationTests {

    @Test
    public void contextLoads() throws IOException {
        //创建工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();

        //获取单元格样式对象（一个单元格样式对象只能对一个样式进行设置）
        HSSFCellStyle cellStyle1 = workbook.createCellStyle();

        //设置字体居中
        cellStyle1.setAlignment(HorizontalAlignment.CENTER);

        //对于字体的调整可以和单元格样式同时设置
        //创建字体样式
        HSSFFont font = workbook.createFont();
        //加粗
        font.setBold(true);
        //颜色
        font.setColor(Font.COLOR_RED);
        //楷体
        font.setFontName("宋体");
        //斜体
        font.setItalic(true);
        //添加设置到单元格
        cellStyle1.setFont(font);


        //创建表名
        HSSFSheet sheet = workbook.createSheet("用户表");
        //创建行
        HSSFRow row = sheet.createRow(0);
        //设置列宽
        sheet.setColumnWidth(3, 15 * 256);
        //获取单元格样式1
        HSSFCellStyle cellStyle = workbook.createCellStyle();

        //获取时间样式
        HSSFDataFormat format = workbook.createDataFormat();
        //设置时间样式
        short format1 = format.getFormat("yyyy-MM-dd");

        //将时间样式设置进单元格当中
        cellStyle.setDataFormat(format1);

        String[] strings = {"主键", "年龄", "姓名", "生日"};
        for (int i = 0; i < strings.length; i++) {
            //创建列
            HSSFCell cell = row.createCell(i);
            cell.setCellStyle(cellStyle1);
            cell.setCellValue(strings[i]);
        }
       /* //创建列
        HSSFCell cell = row.createCell(0);
        //添加内容
        cell.setCellValue("野花");*/
        //写出数据
        User user = new User("1", "zj", "18", new Date());
        User user1 = new User("2", "zj", "18", new Date());
        User user2 = new User("3", "zj", "18", new Date());
        List<User> users = new ArrayList<>();
        users.add(user);
        users.add(user1);
        users.add(user2);

        for (int i = 0; i < users.size(); i++) {
            //创建行
            HSSFRow row1 = sheet.createRow(i + 1);
            //创建列
            //HSSFCell cell = row1.createCell(0);
           /* cell.setCellValue(users.get(0).getId());
            cell.setCellValue(users.get(0).getAge());
            cell.setCellValue(users.get(0).getAge());
            cell.setCellValue(users.get(0).getBir());*/
            row1.createCell(0).setCellValue(users.get(i).getId());
            row1.createCell(1).setCellValue(users.get(i).getAge());
            row1.createCell(2).setCellValue(users.get(i).getName());
            HSSFCell cell = row1.createCell(3);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(users.get(i).getBir());
            //row1.createCell(3).setCellValue(users.get(i).getBir());

        }

        workbook.write(new FileOutputStream(new File("E:/zj.xls")));

    }

}

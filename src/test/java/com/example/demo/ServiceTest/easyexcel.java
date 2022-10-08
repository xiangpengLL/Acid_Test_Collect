package com.example.demo.ServiceTest;

import com.alibaba.excel.util.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class easyexcel {
    @Test
    public void createexcel(){
        FileOutputStream out=null;
        XSSFWorkbook workbook=null;
        try {
            out=new FileOutputStream("D:\\桌面\\temp\\1.xlsx");
            workbook=new XSSFWorkbook();
            XSSFSheet sheet=workbook.createSheet("first");//创建sheet
            XSSFRow row=sheet.createRow(0);//在本sheet中操作第0行
            XSSFCell cellHeader;
            String[] title={"学号","姓名"};
            for (int i=0;i<title.length;i++){//添加表头
                cellHeader=row.createCell(i);
                cellHeader.setCellValue(new XSSFRichTextString(title[i]));
            }
            XSSFCell cell;
            for (int i=1;i<10;i++){//自第1行开始写入数据
                row=sheet.createRow(i);
                int index=5;
                cell=row.createCell(index++);
                cell.setCellValue(i);
            }
            workbook.write(out);//创建
        }catch (Exception e){
            System.out.println(e.getMessage());
        }
    }

    @Test
    public void createNullExcel(){
        FileOutputStream out=null;
        XSSFWorkbook workbook=null;
        try {
            out=new FileOutputStream("D:\\桌面\\temp\\10.xlsx");
            workbook=new XSSFWorkbook();
            XSSFSheet sheet=workbook.createSheet("用户信息");//创建sheet
            XSSFRow row=sheet.createRow(0);//在本sheet中操作第0行
            XSSFCell cellHeader;
            String[] title={"用户编号","姓名","性别","工资","入职时间"};
            for (int i=0;i<title.length;i++){//添加表头
                cellHeader=row.createCell(i);
                cellHeader.setCellValue(new XSSFRichTextString(title[i]));
            }
            workbook.write(out);//创建
        }catch (Exception e){
            System.out.println(e.getMessage());
        }
    }
}

package com.example.demo.ServiceTest;

import cn.hutool.poi.excel.ExcelUtil;
import com.alibaba.excel.EasyExcel;
import com.example.demo.Bean.AllType;
import com.example.demo.Bean.Basic;
import com.example.demo.Bean.Person;
import com.example.demo.Bean.Person_;
import jdk.internal.org.xml.sax.XMLReader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.xb.xsdschema.All;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

public class PoiTest {
    public List<Person_> createData(){
        List<Person_> list=new ArrayList<>();
        list.add(new Person_(1007,"杨玲玲","女"));
        list.add(new Person_(1008,"李丽萍","女"));
        list.add(new Person_(1009,"Jack","男"));
        list.add(new Person_(1010,"tom","男"));
        return list;
    }
    @Test
    public void appendRows() throws Exception {
        List<Person_> list=createData();
        File file=new File("D:\\桌面\\temp\\6.xlsx");
        FileInputStream inputStream=new FileInputStream("D:\\桌面\\temp\\6.xlsx");
        XSSFWorkbook workbook=new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        int lastRow=sheet.getLastRowNum()+1;
        System.out.println("最后一行为："+lastRow);
        Row header=sheet.getRow(0);
        Cell cell=header.getCell(0);
        CellStyle headerStyle = cell.getCellStyle();
        int nullCell=getNullCell(header);
        Cell newCell= header.createCell(nullCell);
        newCell.setCellValue("新增列");
        newCell.setCellStyle(headerStyle);
        FileOutputStream outputStream=new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();
    }

    public AllType getCellValue(Cell cell){
        CellType cellType=cell.getCellType();
        AllType type=new AllType();
        switch (cellType){
            case STRING://值为字符串
                type.setString(cell.getStringCellValue());
                break;
            case BOOLEAN://值为布尔类型
                type.setBoolean(cell.getBooleanCellValue());
                break;
            case NUMERIC://值为数值类型
                type.setDouble(cell.getNumericCellValue());
                break;
            case BLANK://值为空
                type.setIsBlank(true);
                break;
            default://未定义类型
                type.setUndefine(true);
                break;
        }
        return type;
    }

    public int getNullCell(Row row){//获取空列
        int num=0;
        for (int i=0;;i++){
            if (row.getCell(i)==null){
                num=i;
                break;
            }
        }
        return num;
    }

    @Test
    public void optionNullExcel(){//
        String fileName="D:\\桌面\\temp\\20.xlsx";
        EasyExcel.write(fileName, Basic.class).sheet("9.28返校").doWrite(null);
    }

    public void creatNullExcel(){
        String filePath="D:\\桌面\\temp\\20.xlsx";
        File file=new File(filePath);
        try {
            if (file.exists()) file.delete();
            file.createNewFile();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

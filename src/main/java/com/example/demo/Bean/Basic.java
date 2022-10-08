package com.example.demo.Bean;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.ToString;

@Data
@AllArgsConstructor
@NoArgsConstructor
@ToString
public class Basic {
    @ExcelProperty("班级")
    private String classes;//班级
    @ExcelProperty("学号")
    private String number;//学号
    @ExcelProperty("姓名")
    private String name;//姓名
}

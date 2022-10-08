package com.example.demo.Bean;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.ToString;

import java.util.Date;

@Data
@AllArgsConstructor
@NoArgsConstructor
@ToString
public class Person {
    @ExcelProperty("提交时间（自动）")
    private Date submitDate;
    @ExcelProperty("学号（必填）")
    private String code;
    @ExcelProperty("姓名（必填）")
    private String name;
    @ExcelProperty("返校时间（必填）")
    private  Date backTime;
    @ExcelProperty("核酸检测截屏（必填）")
    private String imgUrl;
    @ExcelProperty("提交者（自动）")
    private String submitName;
}

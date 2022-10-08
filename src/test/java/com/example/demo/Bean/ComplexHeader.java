package com.example.demo.Bean;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.ToString;

import java.util.Date;
@Data
@ToString
@AllArgsConstructor
@NoArgsConstructor
public class ComplexHeader {
    @ExcelProperty({"用户主题1", "用户编号"})
    private Integer userId;
    @ExcelProperty({"用户主题1", "用户名称"})
    private String userName;
    @ExcelProperty({"用户主题2", "用户入职时间"})
    @DateTimeFormat("yyyy年MM月dd日")
    private Date hirDate;
}

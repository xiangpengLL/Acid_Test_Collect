package com.example.demo.Bean;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.ToString;

import java.util.Date;
@Data
@ToString
@AllArgsConstructor
@NoArgsConstructor
public class User {
    @ExcelProperty("用户编号")//描述表头
    private Integer userId;
    @ExcelProperty("姓名")
    private String userName;
    @ExcelProperty("性别")
    private String gender;
    @ExcelProperty("工资")
    private Double salary;
    @ExcelProperty("入职时间")
    private Date hirDate;
}

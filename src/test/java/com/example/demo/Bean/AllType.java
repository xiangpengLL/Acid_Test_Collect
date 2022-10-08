package com.example.demo.Bean;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.ToString;

import java.util.Date;
@Data
@AllArgsConstructor
@NoArgsConstructor
@ToString
public class AllType {
    private Double Double;//数值类型
    private String string;//字符串类型
    private Boolean Boolean;//布尔类型
    private Boolean isBlank;//空类型
    private Boolean undefine;//未被定义
}

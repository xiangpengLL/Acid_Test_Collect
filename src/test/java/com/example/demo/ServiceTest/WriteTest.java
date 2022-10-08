package com.example.demo.ServiceTest;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.read.metadata.ReadSheet;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.example.demo.Bean.ComplexHeader;
import com.example.demo.Bean.Person_;
import com.example.demo.Bean.User;
import org.junit.jupiter.api.Test;

import java.util.*;

public class WriteTest {
    @Test//创建并写
    public void test01(){
        String fileName="D:\\桌面\\temp\\6.xlsx";
        List<User> users=new ArrayList<>();
        User user1=new User(1001,"张三","男",10670.1,new Date());
        User user2=new User(1002,"李四","男",10000.1,new Date());
        User user3=new User(1003,"王五","男",10500.1,new Date());
        User user4=new User(1004,"赵六","女",10000.78,new Date());
        User user5=new User(1005,"李七","男",13005.1,new Date());
        User user6=new User(1006,"哈哈","女",467.12,new Date());
        users.add(user1);
        users.add(user2);
        users.add(user3);
        users.add(user4);
        users.add(user5);
        users.add(user6);
        EasyExcel.write(fileName,User.class).sheet("用户信息").doWrite(users);//若该excel不存在那么会自动创建
    }
    @Test//创建并写
    public void test02(){
        String fileName="D:\\桌面\\temp\\8.xlsx";
        List<User> users=new ArrayList<>();
        User user1=new User(1001,"张三","男",10670.1,new Date());
        User user2=new User(1002,"李四","男",10000.1,new Date());
        User user3=new User(1003,"王五","男",10500.1,new Date());
        User user4=new User(1004,"赵六","女",10000.78,new Date());
        User user5=new User(1005,"李七","男",13005.1,new Date());
        User user6=new User(1006,"哈哈","女",467.12,new Date());
        users.add(user1);
        users.add(user2);
        users.add(user3);
        users.add(user4);
        users.add(user5);
        users.add(user6);
        ExcelWriter writer = EasyExcel.write(fileName, User.class).build();//创建写对象
        WriteSheet writeSheet = EasyExcel.writerSheet("用户信息").build();//构建sheet信息
        writer.write(null,writeSheet);
        writer.finish();//关闭
    }
    @Test//写入部分列
    public void test03(){
        String fileName="D:\\桌面\\temp\\6.xlsx";
        List<User> users=new ArrayList<>();
        User user1=new User(1001,"张三","男",10670.1,new Date());
        User user2=new User(1002,"李四","男",10000.1,new Date());
        User user3=new User(1003,"王五","男",10500.1,new Date());
        User user4=new User(1004,"赵六","女",10000.78,new Date());
        User user5=new User(1005,"李七","男",13005.1,new Date());
        User user6=new User(1006,"哈哈","女",467.12,new Date());
        users.add(user1);
        users.add(user2);
        users.add(user3);
        users.add(user4);
        users.add(user5);
        users.add(user6);
        Set<String> set=new HashSet<>();//排除某个信息写入
        set.add("hirDate");
        set.add("salary");
        EasyExcel.write(fileName,User.class)
                .excludeColumnFiledNames(set)
                .sheet("用户信息01").doWrite(users);
    }
    @Test//复杂表头写入
    public void test04(){
        String fileName="D:\\桌面\\temp\\9.xlsx";
        List<ComplexHeader> complexHeaders=new ArrayList<>();
        ComplexHeader complexHeader1=new ComplexHeader(1001,"李磊",new Date());
        ComplexHeader complexHeader2=new ComplexHeader(1002,"张三",new Date());
        ComplexHeader complexHeader3=new ComplexHeader(1003,"李四",new Date());
        complexHeaders.add(complexHeader1);
        complexHeaders.add(complexHeader2);
        complexHeaders.add(complexHeader3);
        EasyExcel.write(fileName,ComplexHeader.class).sheet("复杂头写入").doWrite(complexHeaders);
    }

    /*
    * 读
    * */
    @Test
    public void test05(){//6
        String fileName="D:\\桌面\\temp\\6.xlsx";
        EasyExcel.read(fileName, Person_.class, new AnalysisEventListener<Person_>() {
            @Override//每解析一行数据就调用一次
            public void invoke(Person_ person, AnalysisContext analysisContext) {
                System.out.println("数据为:"+person.toString());
            }

            @Override//解析完成调用
            public void doAfterAllAnalysed(AnalysisContext analysisContext) {
                System.out.println("解析完成");
            }
        }).sheet().doRead();
    }
    @Test
    public void test06(){
        String fileName="D:\\桌面\\temp\\6.xlsx";
        ExcelReader reader = EasyExcel.read(fileName, Person_.class, new AnalysisEventListener<Person_>() {
            @Override//每解析一行数据就调用一次
            public void invoke(Person_ person, AnalysisContext analysisContext) {
                System.out.println("数据为:"+person.toString());
            }

            @Override//解析完成调用
            public void doAfterAllAnalysed(AnalysisContext analysisContext) {
                System.out.println("解析完成");
            }
        }).build();
        ReadSheet readSheet = EasyExcel.readSheet(0).build();
        reader.read(readSheet);
        reader.finish();
    }
    @Test
    public void test07(){
        String fileName="D:\\桌面\\temp\\6.xlsx";
        ExcelReader excelReader = EasyExcel.read(fileName).build();
        //读取第0个sheet
        ReadSheet readSheet = EasyExcel.readSheet(0)
                .head(Person_.class)
                .registerReadListener(new AnalysisEventListener<Person_>() {
                    @Override
                    public void invoke(Person_ person, AnalysisContext analysisContext) {

                    }

                    @Override
                    public void doAfterAllAnalysed(AnalysisContext analysisContext) {

                    }
                })
                .build();
    }

    /*
    * 填充
    * */
}

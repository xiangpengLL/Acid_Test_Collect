package com.example.demo.Controller;

import com.example.demo.Service.MainService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;

@RestController
public class MyController {
    @Autowired
    private MainService mainService;
    @RequestMapping("/process")
    public String index(String handExcel,String holdPosition,String imgPosition){
        System.out.println("访问完成");
        System.out.println("handExcel："+handExcel);
        System.out.println("holdPosition:"+holdPosition);
        System.out.println("imgPosition："+imgPosition);
        System.out.println(mainService.collectImg(handExcel,holdPosition,imgPosition));
        return "123";
    }
}

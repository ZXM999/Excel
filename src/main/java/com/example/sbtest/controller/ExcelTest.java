package com.example.sbtest.controller;

import com.example.sbtest.service.ExcelTestService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping
public class ExcelTest {
    @Autowired
    ExcelTestService excelTestService;
    @GetMapping("/test")
    public void test(){
        excelTestService.makeExcel();
    }
}

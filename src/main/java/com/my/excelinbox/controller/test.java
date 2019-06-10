package com.my.excelinbox.controller;

import com.my.excelinbox.excel.ReadExcel;
import com.my.excelinbox.service.ExcelService;
import com.my.excelinbox.service.Student;
import io.netty.util.internal.StringUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.util.Arrays;
import java.util.List;

@Controller
public class test {

    private ExcelService excelService;

    public test(ExcelService excelService) {
        this.excelService = excelService;
    }

    @PostMapping("/excel")
    public List<Student> addTrades(@RequestParam("file") MultipartFile multipartFile) {

        if (StringUtil.isNullOrEmpty(multipartFile.getOriginalFilename())) {
            throw new RuntimeException("empty file name!");
        }

        String suffixName = Arrays.asList(multipartFile.getOriginalFilename().split("\\.")).get(1);
        if (!Arrays.asList("xlsx", "xls").contains(suffixName)) {
            throw new RuntimeException("unsupported file type:" + suffixName);
        }

        try {
            return ReadExcel.getObjects(new XSSFWorkbook(multipartFile.getInputStream()), Student.class);
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return null;
    }

    @GetMapping("/getExcel")
    public ResponseEntity<byte[]> excelDownload() {

    }
}

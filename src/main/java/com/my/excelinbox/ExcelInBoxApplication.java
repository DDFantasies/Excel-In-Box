package com.my.excelinbox;

import org.springframework.boot.SpringApplication;

import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.jdbc.DataSourceAutoConfiguration;

@SpringBootApplication(exclude= {DataSourceAutoConfiguration.class})
public class ExcelInBoxApplication {

    public static void main(String[] args) {
        SpringApplication.run(ExcelInBoxApplication.class, args);
    }

}


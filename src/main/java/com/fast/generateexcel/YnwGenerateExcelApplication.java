package com.fast.generateexcel;

import com.fast.generateexcel.generate.ExcelWriterOptimized;
import com.fast.generateexcel.generate.ExcelWriterOptimizedQQ;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class YnwGenerateExcelApplication {

    public static void main(String[] args) throws InterruptedException {
//        SpringApplication.run(YnwGenerateExcelApplication.class, args);
//        ExcelWriterOptimized.wyy();
        ExcelWriterOptimizedQQ.qq();
    }

}

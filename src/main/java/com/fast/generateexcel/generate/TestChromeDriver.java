package com.fast.generateexcel.generate;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import java.util.concurrent.TimeUnit;

public class TestChromeDriver {
    public static void main(String[] args) {
        // 设置ChromeDriver路径
        System.setProperty("webdriver.chrome.driver", "D:\\All_File_storage\\PROGRESS\\IdeaObject\\YNW-generateExcel\\chromedriver-win64\\chromedriver.exe");

        // 创建ChromeOptions实例
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--headless"); // 可选：如果你想在后台运行浏览器

        // 创建WebDriver实例
        WebDriver driver = new ChromeDriver(options);
        driver.manage().timeouts().pageLoadTimeout(5, TimeUnit.MINUTES); // 设置页面加载超时时间

        try {
            // 打开一个测试网页
            driver.get("https://www.google.com");
            System.out.println("页面标题csss是: " + driver.getTitle());


        } finally {
            // 关闭浏览器
            driver.quit();
        }
    }
}

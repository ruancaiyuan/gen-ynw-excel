package com.fast.generateexcel.generate;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.TimeUnit;

public class ExcelWriterOptimizedQQ {

    public static void main(String[] args) throws InterruptedException {
        qq();
    }

    public static void qq() throws InterruptedException {
        // Excel文件路径
        String excelFilePath = "C:\\Users\\Administrator\\Desktop\\YNW_MUSIC_EXCEL\\urlsqq.xlsx";
        // 读取excel文件的数据的第一列
        List<String> urls = readExcelFile(excelFilePath);

        // 配置ChromeOptions
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--disable-infobars");
        options.addArguments("--start-maximized");
        options.addArguments("--disable-extensions");
        options.addArguments("--disable-gpu");
//        options.addArguments("--headless"); // 如果需要无头模式可以启用

        System.setProperty("webdriver.chrome.driver", "D:\\All_File_storage\\PROGRESS\\IdeaObject\\YNW-generateExcel\\chromedriver-win64\\chromedriver.exe");

        // 创建WebDriver实例
        // 创建WebDriver实例
        WebDriver driver = new ChromeDriver(options);

        // 设置超时
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        driver.manage().timeouts().pageLoadTimeout(60, TimeUnit.SECONDS);
        driver.manage().timeouts().setScriptTimeout(30, TimeUnit.SECONDS);

        // 打开 QQ 音乐登录页面
        driver.get("https://y.qq.com/");

        // 查找并点击登录按钮
        WebElement loginButton = driver.findElement(By.cssSelector("div.header__opt span.mod_top_login a.top_login__link"));
        loginButton.click();

        // 延迟等待登录 休眠20秒。

        System.out.println("手机扫码快捷登录……25秒");
        TimeUnit.SECONDS.sleep(25);

        List<String[]> dataList = new ArrayList<>();

        for (int i = 0; i < urls.size(); i++) {
            System.out.println("正在处理第 " + (i + 1) + " / " + urls.size() + " 个链接：" + urls.get(i));
            driver.get(urls.get(i));
            // 获取页面源代码
            String pageSource = driver.getPageSource();
            // 使用Jsoup解析页面源代码
            Document doc = Jsoup.parse(pageSource);
            String[] data = getArtistData(i, doc);
            data[5] = urls.get(i);
            dataList.add(data);
        }

        LocalDateTime currentTime = LocalDateTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
        String name = "qq_music_info" + currentTime.format(formatter) + ".xlsx";

        // 生成excel
        String filePath = "C:\\Users\\Administrator\\Desktop\\YNW_MUSIC_EXCEL\\" + name;

        //生成excel
        //每行行高25
        //第一行的标题：设置天蓝色背景，列宽分别为15，21，21，20，20，50，
        String[] title = {"序列", "歌曲名称", "专辑名称", "歌手名", "发行时间", "专辑链接"};
        // 列宽度
        int[] columnWidths = {15 * 256, 50 * 256, 25 * 256, 25 * 256, 25 * 256, 80 * 256};

        // 写入Excel
        try (Workbook workbook = new XSSFWorkbook(); FileOutputStream fileOut = new FileOutputStream(filePath)) {
            // 创建Sheet
            Sheet sheet = workbook.createSheet("歌曲信息");

            // 设置标题行的样式
            Font titleFont = workbook.createFont();
            titleFont.setFontName("宋体");
            titleFont.setFontHeightInPoints((short) 16); // 设置字体大小为16
            titleFont.setBold(true); // 加粗

            CellStyle titleStyle = workbook.createCellStyle();
            titleStyle.setFont(titleFont);

            // 设置数据行的样式
            Font dataFont = workbook.createFont();
            dataFont.setFontName("宋体");
            dataFont.setFontHeightInPoints((short) 15); // 设置字体大小为15

            CellStyle dataStyle = workbook.createCellStyle();
            dataStyle.setFont(dataFont);

            // 设置标题行
            Row titleRow = sheet.createRow(0);
            titleRow.setHeightInPoints(23); // 设置行高为23
            for (int i = 0; i < title.length; i++) {
                Cell cell = titleRow.createCell(i);
                cell.setCellValue(title[i]);
                cell.setCellStyle(titleStyle); // 应用样式
                // 设置列宽度
                sheet.setColumnWidth(i, columnWidths[i]);
            }

            // 写入数据行
            for (int rowIndex = 0; rowIndex < dataList.size(); rowIndex++) {
                Row dataRow = sheet.createRow(rowIndex + 1);
                dataRow.setHeightInPoints(25); // 设置行高为25
                String[] rowData = dataList.get(rowIndex);
                for (int cellIndex = 0; cellIndex < rowData.length; cellIndex++) {
                    Cell cell = dataRow.createCell(cellIndex);
                    cell.setCellValue(rowData[cellIndex]);
                    cell.setCellStyle(dataStyle); // 应用样式
                }
            }

            // 写入文件
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // 关闭WebDriver
            if (driver != null) {
                driver.quit();
            }
        }
        System.out.println("Excel文件已成功写入：" + filePath);
    }

    // 获取艺术家数据
    public static String[] getArtistData(int i, Document doc) {
        String[] data = new String[6];
        data[0] = String.valueOf(i + 1);
        // 提取歌曲名称
        Element songNameElement = doc.selectFirst("h1.data__name_txt");
        data[1] = songNameElement != null ? songNameElement.attr("title") : "";

        // 提取第二个 "所属专辑" 标签中的内容
        Element albumElement = doc.selectFirst("ul.data__info li a[title]");
        data[2] = albumElement != null ? albumElement.attr("title") : "";

        // 提取歌手名称
        Elements singerElements = doc.select("div.data__singer a.data__singer_txt");
        StringBuilder singers = new StringBuilder();
        for (Element singer : singerElements) {
            if (singers.length() > 0) {
                singers.append("、");
            }
            singers.append(singer.attr("title"));
        }
        data[3] = singers.toString();

        Element releaseDateElement = doc.selectFirst("li:contains(发行时间)");
        data[4] = releaseDateElement.text().split("：")[1].trim(); //发行时间
        return data;
    }

    //传输一个文件路径
    public static List<String> readExcelFile(String excelFilePath) {
        List<String> urls = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // 获取第一个工作表
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell cell = row.getCell(0); // 获取第一列
                if (cell != null) {
                    urls.add(cell.getStringCellValue());
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return urls;
    }

}

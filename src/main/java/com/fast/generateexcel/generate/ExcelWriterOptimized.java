package com.fast.generateexcel.generate;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.openqa.selenium.WebDriver;
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

public class ExcelWriterOptimized {

    public static void main(String[] args) {
        wyy();
    }

    public static void wyy() {
        // 设置ChromeDriver路径
//        System.setProperty("webdriver.chrome.driver", "C:\\Users\\Administrator\\Desktop\\YNW_MUSIC_EXCEL\\chromedriver-win64\\chromedriver.exe");
        System.setProperty("webdriver.chrome.driver", "D:\\All_File_storage\\PROGRESS\\IdeaObject\\YNW-generateExcel\\chromedriver-win64\\chromedriver.exe");

        // Excel文件路径
        String excelFilePath = "C:\\Users\\Administrator\\Desktop\\YNW_MUSIC_EXCEL\\urls.xlsx";
        // 读取excel文件的数据的第一列
        List<String> urls = readExcelFile(excelFilePath);

        // 配置ChromeOptions
        ChromeOptions options = new ChromeOptions();
//        options.addArguments("--headless");  // 可选：如果你不需要显示浏览器窗口

        // 创建WebDriver实例
        WebDriver driver = new ChromeDriver(options);
        // 设置页面加载超时时间为 5 分钟
        driver.manage().timeouts().pageLoadTimeout(5, TimeUnit.MINUTES);

        // 设置异步脚本执行超时时间为 30 秒
        driver.manage().timeouts().setScriptTimeout(30, TimeUnit.SECONDS);

        List<String[]> dataList = new ArrayList<>();
        for (int i = 0; i < urls.size(); i++) {
            System.out.println("正在处理第 " + (i + 1) + " / " + urls.size() + " 个链接：" + urls.get(i));
            driver.get(urls.get(i));
            // 切换到iframe，因为网易云音乐页面嵌套了iframe
            driver.switchTo().frame("contentFrame");
            // 获取页面源代码
            String pageSource = driver.getPageSource();
            // 使用Jsoup解析页面源代码
            Document doc = Jsoup.parse(pageSource);
            String[] data = getArtistData(i, doc, driver);
            data[5] = urls.get(i);
            dataList.add(data);
        }
        LocalDateTime currentTime = LocalDateTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
        String name = "wyy_music_info" + currentTime.format(formatter) + ".xlsx";
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
    public static String[] getArtistData(int i, Document doc, WebDriver driver) {
        String[] data = new String[6];
        data[0] = String.valueOf(i + 1);
        // 提取歌曲名称
        Element songNameElement = doc.selectFirst("div.tit em.f-ff2");
        String name1 = songNameElement != null ? songNameElement.text() : "";

        Element songNameElement2 = doc.selectFirst("div.tit div.subtit.f-fs1.f-ff2");
        String name2 = (songNameElement2 != null && !songNameElement2.text().isEmpty()) ? "(" + songNameElement2.text() + ")" : "";

        data[1] = name1 + " " + name2;

        // 提取歌手名称
        Element artistElement = doc.selectFirst("p.des.s-fc4 span[title]");
        data[3] = artistElement != null ? artistElement.attr("title") : "";

        // 提取第二个 "所属专辑" 标签中的内容
        Element albumElement = doc.selectFirst("p.des.s-fc4:contains(所属专辑) a.s-fc7");
        data[2] = albumElement != null ? albumElement.text() : "";
        String albumHref = albumElement != null ? albumElement.attr("href") : "";
        if (!albumHref.isEmpty()) {
            albumHref = "https://music.163.com/#" + albumHref;
            driver.get(albumHref);
            // 切换到iframe
            driver.switchTo().frame("contentFrame");
            // 获取页面源代码
            String pageSource = driver.getPageSource();
            // 使用Jsoup解析页面源代码
            Document doc2 = Jsoup.parse(pageSource);
            // 提取发行时间
            Element releaseTime = doc2.select("p.intr").get(1);
            data[4] = releaseTime != null ? releaseTime.text().replaceFirst("^发行时间：", "") : "";
        } else {
            data[4] = "";
        }
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

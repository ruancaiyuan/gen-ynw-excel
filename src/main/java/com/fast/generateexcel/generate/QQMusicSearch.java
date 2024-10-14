package com.fast.generateexcel.generate;

import org.apache.http.HttpEntity;
import org.apache.http.HttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.IOException;

/**
 * @author ruan cai yuan
 * @version 1.0
 * @fileName com.fast.generateexcel.generate.QQMusicSearch
 * @description: TODO
 * @since 2024/7/7 16:03
 */
public class QQMusicSearch {

    public static void main(String[] args) {
        String songName = "你的歌名"; // 替换为你要查找的歌曲名称
        String searchUrl = "https://c.y.qq.com/soso/fcgi-bin/client_search_cp?p=1&n=10&w=" + songName;

        try {
            // 发送HTTP请求
            CloseableHttpClient httpClient = HttpClients.createDefault();
            HttpGet httpGet = new HttpGet(searchUrl);
            HttpResponse response = httpClient.execute(httpGet);
            HttpEntity entity = response.getEntity();

            if (entity != null) {
                String responseString = EntityUtils.toString(entity);
                Document doc = Jsoup.parse(responseString);

                // 解析HTML并查找相似歌曲链接
                Elements links = doc.select("a[href]"); // 假设相似歌曲链接在<a>标签中
                for (Element link : links) {
                    String linkHref = link.attr("href");
                    String linkText = link.text();
                    System.out.println(linkText + ": " + linkHref);
                }
            }

            httpClient.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


}

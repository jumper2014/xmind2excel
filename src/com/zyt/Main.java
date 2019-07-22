/*
主程序，将xmind文件转换成Excel文件
使用方法：java -jar xmind2excel.jar Test.xmind
 */
package com.zyt;

import java.io.*;
import java.util.List;


public class Main {

    public static void main(String[] args) {
        try {

            //获得xmind文件名
            String xmindFile = args[0];
            System.out.println("mind file is: " + xmindFile);

            //获得当前目录
            String xmlPath = System.getProperty("user.dir");
            String xmindFolderPath = System.getProperty("user.dir");

            // 调用unZip()进行解压
            File srcZipFile = new File(xmlPath, xmindFile);
            UnZipUtil.unZip(srcZipFile, xmindFolderPath);

            // 读取Xml文件，获取所有用例集合
            List<List<String>> allCaseList = ReadXml.readXml(xmindFolderPath);

            // 通过调用writeToExcel方法写入Excel
//            WriteToExcel.writeToExcel(allCaseList, xmindFolderPath);
            WriteToExcel.writeToExcelNew(allCaseList, xmindFolderPath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

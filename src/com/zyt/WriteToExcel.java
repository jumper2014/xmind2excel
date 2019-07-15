package com.zyt;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

// the Java API for Microsoft Documents
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;


/**
 * 将用例写入Excel
 */
public class WriteToExcel {

    /**
     * 将用例写入Excel
     *
     * @return
     */
    public static HSSFWorkbook writeToExcel(List<List<String>> allCaseList, String xmindFolderPath) {

        int caseCount = 0;

        // 第一步：创建Excel工作簿对象
        HSSFWorkbook workbook = new HSSFWorkbook();

        // 第二步：创建工作表
        HSSFSheet sheet = workbook.createSheet("测试用例");

        // 第三步：在sheet中添加表头第0行
        HSSFRow row = sheet.createRow(0);

        int testCasePathCol = 1;
        int testCaseNameCol = 2;
        int testCaseDescCol = 3;
        int executeDurationCol = 4;
        int testStepCol = 5;
        int expectionCol = 6;

        // 第四步:声明列对象
        HSSFCell cell1 = row.createCell(testCasePathCol - 1);
        HSSFCell cell2 = row.createCell(testCaseNameCol - 1);
        HSSFCell cell3 = row.createCell(testCaseDescCol - 1);
        HSSFCell cell4 = row.createCell(executeDurationCol - 1);
        HSSFCell cell5 = row.createCell(testStepCol - 1);
        HSSFCell cell6 = row.createCell(expectionCol - 1);

        cell1.setCellValue("测试案例路径");
        cell2.setCellValue("测试案例名称");
        cell3.setCellValue("测试案例描述");
        cell4.setCellValue("预计执行工时（分钟）");
        cell5.setCellValue("步骤描述");
        cell6.setCellValue("预期结果");


        // 设置表头样式
        HSSFCellStyle styleHead = workbook.createCellStyle();
        // 设置表头字体
        HSSFFont fontHead = workbook.createFont();

        cell1.setCellStyle(getHeadStyle(styleHead, fontHead));
        cell2.setCellStyle(getHeadStyle(styleHead, fontHead));
        cell3.setCellStyle(getHeadStyle(styleHead, fontHead));
        cell4.setCellStyle(getHeadStyle(styleHead, fontHead));
        cell5.setCellStyle(getHeadStyle(styleHead, fontHead));
        cell6.setCellStyle(getHeadStyle(styleHead, fontHead));

        // 设置列宽
        sheet.setColumnWidth(testCasePathCol - 1, 15 * 256);
        sheet.setColumnWidth(testCaseNameCol - 1, 30 * 256);
        sheet.setColumnWidth(testCaseDescCol - 1, 30 * 256);
        sheet.setColumnWidth(executeDurationCol - 1, 20 * 256);
        sheet.setColumnWidth(testStepCol - 1, 60 * 256);
        sheet.setColumnWidth(expectionCol - 1, 30 * 256);

        int caseNameIndex4Xml = 3;
        // 设置单元格样式
        HSSFCellStyle style = workbook.createCellStyle();
        // 设置单元格字体
        HSSFFont font = workbook.createFont();
        // 存储之前一个案例的名字，如果一样表示是同一个案例
        String preCase = "";

        // 遍历所有case集合
        for (int i = 0; i < allCaseList.size(); i++) {
//            System.out.println("-------------------");
//            System.out.println("Row: " + String.valueOf(i));
            // 创建用例内容的行，表头为第0行，因此真正的内容从i+1行开始
            row = sheet.createRow(i + 1);

            // 第一列为测试用例的路径
            HSSFCell cellTestCasePath = row.createCell(testCasePathCol - 1);

            // 取出单条用例
            List<String> caseList = allCaseList.get(i);
            String caseName = caseList.get(caseNameIndex4Xml);
            if (preCase.equals(caseName)) {
                cellTestCasePath.setCellValue("");
            } else {
                caseCount ++;
                cellTestCasePath.setCellValue("TestCasePath");
            }
            cellTestCasePath.setCellStyle(getCellStyle(style, font));

            // 取出每一个用例小步骤
            for (int j = 1; j < caseList.size(); j++) {
//                System.out.println("Column: " + String.valueOf(j));
                if (j <= 4) {
                    // 路径为第0列，因此用例其他信息从j+1列开始，按照顺序把前四个列写入
                    HSSFCell cell = row.createCell(j);
                    // 测试用例名和测试用例描述，保持一致，如果和上一个用例名一样，就留空
                    if (j == 1 || j == 2) {
//                        System.out.println("preCase:" + preCase + " nowCase:" + caseName);
                        if (preCase.equals(caseName)) {
                            cell.setCellValue("");
                        } else {
                            cell.setCellValue(caseName);
                            if (j == 2) {
                                preCase = caseName;
                            }
                        }
                    } else if (j == 3) {  // 执行工时这一行，留空
                        cell.setCellValue("");
                    } else {    // 测试步骤列
                        cell.setCellValue(caseList.get(j));
                    }
                    cell.setCellStyle(getCellStyle(style, font));

                }

                // 获取定位元素 预期结果 的下标,如果没有，则expect = -1
                int expect = caseList.indexOf("预期结果");
                // 填写预期结果
                if (expect != -1) {
                    HSSFCell cellExp = row.createCell(expectionCol - 1);
                    cellExp.setCellValue(caseList.get(expect + 1));
                    cellExp.setCellStyle(getCellStyle(style, font));
                }

            }
        }

        FileOutputStream out;

        try {
            // 生成文件路径: 当前目录
            String filePath = xmindFolderPath;
            // 文件名
            String fileName = allCaseList.get(0).get(0) + "Case.xls";

            // 生成excel文件
            out = new FileOutputStream(filePath + "/" + fileName);
            workbook.write(out);

            System.out.println("Transfer done！Path：" + filePath + fileName);
            out.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Total case number is: "+ String.valueOf(caseCount));

        return workbook;

    }

    /**
     * 设置表头格式 颜色可参照：https://blog.csdn.net/w405722907/article/details/76915903
     *
     * @param styleHead
     * @return
     */
    public static HSSFCellStyle getHeadStyle(HSSFCellStyle styleHead, HSSFFont fontHead) {

        // 水平居中
        styleHead.setAlignment(HorizontalAlignment.CENTER);
        // 垂直居中
        styleHead.setVerticalAlignment(VerticalAlignment.CENTER);

        // 设置标题背景色
        styleHead.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styleHead.setFillForegroundColor(IndexedColors.LIME.getIndex());// 绿色
//        style.setFillForegroundColor(IndexedColors.PALE_BLUE.index);// 蓝色

        // 设置四周边框
        styleHead.setBorderBottom(BorderStyle.THIN);// 下边框
        styleHead.setBorderLeft(BorderStyle.THIN);// 左边框
        styleHead.setBorderTop(BorderStyle.THIN);// 上边框
        styleHead.setBorderRight(BorderStyle.THIN);// 右边框

        // 设置自动换行;
        styleHead.setWrapText(true);

        // 设置字体
        fontHead.setFontName("微软雅黑");
        fontHead.setBold(true);
        styleHead.setFont(fontHead);

        // 自定义一个原谅色
//        HSSFPalette customPalette = workbook.getCustomPalette();
//        HSSFColor yuanLiangColor = customPalette.addColor((byte) 146, (byte) 208, (byte) 80);

        return styleHead;
    }

    /**
     * 设置单元格格式 颜色可参照：https://blog.csdn.net/w405722907/article/details/76915903
     * 对格式的设置进行优化，提升了性能：https://blog.csdn.net/qq592304796/article/details/52608714/
     *
     * @param style
     * @return
     */
    public static HSSFCellStyle getCellStyle(HSSFCellStyle style, HSSFFont font) {

        // 垂直居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        // 设置自动换行;
        style.setWrapText(true);

        font.setFontName("微软雅黑");
        style.setFont(font);

        return style;
    }

}

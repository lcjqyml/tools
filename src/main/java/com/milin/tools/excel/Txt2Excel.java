package com.milin.tools.excel;


import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Arrays;

/**
 *
 * @author changjun.liao
 */
public class Txt2Excel {

    public static void main(String[] args) throws IOException {
        readAndWrite();
    }

    private static void readAndWrite() throws IOException {
        System.out.println("Start");
        String folderStr = "D:\\work\\hy10086\\Docs\\大数据智能地图\\数据参考\\tab\\res\\";
        File folder = new File(folderStr);
        File[] files = folder.listFiles();
        if (ArrayUtils.isNotEmpty(files)) {
            ExcelWriter writer = new ExcelWriter(new FileOutputStream(new File(folderStr + "all.xls")));
            Arrays.stream(files).filter(file -> file.getName().endsWith(".txt")).forEach(file -> {
                try (BufferedReader headReader = new BufferedReader(new InputStreamReader(
                        new FileInputStream(file), "UTF-8"
                )); BufferedReader detailReader = new BufferedReader(
                             new InputStreamReader(
                                     new FileInputStream(
                                             new File(file.getAbsolutePath().replace(".txt", ""))
                                     ), "utf-8"
                             )
                     )) {
                    ExcelWriter.Sheet sheetAct = writer.getOrCreateSheet(file.getName().replace(".txt", "") + "样例数据");
                    String headStr;
                    int cellId = 0;
                    while ((headStr = headReader.readLine()) != null) {
                        if (StringUtils.isBlank(headStr)) {
                            continue;
                        }
                        sheetAct.getOrCreateCell(0, cellId++).setCellValue(headStr);
                    }
                    String detailStr;
                    int excelLine = 1;
                    while ((detailStr = detailReader.readLine()) != null) {
                        if (StringUtils.isBlank(detailStr)) {
                            continue;
                        }
                        String[] dataArr = detailStr.split("\\|");
                        for (int idx = 0; idx < dataArr.length; idx++) {
                            sheetAct.getOrCreateCell(excelLine, idx).setCellValue(dataArr[idx]);
                        }
                        excelLine++;
                    }
                } catch (IOException e) {
                    e.printStackTrace();
                }
            });
            writer.write();
        }
        System.out.println("End");
    }


    // public static void main(String[] args) throws IOException {
    // System.out.println("Start");
    // File file = new File("D:\\data_3452.txt");
    // File outFile = new File("D:\\data_3452.xls");
    // FileOutputStream out = new FileOutputStream(outFile);
    // BufferedReader reader = new BufferedReader(new FileReader(file));
    // ExcelWriter writer = new ExcelWriter(out);
    // ExcelWriter.Sheet sheetAct = writer
    // .getOrCreateSheet("data_3452");
    //
    // sheetAct.getOrCreateCell(0, 0).setCellValue("Email");
    // sheetAct.getOrCreateCell(0, 1).setCellValue("OrderId");
    // sheetAct.getOrCreateCell(0, 2).setCellValue("PurchaseDate");
    // sheetAct.getOrCreateCell(0, 3).setCellValue("Source");
    // String dataStr = null;
    // int excelLine = 1;
    // while ((dataStr = reader.readLine()) != null) {
    // if (StringUtils.isBlank(dataStr)) {
    // continue;
    // }
    // String[] dataArr = dataStr.split(",");
    // sheetAct.getOrCreateCell(excelLine, 0).setCellValue(dataArr[1]);
    // sheetAct.getOrCreateCell(excelLine, 1).setCellValue(dataArr[2]);
    // sheetAct.getOrCreateCell(excelLine, 2).setCellValue(dataArr[0]);
    // sheetAct.getOrCreateCell(excelLine, 3).setCellValue(dataArr[3]);
    // excelLine++;
    // }
    // writer.write();
    // reader.close();
    // writer.close();
    // System.out.println("End");
    // }

    // public static void main(String[] args) throws IOException {
    // File outFile = new
    // File("D:\\work\\fengxingsoftware\\share\\data\\3452_0901_multi_p.xls");
    // FileOutputStream out = new FileOutputStream(outFile);
    // ExcelWriter writer = new ExcelWriter(out);
    // FileInputStream fis = new
    // FileInputStream("D:\\work\\fengxingsoftware\\share\\data\\3452_0901.xls");
    // // 根据excel文件路径创建文件流
    // POIFSFileSystem fs = new POIFSFileSystem(fis); // 利用poi读取excel文件流
    // HSSFWorkbook wb = new HSSFWorkbook(fs); // 读取excel工作簿
    // for (int i = 0; i < 2; i++) {
    // HSSFSheet sheet = wb.getSheetAt(i); // 读取excel的sheet，0表示读取第一个、1表示第二个.....
    // String sName = sheet.getSheetName() + "_multi_psid";
    // ExcelWriter.Sheet sheetAct = writer.getOrCreateSheet(sName);
    // sheetAct.getOrCreateCell(0, 0).setCellValue("MasterId");
    // sheetAct.getOrCreateCell(0, 1).setCellValue("Email");
    // sheetAct.getOrCreateCell(0, 2).setCellValue("ExternalEmail");
    // sheetAct.getOrCreateCell(0, 3).setCellValue("testingItemId");
    // Map<String, String> emailPInfo = new HashMap<String, String>();
    // Set<String> multiEmail = new HashSet<String>();
    // // 获取sheet中总共有多少行数据sheet.getPhysicalNumberOfRows()
    // for (int rowIndex = 1; rowIndex < sheet.getPhysicalNumberOfRows();
    // rowIndex++) {
    // HSSFRow row = sheet.getRow(rowIndex); // 取出sheet中的某一行数据
    // if (row != null) {
    // HSSFCell cell = row.getCell(2);
    // String pInfo = cell.getStringCellValue();
    // if (!pInfo.contains("p:")) {
    // continue;
    // }
    // String email = row.getCell(1).getStringCellValue().trim().toLowerCase();
    // String psid = emailPInfo.get(email);
    // for (String info : pInfo.split(",")) {
    // if (info.contains("p:")) {
    // if (psid == null) {
    // emailPInfo.put(email, info.trim().toLowerCase());
    // } else {
    // if (!psid.equals(info.trim().toLowerCase())) {
    // multiEmail.add(email);
    // }
    // }
    // break;
    // }
    // }
    // }
    // }
    // int excelline = 1;
    // for (int rowIndex = 1; rowIndex < sheet.getPhysicalNumberOfRows();
    // rowIndex++) {
    // HSSFRow row = sheet.getRow(rowIndex); // 取出sheet中的某一行数据
    // if (row != null) {
    // String email = row.getCell(1).getStringCellValue().trim().toLowerCase();
    // if (multiEmail.contains(email)) {
    // sheetAct.getOrCreateCell(excelline,
    // 0).setCellValue(row.getCell(0).getStringCellValue());
    // sheetAct.getOrCreateCell(excelline,
    // 1).setCellValue(row.getCell(1).getStringCellValue());
    // sheetAct.getOrCreateCell(excelline,
    // 2).setCellValue(row.getCell(2).getStringCellValue());
    // sheetAct.getOrCreateCell(excelline,
    // 3).setCellValue(row.getCell(3).getStringCellValue());
    // excelline++;
    // }
    // }
    // }
    // }
    // writer.write();
    // writer.close();
    // fis.close();
    // }
}
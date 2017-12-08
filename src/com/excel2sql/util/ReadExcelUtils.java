package com.excel2sql.util;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ReadExcelUtils {

    private static XSSFWorkbook xssfWorkbook;
    private static XSSFSheet xssfSheet;
    private static XSSFRow xssfRow;

    public static List<Map<String, Object>> readExcelXlsx(InputStream is, int sheetIndex) {
        List<Map<String, Object>> content = new ArrayList<>();
        try {
            xssfWorkbook = new XSSFWorkbook(is);
        } catch (IOException e) {
        }
        xssfSheet = xssfWorkbook.getSheetAt(sheetIndex);
        // 得到总行数
        int rowNum = xssfSheet.getLastRowNum();
        System.out.println("rows -> " + rowNum);
        xssfRow = xssfSheet.getRow(0);
        int colNum = xssfRow.getPhysicalNumberOfCells();

        List<String> colList = new ArrayList<>();

        for (int c = 0; c < colNum; c++) {
            colList.add(xssfRow.getCell(c).toString());
        }

        for (int i = 1; i <= rowNum; i++) {
            xssfRow = xssfSheet.getRow(i);
            Map<String, Object> map = new HashMap<>();
            int j = 0;
            while (j < colNum) {
                map.put(colList.get(j), xssfRow.getCell(j));
                ++j;
            }
            content.add(map);
        }
        return content;
    }


    private static HSSFWorkbook hssfWorkbook;
    private static HSSFSheet hssfSheet;
    private static HSSFRow hssfRow;

    public static List<Map<String, Object>> readExcelXls(InputStream is, int sheetIndex) {
        List<Map<String, Object>> content = new ArrayList<>();
        try {
            hssfWorkbook = new HSSFWorkbook(is);
        } catch (IOException e) {
            e.printStackTrace();
        }
        hssfSheet = hssfWorkbook.getSheetAt(sheetIndex);
        // 得到总行数
        int rowNum = hssfSheet.getLastRowNum();
        System.out.println("rows -> " + rowNum);
        hssfRow = hssfSheet.getRow(0);
        int colNum = hssfRow.getPhysicalNumberOfCells();

        List<String> colList = new ArrayList<>();

        for (int c = 0; c < colNum; c++) {
            colList.add(hssfRow.getCell(c).toString());
        }

        for (int i = 1; i <= rowNum; i++) {
            hssfRow = hssfSheet.getRow(i);
            Map<String, Object> map = new HashMap<>();
            int j = 0;
            while (j < colNum) {
                map.put(colList.get(j), hssfRow.getCell(j));
                ++j;
            }
            content.add(map);
        }
        return content;
    }

}

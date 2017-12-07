package com.excel2sql.main;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Main {

    //project_id,source,related_name,related_id,related_url,status
    private static final String PROJECT_ID = "project_id";
    private static final String SOURCE = "source";
    private static final String RELATED_NAME = "related_name";
    private static final String RELATED_ID = "related_id";
    private static final String RELATED_URL = "related_url";
    private static final String STATUS = "status";

    private static final String FILENAME = "12-06匹配失败结果.xlsx";

    private static final String INSERT_TREE = "insert into `project_crawl_tree` (%s) VALUES(%s); ";
    private static final String UPDATE_TREE = "update `project_crawl_tree` set project_id = %s , status = %s where id = %s;";
    private static final String UPDATE_NAME = "update `project_name_crawl` set status = %s where id = %s;";

    private static final String EXCEL_PATH = "src/excel/";
    private static final String OUT_FILE_PATH = "src/out/";

    private static final String INSERT_TREE_FILE_NAME = "_insert_tree.sql";
    private static final String UPDATE_TREE_FILE_NAME = "_update_tree.sql";
    private static final String UPDATE_NAME_FILE_NAME = "_update_name.sql";

    private static final String XLS_SUFFIX = ".xls";
    private static final String XLSX_SUFFIX = ".xlsx";


    public static void main(String[] args) throws Exception {
        Main main = new Main();
        String excelPathComplete = EXCEL_PATH + FILENAME;
        main.nameMatchDeal(excelPathComplete);
    }

    public void nameMatchDeal(String completePath) throws Exception {
        File file = new File(completePath);
        InputStream fileInputStream = new FileInputStream(file);
        List<Map<String, Object>> resultList;

        String insetTreeSql = null;
        String updateTreeSql = null;
        String updateNameSql = null;
        if (completePath.contains(XLSX_SUFFIX)) {
            resultList = readExcelXlsx(fileInputStream, 0);
            Map<String, String> nameMatchResult = getNameMatchDealSql(resultList);
            if (nameMatchResult.get(TREE_KEY) != null) {
                insetTreeSql = nameMatchResult.get(TREE_KEY);

            }
            if (nameMatchResult.get(NAME_KEY) != null) {
                updateNameSql = nameMatchResult.get(NAME_KEY);

            }
            resultList = readExcelXlsx(fileInputStream, 1);
            updateTreeSql = getTreeDealSql(resultList);
        } else if (completePath.contains(XLS_SUFFIX)) {
            System.out.println("-------------" + XLS_SUFFIX);
            resultList = readExcelXls(fileInputStream, 0);
            Map<String, String> nameMatchResult = getNameMatchDealSql(resultList);
            if (nameMatchResult.get(TREE_KEY) != null) {
                insetTreeSql = nameMatchResult.get(TREE_KEY);

            }
            if (nameMatchResult.get(NAME_KEY) != null) {
                updateNameSql = nameMatchResult.get(NAME_KEY);

            }
            resultList = readExcelXlsx(fileInputStream, 1);
            updateTreeSql = getTreeDealSql(resultList);
        }

        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        String time = null;
        try {
            time = sdf.format(new Date());
        } catch (Exception e) {
            e.printStackTrace();
        }


        if (insetTreeSql != null) {
            write(time + INSERT_TREE_FILE_NAME, insetTreeSql);
        }
        if (updateNameSql != null) {
            write(time + UPDATE_NAME_FILE_NAME, updateNameSql);
        }
        if (updateTreeSql != null) {
            write(time + UPDATE_TREE_FILE_NAME, updateTreeSql);
        }

    }


    public void write(String fileName, String content) {
        String path = OUT_FILE_PATH + fileName;
        File f = new File(path);
        OutputStream out = null;
        try {
            out = new FileOutputStream(f);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        // 将字符串转成字节数组
        byte b[] = content.getBytes();
        try {
            // 将byte数组写入到文件之中
            out.write(b);
            out.close();
        } catch (IOException e1) {
            e1.printStackTrace();
        }
    }

    private static final String TREE_KEY = "treeKey";
    private static final String NAME_KEY = "updateKey";

    public Map<String, String> getNameMatchDealSql(List<Map<String, Object>> paramList) {

        Map<String, String> resultMap = new HashMap<>();

        int treeStatus = 1;

        StringBuffer insertTreeSb = new StringBuffer();
        StringBuffer updateNameSb = new StringBuffer();


        String insertSql = null;
        String updateSql = null;
        for (Map<String, Object> map : paramList) {
            String id = null;
            String projectId;
            String source;
            String relatedName;
            String relatedId;
            String relatedUrl;

            StringBuffer column = new StringBuffer();
            StringBuffer value = new StringBuffer();

            if (map.get("id") != null) {
                id = map.get("id").toString();
            }
            if (map.get("project_id") != null) {
                projectId = map.get("project_id").toString();
                column.append(PROJECT_ID);
                value.append(projectId);
            }
            if (map.get("source") != null) {
                source = map.get("source").toString();
                if (column.length() > 0) {
                    column.append(",").append(SOURCE);
                } else {
                    column.append(SOURCE);
                }
                if (value.length() > 0) {
                    value.append(",\"").append(source).append("\"");
                } else {
                    value.append("\"").append(source).append("\"");
                }
            }
            if (map.get("related_name") != null) {
                relatedName = map.get("related_name").toString();
                if (column.length() > 0) {
                    column.append(",").append(RELATED_NAME);
                } else {
                    column.append(RELATED_NAME);
                }
                if (value.length() > 0) {
                    value.append(",\"").append(relatedName).append("\"");
                } else {
                    value.append("\"").append(relatedName).append("\"");
                }
            }
            if (map.get("related_id") != null) {
                relatedId = map.get("related_id").toString();
                if (column.length() > 0) {
                    column.append(",").append(RELATED_ID);
                } else {
                    column.append(RELATED_ID);
                }
                if (value.length() > 0) {
                    value.append(",\"").append(relatedId).append("\"");
                } else {
                    value.append("\"").append(relatedId).append("\"");
                }
            }
            if (map.get("related_url") != null) {
                relatedUrl = map.get("related_url").toString();
                if (column.length() > 0) {
                    column.append(",").append(RELATED_URL);
                } else {
                    column.append(RELATED_URL);
                }
                if (value.length() > 0) {
                    value.append("").append(",\"").append(relatedUrl).append("\"");
                } else {
                    value.append("\"").append(relatedUrl).append("\"");
                }

            }
            if (SOURCE.equals(column.toString())) {
                continue;
            }

            if (column.length() > 0) {
                column.append(",").append(STATUS);
            } else {
                column.append(STATUS);
            }
            if (value.length() > 0) {
                value.append(",").append(treeStatus);
            } else {
                value.append(treeStatus);
            }


            insertSql = String.format(INSERT_TREE, column.toString(), value.toString());
            insertTreeSb.append(insertSql + "\n");
            if (id != null) {
                updateSql = String.format(UPDATE_NAME, "2", id);
                updateNameSb.append(updateSql + "\n");
            }

        }

        resultMap.put(TREE_KEY, insertTreeSb.toString());
        resultMap.put(NAME_KEY, updateNameSb.toString());

        return resultMap;
    }

    public String getTreeDealSql(List<Map<String, Object>> paramList) {

        StringBuffer stringBuffer = new StringBuffer();
        String status = "1";
        for (Map<String, Object> map : paramList) {
            String id;
            String projectId;

            if (map.get("project_id") == null) {
                continue;
            }
            if (map.get("id") == null) {
                continue;
            }
            projectId = map.get("project_id").toString();
            id = map.get("id").toString();


            String insertSql = String.format(UPDATE_TREE, projectId, status, id);
            stringBuffer.append(insertSql + "\n");

        }
        return stringBuffer.toString();
    }

    private POIFSFileSystem fs;
    private HSSFWorkbook hssfWorkbook;
    private HSSFSheet hssfSheet;
    private HSSFRow hssfRow;

    public List<Map<String, Object>> readExcelXls(InputStream is, int sheetIndex) {
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

    private XSSFWorkbook xssfWorkbook;
    private XSSFSheet xssfSheet;
    private XSSFRow xssfRow;

    public List<Map<String, Object>> readExcelXlsx(InputStream is, int sheetIndex) {
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


}

package com.excel2sql.main;

import com.excel2sql.util.FileIOUtils;
import com.excel2sql.util.ReadExcelUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class NameTreeMain {

    //project_id,source,related_name,related_id,related_url,status
    private static final String PROJECT_ID = "project_id";
    private static final String SOURCE = "source";
    private static final String RELATED_NAME = "related_name";
    private static final String RELATED_ID = "related_id";
    private static final String RELATED_URL = "related_url";
    private static final String STATUS = "status";

    private static final String FILENAME = "12-07匹配失败结果.xlsx";

    private static final String INSERT_TREE = "insert into `project_crawl_tree` (%s) VALUES(%s); ";
    private static final String UPDATE_TREE = "update `project_crawl_tree` set project_id = %s , status = %s where id = %s;";
    private static final String UPDATE_NAME = "update `project_name_crawl` set status = %s where id = %s;";

    private static final String EXCEL_PATH = "src/excel/";

    private static final String INSERT_TREE_FILE_NAME = "_insert_tree.sql";
    private static final String UPDATE_TREE_FILE_NAME = "_update_tree.sql";
    private static final String UPDATE_NAME_FILE_NAME = "_update_name.sql";

    private static final String XLS_SUFFIX = ".xls";
    private static final String XLSX_SUFFIX = ".xlsx";

    public static void main(String[] args) throws Exception {
        NameTreeMain nameTreeMain = new NameTreeMain();
        String excelPathComplete = EXCEL_PATH + FILENAME;
        nameTreeMain.nameMatchDeal(excelPathComplete);
    }

    public void nameMatchDeal(String completePath) throws Exception {
        File file = new File(completePath);
        InputStream fileInputStream = new FileInputStream(file);
        List<Map<String, Object>> resultList;

        String insetTreeSql = null;
        String updateTreeSql = null;
        String updateNameSql = null;
        if (completePath.contains(XLSX_SUFFIX)) {
            resultList = ReadExcelUtils.readExcelXlsx(fileInputStream, 0);
            Map<String, String> nameMatchResult = getNameMatchDealSql(resultList);
            if (nameMatchResult.get(TREE_KEY) != null) {
                insetTreeSql = nameMatchResult.get(TREE_KEY);

            }
            if (nameMatchResult.get(NAME_KEY) != null) {
                updateNameSql = nameMatchResult.get(NAME_KEY);

            }
            resultList = ReadExcelUtils.readExcelXlsx(fileInputStream, 1);
            updateTreeSql = getTreeDealSql(resultList);
        } else if (completePath.contains(XLS_SUFFIX)) {
            System.out.println("-------------" + XLS_SUFFIX);
            resultList = ReadExcelUtils.readExcelXls(fileInputStream, 0);
            Map<String, String> nameMatchResult = getNameMatchDealSql(resultList);
            if (nameMatchResult.get(TREE_KEY) != null) {
                insetTreeSql = nameMatchResult.get(TREE_KEY);

            }
            if (nameMatchResult.get(NAME_KEY) != null) {
                updateNameSql = nameMatchResult.get(NAME_KEY);

            }
            resultList = ReadExcelUtils.readExcelXlsx(fileInputStream, 1);
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
            FileIOUtils.fileWrite( time + INSERT_TREE_FILE_NAME, insetTreeSql);
        }
        if (updateNameSql != null) {
            FileIOUtils.fileWrite(time + UPDATE_NAME_FILE_NAME, updateNameSql);
        }
        if (updateTreeSql != null) {
            FileIOUtils.fileWrite(time + UPDATE_TREE_FILE_NAME, updateTreeSql);
        }

    }


    private static final String TREE_KEY = "treeKey";
    private static final String NAME_KEY = "updateKey";

    public Map<String, String> getNameMatchDealSql(List<Map<String, Object>> paramList) {

        Map<String, String> resultMap = new HashMap<>();

        int treeStatus = 1;

        StringBuffer insertTreeSb = new StringBuffer();
        StringBuffer updateNameSb = new StringBuffer();


        String insertSql;
        String updateSql;
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

            } else {
                continue;
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


}

package com.excel2sql.util;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class FileIOUtils {

    private static final String OUT_FILE_PATH = "src/out/";

    public static void fileWrite(String fileName, String content) {
        fileWrite(fileName,content,false);
    }

    public static void fileWrite(String fileName, String content, boolean isFullName) {
        String path;
        if(isFullName){
            path = fileName;
        }else{
            path = OUT_FILE_PATH + fileName;
        }
        File file = new File(path);
        OutputStream out = null;
        try {
            out = new FileOutputStream(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        byte b[] = content.getBytes();
        try {
            out.write(b);
            out.close();
        } catch (IOException e1) {
            e1.printStackTrace();
        }
    }

}

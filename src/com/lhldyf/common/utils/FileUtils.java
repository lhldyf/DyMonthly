package com.lhldyf.common.utils;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by LH on 17/6/27.
 */
public class FileUtils {

    static List<String> filePathlist = null;
    public static List<String> getFilePathList(String strPath) {
        getFilePathList(strPath,true);
        return filePathlist;
    }

    private static void getFilePathList(String strPath, boolean initArray) {
        if(initArray) {
            filePathlist = new ArrayList<>();
        }
        File dir = new File(strPath);
        File[] files = dir.listFiles(); // 该文件目录下文件全部放入数组
        if (files != null) {
            for (int i = 0; i < files.length; i++) {
                String fileName = files[i].getName();
                if (files[i].isDirectory()) { // 判断是文件还是文件夹
                    getFilePathList(files[i].getAbsolutePath(), false); // 获取文件绝对路径
                } else if (fileName.endsWith("xlsx")) { // 判断文件名是否以.xlsx结尾
                    String strFileName = files[i].getAbsolutePath();
                    if(fileName.indexOf("~$")>-1) {
                        System.out.println("文件以特殊字符~开头，将被认为临时文件，请主公明察，文件路径："+strFileName);
                        continue;
                    }
                    //System.out.println("---" + strFileName);
                    filePathlist.add(strFileName);
                } else {
                    continue;
                }
            }

        }
    }

}

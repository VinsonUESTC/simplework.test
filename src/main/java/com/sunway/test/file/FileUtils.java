/*
 * Copyright (c) 2018. www.sunway.com Edit by Vinson
 */

package com.sunway.test.file;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class FileUtils {

    public static void main(String[] args) {
        getAllFiles("/Users/vinson/Downloads/");
    }

    public static List<File> getAllFiles(String filepath){
        List<String> paths = new ArrayList<>();
        List<File> files = new ArrayList<>();
        paths = getAllFilePaths(new File(filepath),paths);
        for(String filepaths : paths){
            files.add(new File(filepaths));
            System.out.println(filepaths);
        }
        return files;
    }

    public static List<String> getAllFilePaths(File filePath,List<String> filePaths){
        File[] files = filePath.listFiles();
        if(files == null){
            return filePaths;
        }
        for(File f:files){
            if(f.isDirectory()){
                getAllFilePaths(f,filePaths);
            }else{
                filePaths.add(f.getPath());
            }
        }
        return filePaths;
    }
}

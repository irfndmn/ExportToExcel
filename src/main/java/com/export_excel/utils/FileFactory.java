package com.export_excel.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Component;
import org.springframework.util.ResourceUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

@Component
public class FileFactory {
    public static final String PATH_TEMPLATE;
    static {

        Path targetDirectory= Paths.get(System.getProperty("user.home") + File.separator + "Downloads");
        Path templateFolder = targetDirectory.resolve("Template");
        if(!Files.exists(templateFolder)){
            try {
                Files.createDirectory(templateFolder);
            } catch (IOException e) {
                System.out.println(""+e);
            }
        }
        PATH_TEMPLATE=System.getProperty("user.home") + File.separator + "Downloads" + File.separator + "Template" + File.separator;
    }


  // public static final String PATH_TEMPLATE="C:\\Users\\irfndmn\\Downloads\\Template\\";   For write clean and more dynamically code we use other way.

    public static File createFile(String fileName, Workbook workbook) throws Exception {
        OutputStream out = new FileOutputStream(PATH_TEMPLATE+fileName);
        workbook.write(out);
        return ResourceUtils.getFile(PATH_TEMPLATE+fileName);
    }
}

package com.export_excel.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Component;
import org.springframework.util.ResourceUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

@Component
public class FileFactory {

    public static final String PATH_TEMPLATE = System.getProperty("user.home") + File.separator + "Downloads" + File.separator + "Template" + File.separator;
  // public static final String PATH_TEMPLATE="C:\\Users\\iduman\\Downloads\\Template\\";   For write clean and more dynamically code we use other way.

    public static File createFile(String fileName, Workbook workbook) throws Exception {

        workbook = new HSSFWorkbook();

        OutputStream out = new FileOutputStream(PATH_TEMPLATE+fileName);

        workbook.write(out);

        return ResourceUtils.getFile(PATH_TEMPLATE+fileName);

    }











}

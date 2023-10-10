package com.export_excel.utils;

import com.export_excel.entity.ContactMessage;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;
import org.springframework.util.ResourceUtils;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.List;

import static com.export_excel.utils.FileFactory.PATH_TEMPLATE;

@Component
@Slf4j
public class ExcelUtils {

    public static ByteArrayInputStream exportContactMessage(List<ContactMessage> contactMessageList, String fileName) throws Exception {

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();


        // Get file -> not found -> create file
        File file;

        FileInputStream fileInputStream;

        try {
            file = ResourceUtils.getFile(PATH_TEMPLATE+fileName);
            fileInputStream=new FileInputStream(file);
        } catch (FileNotFoundException e) {
            log.info("FILE NOT FOUND");
            file = FileFactory.createFile(fileName,xssfWorkbook);
            fileInputStream=new FileInputStream(file);
        }


        // freeze pane
        XSSFSheet newSheet = xssfWorkbook.createSheet("sheet1");
        newSheet.createFreezePane(4,2,4,2);



        // create title font
        XSSFFont titleFont = xssfWorkbook.createFont();
        titleFont.setFontName("Arial");
        titleFont.setBold(true);
        titleFont.setFontHeightInPoints((short) 13);



        //create style for cell of title
        XSSFCellStyle titleCellStyle = xssfWorkbook.createCellStyle();
        titleCellStyle.setAlignment(HorizontalAlignment.CENTER);
        titleCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleCellStyle.setFillForegroundColor(IndexedColors.SKY_BLUE.index);
        titleCellStyle.setBorderBottom(BorderStyle.MEDIUM);
        titleCellStyle.setBorderLeft(BorderStyle.MEDIUM);
        titleCellStyle.setBorderRight(BorderStyle.MEDIUM);
        titleCellStyle.setBorderTop(BorderStyle.MEDIUM);
        titleCellStyle.setFont(titleFont);
        titleCellStyle.setWrapText(true);



        //font for data

        XSSFFont dataFont=xssfWorkbook.createFont();
        dataFont.setFontName("Arial");
        dataFont.setBold(false);
        dataFont.setFontHeightInPoints((short) 9);


        // create style for data
        XSSFCellStyle dataCellStyle = xssfWorkbook.createCellStyle();
        dataCellStyle.setAlignment(HorizontalAlignment.CENTER);
        dataCellStyle.setBorderBottom(BorderStyle.THIN);
        dataCellStyle.setBorderLeft(BorderStyle.THIN);
        dataCellStyle.setBorderRight(BorderStyle.THIN);
        dataCellStyle.setBorderTop(BorderStyle.THIN);
        dataCellStyle.setFont(dataFont);
        dataCellStyle.setWrapText(true);


        // insert field name as title to excel

        // insert data of the field to excel

        //return







    }


}

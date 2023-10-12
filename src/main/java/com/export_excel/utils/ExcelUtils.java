package com.export_excel.utils;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;
import org.springframework.util.ObjectUtils;
import org.springframework.util.ResourceUtils;

import java.io.*;
import java.lang.reflect.Field;
import java.util.List;

import static com.export_excel.utils.FileFactory.PATH_TEMPLATE;

@Component
@Slf4j
public class ExcelUtils {

    public static <T> ByteArrayInputStream exportContactMessage(List<T> theDataListForExport, String fileName) throws Exception {

        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        // Get file -> not found -> create file
        File file;
        FileInputStream fileInputStream;

        try {
            file = ResourceUtils.getFile(PATH_TEMPLATE + fileName);
            fileInputStream = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            log.info("FILE NOT FOUND");
            file = FileFactory.createFile(fileName, xssfWorkbook);
            fileInputStream = new FileInputStream(file);
        }

        // freeze pane
        XSSFSheet newSheet = xssfWorkbook.createSheet("sheet1");
        newSheet.createFreezePane(4, 2, 4, 2);

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

        XSSFFont dataFont = xssfWorkbook.createFont();
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

        Class<?> clazz = theDataListForExport.get(0).getClass();
        ExportConfig newExportConfig = ExportConfig.createExportConfig(clazz);

        insertFieldNameAsTitleToWorkbook(newExportConfig.getCellExportConfigList(), newSheet, titleCellStyle);

        // insert data of the field to excel
        insertDataToWorkbook(xssfWorkbook, newExportConfig, theDataListForExport, dataCellStyle);

        //return
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        xssfWorkbook.write(outputStream);
        outputStream.close();
        fileInputStream.close();

        return new ByteArrayInputStream(outputStream.toByteArray());

    }

    private static <T> void insertDataToWorkbook(Workbook workbook,
                                                 ExportConfig exportConfig,
                                                 List<T> dataList,
                                                 XSSFCellStyle dataCellStyle) {

        int startRowIndex = exportConfig.getStartRow();
        int sheetIndex = exportConfig.getSheetIndex();
        Class clazz = exportConfig.getDataClazz();
        List<CellConfig> cellConfigs = exportConfig.getCellExportConfigList();

        Sheet sheet = workbook.getSheetAt(sheetIndex);
        int currentRowIndex = startRowIndex;

        for (T data : dataList) {
            Row currentRow = sheet.getRow(currentRowIndex);
            if (ObjectUtils.isEmpty(currentRow)) {
                currentRow = sheet.createRow(currentRowIndex);
            }
            // insert data to row
            insertDataToCell(data, currentRow, cellConfigs, clazz, sheet, dataCellStyle);
            currentRowIndex++;
        }
    }


    private static void insertFieldNameAsTitleToWorkbook(List<CellConfig> cellConfigs,
                                                         Sheet sheet,
                                                         XSSFCellStyle titleCellStyle) {

        // title -> first row of  excel
        int currentRow = sheet.getTopRow();
        // create row
        Row row = sheet.createRow(currentRow);
        int i = 0;
        sheet.autoSizeColumn(currentRow);
        //insert field name to each cell
        for (CellConfig cellConfig : cellConfigs) {
            Cell currentCell = row.createCell(i);
            currentCell.setCellValue(cellConfig.getFieldName());
            currentCell.setCellStyle(titleCellStyle);
            sheet.autoSizeColumn(i);
            i++;
        }


    }


    private static <T> void insertDataToCell(T data, Row currentRow, List<CellConfig> cellConfigs,
                                             Class clazz,
                                             Sheet sheet,
                                             XSSFCellStyle dataStyle) {

        for (CellConfig cellConfig : cellConfigs) {
            Cell currentCell = currentRow.getCell(cellConfig.getColumnIndex());
            if (ObjectUtils.isEmpty(currentCell)) {
                currentCell = currentRow.createCell(cellConfig.getColumnIndex());
            }
            // get data for cell
            String cellValue = getCellValue(data, cellConfig, clazz);
            // set data
            currentCell.setCellValue(cellValue);
            sheet.autoSizeColumn(cellConfig.getColumnIndex());
            currentCell.setCellStyle(dataStyle);

        }


    }

    private static <T> String getCellValue(T data, CellConfig cellConfig, Class clazz) {

        String fieldName = cellConfig.getFieldName();
        try {
            Field field = getDeclaredField(clazz, fieldName);
            if (!ObjectUtils.isEmpty(field)) {
                field.setAccessible(true);
                return !ObjectUtils.isEmpty(field.get(data)) ? field.get(data).toString() : "";
            }

            return "";

        } catch (Exception e) {
            log.info("" + e);
            return "";
        }
    }


    private static Field getDeclaredField(Class clazz, String fieldName) {

        if (ObjectUtils.isEmpty(clazz) || ObjectUtils.isEmpty(fieldName)) {
            return null;
        }
        do {
            try {
                Field field = clazz.getDeclaredField(fieldName);
                field.setAccessible(true);
                return field;
            } catch (Exception e) {
                log.info("" + e);
            }
        } while ((clazz = clazz.getSuperclass()) != null);  // if super class not null we check to super class too for field.
        return null;
    }

}

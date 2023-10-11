package com.export_excel.utils;

import com.export_excel.entity.ContactMessage;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class ExportConfig {
    private int sheetIndex;
    private int startRow;
    private Class<?> dataClazz;
    private List<CellConfig> cellExportConfigList;

    public ExportConfig(Class<?> dataClazz) {
        this.sheetIndex = 0;
        this.startRow = 1;
        this.dataClazz = dataClazz;
        this.cellExportConfigList = generateCellConfigList(dataClazz);
    }

    public static ExportConfig createExportConfig(Class<?> dataClazz){
        return new ExportConfig(dataClazz);
    }



    public static List<CellConfig> generateCellConfigList(Class<?> clazz) {
        List<CellConfig> cellConfigList = new ArrayList<>();
        Field[] fields = clazz.getDeclaredFields();

        for (int i = 0; i < fields.length; i++) {
            String fieldName = fields[i].getName();
            cellConfigList.add(new CellConfig(i, fieldName));
        }

        return cellConfigList;
    }



    public static final ExportConfig contactMessageExport;

    static {

        // Initialize default values in a static block
        int defaultSheetIndex=0;
        int defaultStartRowIndex =1;

        // Create a default ExportConfig for ContactMessage class
        Class<?> defaultDataClazz = ContactMessage.class;
        List<CellConfig> defaultCellConfigList = generateCellConfigList(defaultDataClazz);

        // Create a default ExportConfig instance
        contactMessageExport = new ExportConfig(defaultSheetIndex,defaultStartRowIndex,defaultDataClazz,defaultCellConfigList);


    }


}

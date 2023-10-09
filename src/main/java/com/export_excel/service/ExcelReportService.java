package com.export_excel.service;

import com.export_excel.entity.ContactMessage;
import lombok.RequiredArgsConstructor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.stereotype.Service;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

@Service
@RequiredArgsConstructor
public class ExcelReportService {
    private final ContactMessageService contactMessageService;

    public void createExcel(HttpServletResponse response) throws IOException {

        HSSFWorkbook workbook = new HSSFWorkbook();
        Sheet sheet=workbook.createSheet("Contact Message Info");
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue("ID");
        row.createCell(1).setCellValue("DATE");
        row.createCell(2).setCellValue("NAME");
        row.createCell(3).setCellValue("EMAIL");
        row.createCell(4).setCellValue("MESSAGE");
        row.createCell(5).setCellValue("SUBJECT");
        List<ContactMessage> contactMessages=contactMessageService.getAllMessages();

        int rowIndex=1;
        for(ContactMessage message:contactMessages){
            Row rowData=sheet.createRow(rowIndex);
            rowData.createCell(0).setCellValue(message.getMessageId());
            rowData.createCell(1).setCellValue(message.getDate());
            rowData.createCell(2).setCellValue(message.getMessageName());
            rowData.createCell(3).setCellValue(message.getEmail());
            rowData.createCell(4).setCellValue(message.getMessage());
            rowData.createCell(5).setCellValue(message.getSubject());
            rowIndex++;
        }

        ServletOutputStream ops=response.getOutputStream();
        workbook.write(ops);
        workbook.close();
        ops.close();

    }

}

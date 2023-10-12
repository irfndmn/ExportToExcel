package com.export_excel.controller;

import com.export_excel.entity.ContactMessage;
import com.export_excel.repository.ContactMessageRepository;
import com.export_excel.service.ExcelReportService;
import com.export_excel.utils.ExcelUtils;
import com.export_excel.utils.ExportConfig;
import lombok.RequiredArgsConstructor;
import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;

@RestController
@RequiredArgsConstructor
@RequestMapping("/export")
public class ExcelReportController {
    private final ExcelReportService excelReportService;

    @GetMapping("/excelReport")
    public void generateExcelReport(HttpServletResponse response) throws IOException {

        response.setContentType("application/octet-stream");
        String headerKey = "Content-Disposition";
        String headerValue = "attachment;filename = Contact_Messages.xls";
        response.setHeader(headerKey,headerValue);
        excelReportService.createExcel(response);

    }

    @GetMapping("/export")
    public List<ContactMessage> exportData(HttpServletResponse response) {
        Object dataService;
        List<ContactMessage> data = null;
        return data;
    }


    @GetMapping("/exportContactMessage")
    public ResponseEntity<Resource> exportData() throws Exception {

        return excelReportService.exportContactMessage();

    }

}

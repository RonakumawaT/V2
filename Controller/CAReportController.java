package com.RK8.V2.Controller;
import com.RK8.V2.DTO.Gstr2BDTO;
import com.RK8.V2.DTO.PurchaseInvoiceDTO;
import com.RK8.V2.DTO.ReconciliationResult;
import com.RK8.V2.Parser.Gstr2BExcelParser;
import com.RK8.V2.Parser.PurchaseExcelParser;
import com.RK8.V2.Service.CAReportService;
import com.RK8.V2.Service.Purchase2BReconciliationService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

@RestController
@RequestMapping("/api/ca")
public class CAReportController {

    @Autowired
    private PurchaseExcelParser purchaseParser;

    @Autowired
    private Gstr2BExcelParser gstr2BParser;

    @Autowired
    private Purchase2BReconciliationService reconciliationService;

    @Autowired
    private CAReportService caReportService;

    @PostMapping("/generate-report")
    public ResponseEntity<ByteArrayResource> generateCAReport(
            @RequestParam("purchaseFile") MultipartFile purchaseFile,
            @RequestParam("gstr2bFile") MultipartFile gstr2bFile) {

        try {
            // Parse files
            List<PurchaseInvoiceDTO> purchases;
            List<Gstr2BDTO> gstr2bList;

            try (InputStream purchaseStream = purchaseFile.getInputStream();
                 InputStream gstr2bStream = gstr2bFile.getInputStream()) {

                purchases = purchaseParser.parse(purchaseStream);
                gstr2bList = gstr2BParser.parse(gstr2bStream);
            }

            // Reconcile
            List<ReconciliationResult> results = reconciliationService.reconcile(purchases, gstr2bList);

            // Generate CA report
            byte[] reportBytes = caReportService.generateCAReport(purchases, gstr2bList, results);

            // Create filename with timestamp
            String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
            String filename = String.format("GST_Reconciliation_Report_%s.xlsx", timestamp);

            // Prepare response
            ByteArrayResource resource = new ByteArrayResource(reportBytes);

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + filename + "\"")
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .contentLength(reportBytes.length)
                    .body(resource);

        } catch (Exception e) {
            return ResponseEntity.internalServerError().build();
        }
    }

        @PostMapping("/download-report")
    public ResponseEntity<ByteArrayResource> downloadReport(
            @RequestParam("purchaseFile") MultipartFile purchaseFile,
            @RequestParam("gstr2bFile") MultipartFile gstr2bFile) {

        try {
            // Parse files
            List<PurchaseInvoiceDTO> purchases;
            List<Gstr2BDTO> gstr2bList;

            try (InputStream purchaseStream = purchaseFile.getInputStream();
                 InputStream gstr2bStream = gstr2bFile.getInputStream()) {

                purchases = purchaseParser.parse(purchaseStream);
                gstr2bList = gstr2BParser.parse(gstr2bStream);
            }

            // Reconcile
            List<ReconciliationResult> results = reconciliationService.reconcile(purchases, gstr2bList);

            // Generate CA report
            byte[] reportBytes = caReportService.generateCAReport(purchases, gstr2bList, results);

            // Create filename with timestamp
            String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
            String filename = String.format("GST_Reconciliation_Report_%s.xlsx", timestamp);

            // Prepare response
            ByteArrayResource resource = new ByteArrayResource(reportBytes);

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + filename + "\"")
                    .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                    .contentLength(reportBytes.length)
                    .body(resource);

        } catch (Exception e) {
            return ResponseEntity.internalServerError().build();
        }
    }

    // Simple version without CA report service (if you just want mismatches)
    @PostMapping("/download-mismatches")
    public ResponseEntity<ByteArrayResource> downloadMismatchesExcel(
            @RequestParam("purchaseFile") MultipartFile purchaseFile,
            @RequestParam("gstr2bFile") MultipartFile gstr2bFile) {

        try {
            // Parse files
            List<PurchaseInvoiceDTO> purchases;
            List<Gstr2BDTO> gstr2bList;

            try (InputStream purchaseStream = purchaseFile.getInputStream();
                 InputStream gstr2bStream = gstr2bFile.getInputStream()) {

                purchases = purchaseParser.parse(purchaseStream);
                gstr2bList = gstr2BParser.parse(gstr2bStream);
            }

            // Reconcile
            List<ReconciliationResult> results = reconciliationService.reconcile(purchases, gstr2bList);

            // Generate simple Excel with mismatches
            byte[] excelBytes = generateSimpleMismatchExcel(results);

            // Create filename
            String filename = "GST_Mismatches_Report.xlsx";

            ByteArrayResource resource = new ByteArrayResource(excelBytes);

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + filename + "\"")
                    .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                    .contentLength(excelBytes.length)
                    .body(resource);

        } catch (Exception e) {
            return ResponseEntity.internalServerError().build();
        }
    }

    private byte[] generateSimpleMismatchExcel(List<ReconciliationResult> results) throws Exception {
        // Create workbook
        org.apache.poi.xssf.usermodel.XSSFWorkbook workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook();

        // Create sheet
        org.apache.poi.ss.usermodel.Sheet sheet = workbook.createSheet("Mismatches");

        // Create header row
        org.apache.poi.ss.usermodel.Row headerRow = sheet.createRow(0);
        String[] headers = {"Supplier GSTIN", "Invoice No", "Month", "Status",
                "Purchase Tax", "2B Tax", "ITC at Risk", "Remarks"};

        for (int i = 0; i < headers.length; i++) {
            org.apache.poi.ss.usermodel.Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }

        // Add data rows
        int rowNum = 1;
        for (ReconciliationResult result : results) {
            if (!result.getStatus().startsWith("MATCHED")) {
                org.apache.poi.ss.usermodel.Row row = sheet.createRow(rowNum++);

                row.createCell(0).setCellValue(result.getSupplierGstin());
                row.createCell(1).setCellValue(result.getInvoiceNo());
                row.createCell(2).setCellValue(result.getInvoiceMonth().toString());
                row.createCell(3).setCellValue(result.getStatus());
                row.createCell(4).setCellValue(result.getPurchaseTax().doubleValue());
                row.createCell(5).setCellValue(result.getGstr2bTax().doubleValue());
                row.createCell(6).setCellValue(result.getItcAtRisk().doubleValue());
                row.createCell(7).setCellValue(result.getRemarks());
            }
        }

        // Auto-size columns
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write to byte array
        try (java.io.ByteArrayOutputStream baos = new java.io.ByteArrayOutputStream()) {
            workbook.write(baos);
            return baos.toByteArray();
        } finally {
            workbook.close();
        }
    }
}
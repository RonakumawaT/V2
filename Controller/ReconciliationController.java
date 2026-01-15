package com.RK8.V2.Controller;

import com.RK8.V2.DTO.Gstr2BDTO;
import com.RK8.V2.DTO.PurchaseInvoiceDTO;
import com.RK8.V2.DTO.ReconciliationResult;
import com.RK8.V2.Parser.Gstr2BExcelParser;
import com.RK8.V2.Parser.PurchaseExcelParser;
import com.RK8.V2.Service.CAReportService;
import com.RK8.V2.Service.Purchase2BReconciliationService;
import com.RK8.V2.Service.ReconciliationReportService;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@RestController
@RequestMapping("/api/reconcile")
public class ReconciliationController {

    private final PurchaseExcelParser purchaseParser;
    private final Gstr2BExcelParser gstr2bParser;
    private final Purchase2BReconciliationService reconciliationService;
    private final ReconciliationReportService reportService;
    private final CAReportService re;

    public ReconciliationController(
            PurchaseExcelParser purchaseParser,
            Gstr2BExcelParser gstr2bParser,
            Purchase2BReconciliationService reconciliationService, ReconciliationReportService reportService, CAReportService re
    ) {
        this.purchaseParser = purchaseParser;
        this.gstr2bParser = gstr2bParser;
        this.reconciliationService = reconciliationService;
        this.reportService = reportService;
        this.re = re;
    }

    @PostMapping("/upload")
    public Map<String, Object> uploadFiles(
            @RequestParam("purchaseFile") MultipartFile purchaseFile,
            @RequestParam("gstr2bFile") MultipartFile gstr2bFile) {

        Map<String, Object> response = new HashMap<>();

        try {
            // Parse files
            List<PurchaseInvoiceDTO> purchases;
            List<Gstr2BDTO> gstr2bList;

            try (InputStream purchaseStream = purchaseFile.getInputStream();
                 InputStream gstr2bStream = gstr2bFile.getInputStream()) {

                purchases = purchaseParser.parse(purchaseStream);
                gstr2bList = gstr2bParser.parse(gstr2bStream);
            }

            response.put("purchaseCount", purchases.size());
            response.put("gstr2bCount", gstr2bList.size());

            // Reconcile
            List<ReconciliationResult> results = reconciliationService.reconcile(purchases, gstr2bList);

            // Calculate statistics
            long matched = results.stream().filter(r -> r.getStatus().startsWith("MATCHED")).count();
            long mismatch = results.stream().filter(r -> "MISMATCH".equals(r.getStatus())).count();
            long missingIn2B = results.stream().filter(r -> "MISSING_IN_2B".equals(r.getStatus())).count();
            long missingInPurchase = results.stream().filter(r -> "MISSING_IN_PURCHASE".equals(r.getStatus())).count();

            BigDecimal totalItcAtRisk = results.stream()
                    .map(ReconciliationResult::getItcAtRisk)
                    .reduce(BigDecimal.ZERO, BigDecimal::add);

            // Get ALL mismatches (non-matched items)
            List<Map<String, Object>> allMismatches = new ArrayList<>();
            List<Map<String, Object>> matchedItems = new ArrayList<>();
            List<Map<String, Object>> missingInPurchaseList = new ArrayList<>();
            List<Map<String, Object>> missingIn2BList = new ArrayList<>();

            for (ReconciliationResult r : results) {
                Map<String, Object> detail = new HashMap<>();
                detail.put("gstin", r.getSupplierGstin());
                detail.put("invoiceNo", r.getInvoiceNo());
                detail.put("status", r.getStatus());
                detail.put("purchaseTax", r.getPurchaseTax());
                detail.put("gstr2bTax", r.getGstr2bTax());
                detail.put("itcAtRisk", r.getItcAtRisk());
                detail.put("remarks", r.getRemarks());
                detail.put("invoiceMonth", r.getInvoiceMonth().toString());

                // Categorize
                if (r.getStatus().startsWith("MATCHED")) {
                    matchedItems.add(detail);
                } else if ("MISSING_IN_PURCHASE".equals(r.getStatus())) {
                    missingInPurchaseList.add(detail);
                    allMismatches.add(detail); // Add to mismatches
                } else if ("MISSING_IN_2B".equals(r.getStatus())) {
                    missingIn2BList.add(detail);
                    allMismatches.add(detail); // Add to mismatches
                } else if ("MISMATCH".equals(r.getStatus())) {
                    allMismatches.add(detail);
                }
            }

            // Sort mismatches by tax amount (descending)
            allMismatches.sort((a, b) -> {
                BigDecimal taxA = (BigDecimal) a.get("gstr2bTax");
                BigDecimal taxB = (BigDecimal) b.get("gstr2bTax");
                if (taxA == null && taxB == null) return 0;
                if (taxA == null) return 1;
                if (taxB == null) return -1;
                return taxB.compareTo(taxA); // Descending order
            });

            // Sort missing in purchase by tax amount (descending)
            missingInPurchaseList.sort((a, b) -> {
                BigDecimal taxA = (BigDecimal) a.get("gstr2bTax");
                BigDecimal taxB = (BigDecimal) b.get("gstr2bTax");
                if (taxA == null && taxB == null) return 0;
                if (taxA == null) return 1;
                if (taxB == null) return -1;
                return taxB.compareTo(taxA); // Descending order
            });

            response.put("totalResults", results.size());
            response.put("matched", matched);
            response.put("mismatch", mismatch);
            response.put("missingIn2B", missingIn2B);
            response.put("missingInPurchase", missingInPurchase);
            response.put("itcAtRisk", totalItcAtRisk);

            // Group by status for breakdown
            Map<String, Long> statusBreakdown = results.stream()
                    .collect(Collectors.groupingBy(ReconciliationResult::getStatus, Collectors.counting()));
            response.put("statusBreakdown", statusBreakdown);

            // Add all mismatches
            response.put("allMismatches", allMismatches);

            // Add categorized lists
            response.put("missingInPurchaseList", missingInPurchaseList);
            response.put("missingIn2BList", missingIn2BList);
            response.put("matchedItems", matchedItems);

            // Add summary for quick view
            Map<String, Object> summary = new HashMap<>();
            summary.put("totalInvoicesIn2B", gstr2bList.size());
            summary.put("totalInvoicesInPurchase", purchases.size());
            summary.put("matchedInvoices", matched);
            summary.put("unmatchedInvoices", allMismatches.size());
            summary.put("totalItcAvailableIn2B", calculateTotalTax(gstr2bList));
            summary.put("totalItcClaimedInPurchase", calculateTotalTaxFromPurchases(purchases));
            summary.put("itcAtRisk", totalItcAtRisk);
            summary.put("complianceRate", String.format("%.2f%%",
                    (matched * 100.0) / Math.max(gstr2bList.size(), purchases.size())));

            response.put("summary", summary);

        } catch (Exception e) {
            response.put("error", e.getMessage());
            e.printStackTrace();
        }

        return response;
    }

    @PostMapping("/detailed-report")
    public Map<String, Object> getDetailedReport(
            @RequestParam("purchaseFile") MultipartFile purchaseFile,
            @RequestParam("gstr2bFile") MultipartFile gstr2bFile) {

        Map<String, Object> response = new HashMap<>();

        try {
            // Parse files
            List<PurchaseInvoiceDTO> purchases;
            List<Gstr2BDTO> gstr2bList;

            try (InputStream purchaseStream = purchaseFile.getInputStream();
                 InputStream gstr2bStream = gstr2bFile.getInputStream()) {

                purchases = purchaseParser.parse(purchaseStream);
                gstr2bList = gstr2bParser.parse(gstr2bStream);
            }

            // Reconcile
            List<ReconciliationResult> results = reconciliationService.reconcile(purchases, gstr2bList);

            // Prepare detailed report
            List<Map<String, Object>> detailedMismatches = new ArrayList<>();

            for (ReconciliationResult r : results) {
                if (!r.getStatus().startsWith("MATCHED")) {
                    Map<String, Object> detail = new HashMap<>();
                    detail.put("supplierGstin", r.getSupplierGstin());
                    detail.put("invoiceNo", r.getInvoiceNo());
                    detail.put("invoiceMonth", r.getInvoiceMonth().toString());
                    detail.put("status", r.getStatus());
                    detail.put("purchaseTaxAmount", r.getPurchaseTax());
                    detail.put("gstr2bTaxAmount", r.getGstr2bTax());
                    detail.put("taxDifference", r.getGstr2bTax().subtract(r.getPurchaseTax()).abs());
                    detail.put("itcAtRiskAmount", r.getItcAtRisk());
                    detail.put("actionRequired", getActionRequired(r.getStatus()));
                    detail.put("priority", getPriority(r.getItcAtRisk()));
                    detail.put("remarks", r.getRemarks());

                    detailedMismatches.add(detail);
                }
            }

            // Sort by ITC at risk (descending)
            detailedMismatches.sort((a, b) -> {
                BigDecimal riskA = (BigDecimal) a.get("itcAtRiskAmount");
                BigDecimal riskB = (BigDecimal) b.get("itcAtRiskAmount");
                return riskB.compareTo(riskA);
            });

            // Calculate totals
            BigDecimal totalRisk = detailedMismatches.stream()
                    .map(d -> (BigDecimal) d.get("itcAtRiskAmount"))
                    .reduce(BigDecimal.ZERO, BigDecimal::add);

            // Group by action required
            Map<String, List<Map<String, Object>>> groupedByAction = detailedMismatches.stream()
                    .collect(Collectors.groupingBy(d -> (String) d.get("actionRequired")));

            response.put("totalMismatches", detailedMismatches.size());
            response.put("totalItcAtRisk", totalRisk);
            response.put("mismatchesByAction", groupedByAction);
            response.put("allMismatchesDetails", detailedMismatches);

            // Add month-wise summary
            Map<String, Map<String, Object>> monthSummary = new HashMap<>();
            for (Map<String, Object> mismatch : detailedMismatches) {
                String month = (String) mismatch.get("invoiceMonth");
                BigDecimal risk = (BigDecimal) mismatch.get("itcAtRiskAmount");

                monthSummary.computeIfAbsent(month, k -> {
                    Map<String, Object> monthData = new HashMap<>();
                    monthData.put("month", month);
                    monthData.put("count", 0);
                    monthData.put("totalRisk", BigDecimal.ZERO);
                    return monthData;
                });

                Map<String, Object> monthData = monthSummary.get(month);
                monthData.put("count", (Integer) monthData.get("count") + 1);
                monthData.put("totalRisk", ((BigDecimal) monthData.get("totalRisk")).add(risk));
            }

            response.put("monthWiseSummary", new ArrayList<>(monthSummary.values()));

        } catch (Exception e) {
            response.put("error", e.getMessage());
        }

        return response;
    }

    private String getActionRequired(String status) {
        switch (status) {
            case "MISSING_IN_PURCHASE":
                return "Add to Purchase Register";
            case "MISSING_IN_2B":
                return "Follow up with Supplier";
            case "MISMATCH":
                return "Verify Tax Amount";
            default:
                return "Review";
        }
    }

    private String getPriority(BigDecimal itcAtRisk) {
        if (itcAtRisk.compareTo(new BigDecimal("10000")) > 0) {
            return "HIGH";
        } else if (itcAtRisk.compareTo(new BigDecimal("1000")) > 0) {
            return "MEDIUM";
        } else {
            return "LOW";
        }
    }

    private BigDecimal calculateTotalTax(List<Gstr2BDTO> gstr2bList) {
        return gstr2bList.stream()
                .map(g -> g.getIgst().add(g.getCgst()).add(g.getSgst()))
                .reduce(BigDecimal.ZERO, BigDecimal::add);
    }

    private BigDecimal calculateTotalTaxFromPurchases(List<PurchaseInvoiceDTO> purchases) {
        return purchases.stream()
                .map(p -> p.getIgst().add(p.getCgst()).add(p.getSgst()))
                .reduce(BigDecimal.ZERO, BigDecimal::add);
    }

    @PostMapping("/generate-report")
    public Map<String, Object> generateReport(
            @RequestParam("purchaseFile") MultipartFile purchaseFile,
            @RequestParam("gstr2bFile") MultipartFile gstr2bFile) {

        Map<String, Object> response = new HashMap<>();

        try {
            // Parse files
            List<PurchaseInvoiceDTO> purchases;
            List<Gstr2BDTO> gstr2bList;

            try (InputStream purchaseStream = purchaseFile.getInputStream();
                 InputStream gstr2bStream = gstr2bFile.getInputStream()) {

                purchases = purchaseParser.parse(purchaseStream);
                gstr2bList = gstr2bParser.parse(gstr2bStream);
            }

            // Reconcile
            List<ReconciliationResult> results = reconciliationService.reconcile(purchases, gstr2bList);

            // Generate detailed report
            Map<String, Object> report = reportService.generateActionReport(results);

            // Add basic stats
            report.put("purchaseInvoiceCount", purchases.size());
            report.put("gstr2bInvoiceCount", gstr2bList.size());

            return report;

        } catch (Exception e) {
            response.put("error", e.getMessage());
            e.printStackTrace();
            return response;
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
                gstr2bList = gstr2bParser.parse(gstr2bStream);
            }

            // Reconcile
            List<ReconciliationResult> results = reconciliationService.reconcile(purchases, gstr2bList);

            // Generate CA report
            byte[] reportBytes = re.generateCAReport(purchases, gstr2bList, results);

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
                gstr2bList = gstr2bParser.parse(gstr2bStream);
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


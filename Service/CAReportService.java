package com.RK8.V2.Service;

import com.RK8.V2.DTO.Gstr2BDTO;
import com.RK8.V2.DTO.PurchaseInvoiceDTO;
import com.RK8.V2.DTO.ReconciliationResult;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.math.BigDecimal;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

@Service
public class CAReportService {

    private static final DateTimeFormatter DATE_FORMATTER = DateTimeFormatter.ofPattern("dd-MMM-yyyy");

    public byte[] generateCAReport(List<PurchaseInvoiceDTO> purchases,
                                   List<Gstr2BDTO> gstr2bList,
                                   List<ReconciliationResult> results) {

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            // Create professional styles
            Map<String, CellStyle> styles = createStyles(workbook);

            // Sheet 1: Executive Summary
            createExecutiveSummarySheet(workbook, styles, purchases, gstr2bList, results);

            // Sheet 2: Reconciliation Details
            createReconciliationDetailsSheet(workbook, styles, results);

            // Sheet 3: Missing Invoices (Action Required)
            createMissingInvoicesSheet(workbook, styles, results);

            // Sheet 4: Matched Invoices
            createMatchedInvoicesSheet(workbook, styles, results);

            // Sheet 5: ITC Summary by Month
            createMonthlySummarySheet(workbook, styles, results);

            // Sheet 6: Supplier-wise Summary
            createSupplierSummarySheet(workbook, styles, results);

            // Write to byte array
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();

        } catch (Exception e) {
            throw new RuntimeException("Failed to generate CA report", e);
        }
    }

    private Map<String, CellStyle> createStyles(XSSFWorkbook workbook) {
        Map<String, CellStyle> styles = new HashMap<>();

        // Title style
        CellStyle titleStyle = workbook.createCellStyle();
        Font titleFont = workbook.createFont();
        titleFont.setBold(true);
        titleFont.setFontHeightInPoints((short) 16);
        titleStyle.setFont(titleFont);
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        styles.put("title", titleStyle);

        // Header style
        CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        headerStyle.setFont(headerFont);
        headerStyle.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        styles.put("header", headerStyle);

        // Data style
        CellStyle dataStyle = workbook.createCellStyle();
        dataStyle.setBorderBottom(BorderStyle.THIN);
        dataStyle.setBorderTop(BorderStyle.THIN);
        dataStyle.setBorderLeft(BorderStyle.THIN);
        dataStyle.setBorderRight(BorderStyle.THIN);
        styles.put("data", dataStyle);

        // Currency style
        CellStyle currencyStyle = workbook.createCellStyle();
        currencyStyle.cloneStyleFrom(dataStyle);
        DataFormat df = workbook.createDataFormat();
        currencyStyle.setDataFormat(df.getFormat("₹#,##0.00"));
        styles.put("currency", currencyStyle);

        // Highlight style (for issues)
        CellStyle highlightStyle = workbook.createCellStyle();
        highlightStyle.cloneStyleFrom(dataStyle);
        highlightStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        highlightStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.put("highlight", highlightStyle);

        // Red highlight for high risk
        CellStyle redHighlight = workbook.createCellStyle();
        redHighlight.cloneStyleFrom(dataStyle);
        redHighlight.setFillForegroundColor(IndexedColors.CORAL.getIndex());
        redHighlight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.put("red", redHighlight);

        return styles;
    }

    private void createExecutiveSummarySheet(XSSFWorkbook workbook,
                                             Map<String, CellStyle> styles,
                                             List<PurchaseInvoiceDTO> purchases,
                                             List<Gstr2BDTO> gstr2bList,
                                             List<ReconciliationResult> results) {

        Sheet sheet = workbook.createSheet("Executive Summary");

        int rowNum = 0;

        // Title
        Row titleRow = sheet.createRow(rowNum++);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("GSTR-2B vs Purchase Register Reconciliation Report");
        titleCell.setCellStyle(styles.get("title"));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 7));

        // Report Date
        Row dateRow = sheet.createRow(rowNum++);
        dateRow.createCell(0).setCellValue("Report Date: " +
                java.time.LocalDate.now().format(DateTimeFormatter.ofPattern("dd-MMM-yyyy")));

        // Period
        rowNum++;
        Row periodRow = sheet.createRow(rowNum++);
        periodRow.createCell(0).setCellValue("Period: " + getPeriodCovered(results));

        // Summary Table
        rowNum += 2;
        Row summaryHeader = sheet.createRow(rowNum++);
        String[] summaryHeaders = {"Parameter", "Count", "Amount (₹)", "Remarks"};
        for (int i = 0; i < summaryHeaders.length; i++) {
            Cell cell = summaryHeader.createCell(i);
            cell.setCellValue(summaryHeaders[i]);
            cell.setCellStyle(styles.get("header"));
        }

        // Calculate statistics
        Map<String, Long> statusCounts = results.stream()
                .collect(Collectors.groupingBy(ReconciliationResult::getStatus, Collectors.counting()));

        long matched = statusCounts.getOrDefault("MATCHED", 0L) +
                statusCounts.getOrDefault("MATCHED_WITH_TOLERANCE", 0L);
        long missingInPurchase = statusCounts.getOrDefault("MISSING_IN_PURCHASE", 0L);
        long missingIn2B = statusCounts.getOrDefault("MISSING_IN_2B", 0L);

        BigDecimal total2BTax = gstr2bList.stream()
                .map(g -> g.getIgst().add(g.getCgst()).add(g.getSgst()))
                .reduce(BigDecimal.ZERO, BigDecimal::add);

        BigDecimal totalPurchaseTax = purchases.stream()
                .map(p -> p.getIgst().add(p.getCgst()).add(p.getSgst()))
                .reduce(BigDecimal.ZERO, BigDecimal::add);

        BigDecimal itcAtRisk = results.stream()
                .map(ReconciliationResult::getItcAtRisk)
                .reduce(BigDecimal.ZERO, BigDecimal::add);

        BigDecimal itcAvailable = total2BTax;
        BigDecimal itcClaimed = totalPurchaseTax;
        BigDecimal itcUnclaimed = itcAvailable.subtract(itcClaimed).max(BigDecimal.ZERO);

        // Summary rows
        String[][] summaryData = {
                {"Total Invoices in GSTR-2B", String.valueOf(gstr2bList.size()), formatCurrency(total2BTax), "Total ITC available as per GSTR-2B"},
                {"Total Invoices in Purchase Register", String.valueOf(purchases.size()), formatCurrency(totalPurchaseTax), "Total ITC claimed in books"},
                {"Successfully Matched Invoices", String.valueOf(matched), formatCurrency(itcClaimed), "Verified ITC"},
                {"Invoices Missing in Purchase Register", String.valueOf(missingInPurchase), formatCurrency(itcAtRisk), "ITC at risk - Need to add to books"},
                {"Invoices Missing in GSTR-2B", String.valueOf(missingIn2B), "0.00", "Need to follow up with suppliers"},
                {"ITC Unclaimed", "-", formatCurrency(itcUnclaimed), "Potential additional ITC available"},
                {"Compliance Rate", String.format("%.1f%%", (matched * 100.0) / purchases.size()), "-", "Percentage of purchase invoices verified"}
        };

        for (String[] rowData : summaryData) {
            Row row = sheet.createRow(rowNum++);
            for (int i = 0; i < rowData.length; i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(rowData[i]);
                cell.setCellStyle(styles.get("data"));

                if (i == 2 && !rowData[i].equals("-")) { // Currency column
                    cell.setCellStyle(styles.get("currency"));
                }

                if (rowData[0].contains("Missing") || rowData[0].contains("Unclaimed")) {
                    cell.setCellStyle(styles.get("highlight"));
                }
            }
        }

        // Action Required Section
        rowNum += 2;
        Row actionHeader = sheet.createRow(rowNum++);
        actionHeader.createCell(0).setCellValue("Action Required:");
        actionHeader.getCell(0).setCellStyle(styles.get("header"));
        sheet.addMergedRegion(new CellRangeAddress(rowNum-1, rowNum-1, 0, 3));

        String[] actions = {
                "1. Add " + missingInPurchase + " missing invoices to purchase register to claim ₹" + formatCurrency(itcAtRisk) + " ITC",
                "2. Follow up with suppliers for " + missingIn2B + " invoices not reflecting in GSTR-2B",
                "3. Ensure all future purchases are recorded in purchase register promptly",
                "4. Monthly reconciliation recommended to avoid ITC loss"
        };

        for (String action : actions) {
            Row actionRow = sheet.createRow(rowNum++);
            actionRow.createCell(0).setCellValue(action);
        }

        // Auto-size columns
        for (int i = 0; i < 4; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private void createReconciliationDetailsSheet(XSSFWorkbook workbook,
                                                  Map<String, CellStyle> styles,
                                                  List<ReconciliationResult> results) {

        Sheet sheet = workbook.createSheet("Reconciliation Details");
        int rowNum = 0;

        // Header
        Row header = sheet.createRow(rowNum++);
        String[] headers = {"Sr No", "Supplier GSTIN", "Invoice No", "Invoice Month",
                "Purchase Tax (₹)", "GSTR-2B Tax (₹)", "Difference (₹)",
                "ITC at Risk (₹)", "Status", "Remarks/Action"};

        for (int i = 0; i < headers.length; i++) {
            Cell cell = header.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(styles.get("header"));
        }

        // Filter and sort mismatches
        List<ReconciliationResult> mismatches = results.stream()
                .filter(r -> !r.getStatus().startsWith("MATCHED"))
                .sorted((a, b) -> b.getItcAtRisk().compareTo(a.getItcAtRisk()))
                .collect(Collectors.toList());

        // Data rows
        int srNo = 1;
        for (ReconciliationResult r : mismatches) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(srNo++);
            row.createCell(1).setCellValue(r.getSupplierGstin());
            row.createCell(2).setCellValue(r.getInvoiceNo());
            row.createCell(3).setCellValue(r.getInvoiceMonth().toString());

            Cell purchaseTaxCell = row.createCell(4);
            purchaseTaxCell.setCellValue(r.getPurchaseTax().doubleValue());
            purchaseTaxCell.setCellStyle(styles.get("currency"));

            Cell gstr2bTaxCell = row.createCell(5);
            gstr2bTaxCell.setCellValue(r.getGstr2bTax().doubleValue());
            gstr2bTaxCell.setCellStyle(styles.get("currency"));

            Cell diffCell = row.createCell(6);
            BigDecimal diff = r.getGstr2bTax().subtract(r.getPurchaseTax()).abs();
            diffCell.setCellValue(diff.doubleValue());
            diffCell.setCellStyle(styles.get("currency"));

            Cell riskCell = row.createCell(7);
            riskCell.setCellValue(r.getItcAtRisk().doubleValue());
            riskCell.setCellStyle(styles.get("currency"));
            if (r.getItcAtRisk().compareTo(new BigDecimal("10000")) > 0) {
                riskCell.setCellStyle(styles.get("red"));
            }

            row.createCell(8).setCellValue(r.getStatus());
            row.createCell(9).setCellValue(getActionForStatus(r.getStatus()));

            // Apply highlight to entire row for missing in purchase
            if ("MISSING_IN_PURCHASE".equals(r.getStatus())) {
                for (int i = 0; i < 10; i++) {
                    row.getCell(i).setCellStyle(styles.get("highlight"));
                }
            }
        }

        // Add totals row
        Row totalRow = sheet.createRow(rowNum++);
        totalRow.createCell(3).setCellValue("TOTAL:");

        BigDecimal totalRisk = mismatches.stream()
                .map(ReconciliationResult::getItcAtRisk)
                .reduce(BigDecimal.ZERO, BigDecimal::add);

        Cell totalRiskCell = totalRow.createCell(7);
        totalRiskCell.setCellValue(totalRisk.doubleValue());
        totalRiskCell.setCellStyle(styles.get("currency"));

        // Auto-size columns
        for (int i = 0; i < 10; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private void createMissingInvoicesSheet(XSSFWorkbook workbook,
                                            Map<String, CellStyle> styles,
                                            List<ReconciliationResult> results) {

        Sheet sheet = workbook.createSheet("Missing In Purchase Register");
        int rowNum = 0;

        // Title
        Row titleRow = sheet.createRow(rowNum++);
        titleRow.createCell(0).setCellValue("Invoices in GSTR-2B but Missing in Purchase Register");
        titleRow.getCell(0).setCellStyle(styles.get("title"));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 7));

        rowNum++;
        Row subTitle = sheet.createRow(rowNum++);
        subTitle.createCell(0).setCellValue("Action Required: Add these invoices to purchase register to claim ITC");

        // Header
        Row header = sheet.createRow(rowNum++);
        String[] headers = {"Sr No", "Supplier GSTIN", "Invoice No", "Invoice Date",
                "Taxable Value (₹)", "IGST (₹)", "CGST (₹)", "SGST (₹)",
                "Total Tax (₹)", "Priority"};

        for (int i = 0; i < headers.length; i++) {
            Cell cell = header.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(styles.get("header"));
        }

        // Get missing in purchase invoices
        List<ReconciliationResult> missingInPurchase = results.stream()
                .filter(r -> "MISSING_IN_PURCHASE".equals(r.getStatus()))
                .sorted((a, b) -> b.getGstr2bTax().compareTo(a.getGstr2bTax()))
                .collect(Collectors.toList());

        // Data rows
        int srNo = 1;
        for (ReconciliationResult r : missingInPurchase) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(srNo++);
            row.createCell(1).setCellValue(r.getSupplierGstin());
            row.createCell(2).setCellValue(r.getInvoiceNo());
            row.createCell(3).setCellValue(r.getInvoiceMonth().toString() + "-01");

            // Since we don't have taxable value in ReconciliationResult,
            // we'll show placeholder or calculate from tax
            Cell taxableCell = row.createCell(4);
            taxableCell.setCellValue(0.0); // Placeholder
            taxableCell.setCellStyle(styles.get("currency"));

            // Tax breakdown (simplified - assuming IGST only for mismatch)
            Cell igstCell = row.createCell(5);
            igstCell.setCellValue(r.getGstr2bTax().doubleValue());
            igstCell.setCellStyle(styles.get("currency"));

            row.createCell(6).setCellValue(0.0);
            row.createCell(7).setCellValue(0.0);

            Cell totalTaxCell = row.createCell(8);
            totalTaxCell.setCellValue(r.getGstr2bTax().doubleValue());
            totalTaxCell.setCellStyle(styles.get("currency"));

            String priority = getPriority(r.getGstr2bTax());
            row.createCell(9).setCellValue(priority);

            // Color code based on priority
            if ("HIGH".equals(priority)) {
                for (int i = 0; i < 10; i++) {
                    row.getCell(i).setCellStyle(styles.get("red"));
                }
            }
        }

        // Auto-size columns
        for (int i = 0; i < 10; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private void createMatchedInvoicesSheet(XSSFWorkbook workbook,
                                            Map<String, CellStyle> styles,
                                            List<ReconciliationResult> results) {

        Sheet sheet = workbook.createSheet("Matched Invoices");
        int rowNum = 0;

        // Title
        Row titleRow = sheet.createRow(rowNum++);
        titleRow.createCell(0).setCellValue("Successfully Matched Invoices");
        titleRow.getCell(0).setCellStyle(styles.get("title"));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 6));

        // Header
        Row header = sheet.createRow(rowNum++);
        String[] headers = {"Supplier GSTIN", "Invoice No", "Invoice Month",
                "Purchase Tax (₹)", "GSTR-2B Tax (₹)", "Status", "Verification"};

        for (int i = 0; i < headers.length; i++) {
            Cell cell = header.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(styles.get("header"));
        }

        // Get matched invoices
        List<ReconciliationResult> matched = results.stream()
                .filter(r -> r.getStatus().startsWith("MATCHED"))
                .collect(Collectors.toList());

        // Data rows
        for (ReconciliationResult r : matched) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(r.getSupplierGstin());
            row.createCell(1).setCellValue(r.getInvoiceNo());
            row.createCell(2).setCellValue(r.getInvoiceMonth().toString());

            Cell purchaseTaxCell = row.createCell(3);
            purchaseTaxCell.setCellValue(r.getPurchaseTax().doubleValue());
            purchaseTaxCell.setCellStyle(styles.get("currency"));

            Cell gstr2bTaxCell = row.createCell(4);
            gstr2bTaxCell.setCellValue(r.getGstr2bTax().doubleValue());
            gstr2bTaxCell.setCellStyle(styles.get("currency"));

            row.createCell(5).setCellValue(r.getStatus());
            row.createCell(6).setCellValue("✓ Verified");
        }

        // Auto-size columns
        for (int i = 0; i < 7; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private void createMonthlySummarySheet(XSSFWorkbook workbook,
                                           Map<String, CellStyle> styles,
                                           List<ReconciliationResult> results) {

        Sheet sheet = workbook.createSheet("Monthly ITC Summary");
        int rowNum = 0;

        // Title
        Row titleRow = sheet.createRow(rowNum++);
        titleRow.createCell(0).setCellValue("Month-wise ITC Analysis");
        titleRow.getCell(0).setCellStyle(styles.get("title"));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));

        // Header
        Row header = sheet.createRow(rowNum++);
        String[] headers = {"Month", "Total Invoices", "ITC Available (₹)",
                "ITC Claimed (₹)", "ITC at Risk (₹)", "Compliance %"};

        for (int i = 0; i < headers.length; i++) {
            Cell cell = header.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(styles.get("header"));
        }

        // Group by month
        Map<String, List<ReconciliationResult>> byMonth = results.stream()
                .collect(Collectors.groupingBy(r -> r.getInvoiceMonth().toString()));

        // Data rows
        for (Map.Entry<String, List<ReconciliationResult>> entry : byMonth.entrySet()) {
            String month = entry.getKey();
            List<ReconciliationResult> monthResults = entry.getValue();

            long totalInvoices = monthResults.size();
            BigDecimal itcAvailable = monthResults.stream()
                    .map(ReconciliationResult::getGstr2bTax)
                    .reduce(BigDecimal.ZERO, BigDecimal::add);

            BigDecimal itcClaimed = monthResults.stream()
                    .map(ReconciliationResult::getPurchaseTax)
                    .reduce(BigDecimal.ZERO, BigDecimal::add);

            BigDecimal itcAtRisk = monthResults.stream()
                    .map(ReconciliationResult::getItcAtRisk)
                    .reduce(BigDecimal.ZERO, BigDecimal::add);

            double complianceRate = itcAvailable.compareTo(BigDecimal.ZERO) == 0 ? 100.0 :
                    (itcClaimed.doubleValue() / itcAvailable.doubleValue()) * 100;

            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(month);
            row.createCell(1).setCellValue(totalInvoices);

            Cell availableCell = row.createCell(2);
            availableCell.setCellValue(itcAvailable.doubleValue());
            availableCell.setCellStyle(styles.get("currency"));

            Cell claimedCell = row.createCell(3);
            claimedCell.setCellValue(itcClaimed.doubleValue());
            claimedCell.setCellStyle(styles.get("currency"));

            Cell riskCell = row.createCell(4);
            riskCell.setCellValue(itcAtRisk.doubleValue());
            riskCell.setCellStyle(styles.get("currency"));

            Cell complianceCell = row.createCell(5);
            complianceCell.setCellValue(String.format("%.1f%%", complianceRate));

            if (complianceRate < 80) {
                complianceCell.setCellStyle(styles.get("red"));
            }
        }

        // Auto-size columns
        for (int i = 0; i < 6; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private void createSupplierSummarySheet(XSSFWorkbook workbook,
                                            Map<String, CellStyle> styles,
                                            List<ReconciliationResult> results) {

        Sheet sheet = workbook.createSheet("Supplier-wise Summary");
        int rowNum = 0;

        // Title
        Row titleRow = sheet.createRow(rowNum++);
        titleRow.createCell(0).setCellValue("Supplier-wise ITC Analysis");
        titleRow.getCell(0).setCellStyle(styles.get("title"));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 5));

        // Header
        Row header = sheet.createRow(rowNum++);
        String[] headers = {"Supplier GSTIN", "Total Invoices", "Matched",
                "Missing in Purchase", "ITC Amount (₹)", "ITC at Risk (₹)"};

        for (int i = 0; i < headers.length; i++) {
            Cell cell = header.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(styles.get("header"));
        }

        // Group by supplier
        Map<String, List<ReconciliationResult>> bySupplier = results.stream()
                .collect(Collectors.groupingBy(ReconciliationResult::getSupplierGstin));

        // Data rows
        for (Map.Entry<String, List<ReconciliationResult>> entry : bySupplier.entrySet()) {
            String gstin = entry.getKey();
            List<ReconciliationResult> supplierResults = entry.getValue();

            long totalInvoices = supplierResults.size();
            long matched = supplierResults.stream()
                    .filter(r -> r.getStatus().startsWith("MATCHED"))
                    .count();

            long missingInPurchase = supplierResults.stream()
                    .filter(r -> "MISSING_IN_PURCHASE".equals(r.getStatus()))
                    .count();

            BigDecimal totalTax = supplierResults.stream()
                    .map(ReconciliationResult::getGstr2bTax)
                    .reduce(BigDecimal.ZERO, BigDecimal::add);

            BigDecimal risk = supplierResults.stream()
                    .map(ReconciliationResult::getItcAtRisk)
                    .reduce(BigDecimal.ZERO, BigDecimal::add);

            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(gstin);
            row.createCell(1).setCellValue(totalInvoices);
            row.createCell(2).setCellValue(matched);
            row.createCell(3).setCellValue(missingInPurchase);

            Cell taxCell = row.createCell(4);
            taxCell.setCellValue(totalTax.doubleValue());
            taxCell.setCellStyle(styles.get("currency"));

            Cell riskCell = row.createCell(5);
            riskCell.setCellValue(risk.doubleValue());
            riskCell.setCellStyle(styles.get("currency"));

            if (risk.compareTo(BigDecimal.ZERO) > 0) {
                row.getCell(3).setCellStyle(styles.get("highlight"));
                riskCell.setCellStyle(styles.get("red"));
            }
        }

        // Auto-size columns
        for (int i = 0; i < 6; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    // Helper methods
    private String formatCurrency(BigDecimal amount) {
        return String.format("₹%,.2f", amount);
    }

    private String getPeriodCovered(List<ReconciliationResult> results) {
        if (results.isEmpty()) return "N/A";

        String minMonth = results.stream()
                .map(r -> r.getInvoiceMonth().toString())
                .min(String::compareTo)
                .orElse("N/A");

        String maxMonth = results.stream()
                .map(r -> r.getInvoiceMonth().toString())
                .max(String::compareTo)
                .orElse("N/A");

        return minMonth.equals(maxMonth) ? minMonth : minMonth + " to " + maxMonth;
    }

    private String getActionForStatus(String status) {
        switch (status) {
            case "MISSING_IN_PURCHASE":
                return "ADD TO PURCHASE REGISTER";
            case "MISSING_IN_2B":
                return "FOLLOW UP WITH SUPPLIER";
            case "MISMATCH":
                return "VERIFY TAX AMOUNT";
            default:
                return "REVIEW";
        }
    }

    private String getPriority(BigDecimal amount) {
        if (amount.compareTo(new BigDecimal("10000")) > 0) {
            return "HIGH";
        } else if (amount.compareTo(new BigDecimal("1000")) > 0) {
            return "MEDIUM";
        } else {
            return "LOW";
        }
    }
}

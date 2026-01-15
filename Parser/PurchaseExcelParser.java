package com.RK8.V2.Parser;

import com.RK8.V2.DTO.PurchaseInvoiceDTO;
import org.apache.poi.ss.usermodel.*;
import java.io.InputStream;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import org.springframework.stereotype.Component;
import java.time.ZoneId;
import java.util.*;

@Component
public class PurchaseExcelParser {
    private static final String INVOICE_NO = "INVOICE_NO";
    private static final String SUPPLIER_GSTIN = "SUPPLIER_GSTIN";
    private static final String INVOICE_DATE = "INVOICE_DATE";
    private static final String IGST = "IGST";
    private static final String CGST = "CGST";
    private static final String SGST = "SGST";
    private static final String PARTICULARS = "PARTICULARS";
    private static final String GROSS_TOTAL = "GROSS_TOTAL";



    public List<PurchaseInvoiceDTO> parse(InputStream is) throws Exception {
        try (Workbook workbook = WorkbookFactory.create(is)) {
            Sheet sheet = workbook.getSheetAt(0);
            int headerRowIndex = findHeaderRow(sheet);
            Row headerRow = sheet.getRow(headerRowIndex);
            Map<String, Integer> colIndex = buildColumnIndexMap(headerRow);
            validateColumns(colIndex);

            List<PurchaseInvoiceDTO> list = new ArrayList<>();
            for (int i = headerRowIndex + 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null || isEmptyRow(row)) continue;

                // Check for total row
                Cell firstCell = row.getCell(0);
                if (firstCell != null && "Grand Total".equalsIgnoreCase(getStringValue(firstCell).trim())) {
                    continue;
                }

                String rawInvoice = getStringValue(row.getCell(colIndex.get(INVOICE_NO)));
                if (rawInvoice == null || rawInvoice.trim().isEmpty() ||
                        rawInvoice.equalsIgnoreCase("Grand Total")) {
                    continue;
                }

                // IMPORTANT: Don't skip zero tax rows
                BigDecimal igst = parseBigDecimal(row.getCell(colIndex.get(IGST)));
                BigDecimal cgst = parseBigDecimal(row.getCell(colIndex.get(CGST)));
                BigDecimal sgst = parseBigDecimal(row.getCell(colIndex.get(SGST)));

                PurchaseInvoiceDTO dto = new PurchaseInvoiceDTO();
                dto.setInvoiceNo(normalizeInvoice(rawInvoice));
                dto.setSupplierGstin(normalizeGstin(
                        getStringValue(row.getCell(colIndex.get(SUPPLIER_GSTIN)))
                ));

                LocalDate invoiceDate = parseDate(row.getCell(colIndex.get(INVOICE_DATE)));
                if (invoiceDate == null) {
                    System.err.println("Skipping row due to invalid date: " + rawInvoice);
                    continue;
                }
                dto.setInvoiceDate(invoiceDate);

                dto.setIgst(igst);
                dto.setCgst(cgst);
                dto.setSgst(sgst);
                dto.setParticulars(getStringValue(row.getCell(colIndex.get(PARTICULARS))));

                Integer grossTotalIdx = colIndex.get(GROSS_TOTAL);
                if (grossTotalIdx != null) {
                    dto.setGrossTotal(parseBigDecimal(row.getCell(grossTotalIdx)));
                }

                list.add(dto);
            }

            System.out.println("Purchase Parser: Loaded " + list.size() + " invoices");
            return list;
        }
    }

    private boolean isEmptyRow(Row row) {
        if (row == null) return true;
        for (Cell cell : row) {
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                String val = getStringValue(cell).trim();
                if (!val.isEmpty()) {
                    return false;
                }
            }
        }
        return true;
    }

    private int findHeaderRow(Sheet sheet) {
        for (int i = 0; i <= 20; i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            int hits = 0;
            for (Cell cell : row) {
                String v = getStringValue(cell).toUpperCase();
                if (v.contains("INVOICE") && v.contains("NO")) hits++;
                if (v.contains("SUPPLIER") && v.contains("GST")) hits++;
                if (v.contains("INVOICE") && v.contains("DATE")) hits++;
            }
            if (hits >= 2) {
                System.out.println("Found Purchase header at row: " + i);
                return i;
            }
        }
        throw new RuntimeException("Purchase header row not found");
    }

    private Map<String, Integer> buildColumnIndexMap(Row headerRow) {
        Map<String, Integer> map = new HashMap<>();
        for (Cell cell : headerRow) {
            String h = getStringValue(cell).toUpperCase().trim();

            if (h.contains("INVOICE") && h.contains("NO")) {
                map.put(INVOICE_NO, cell.getColumnIndex());
                System.out.println("Purchase - Found INVOICE_NO at column: " + cell.getColumnIndex());
            }
            if (h.contains("SUPPLIER") && h.contains("GST")) {
                map.put(SUPPLIER_GSTIN, cell.getColumnIndex());
                System.out.println("Purchase - Found SUPPLIER_GSTIN at column: " + cell.getColumnIndex());
            }
            if (h.contains("INVOICE") && h.contains("DATE")) {
                map.put(INVOICE_DATE, cell.getColumnIndex());
                System.out.println("Purchase - Found INVOICE_DATE at column: " + cell.getColumnIndex());
            }
            if (h.contains("IGST")) {
                map.put(IGST, cell.getColumnIndex());
                System.out.println("Purchase - Found IGST at column: " + cell.getColumnIndex());
            }
            if (h.contains("CGST")) {
                map.put(CGST, cell.getColumnIndex());
                System.out.println("Purchase - Found CGST at column: " + cell.getColumnIndex());
            }
            if (h.contains("SGST")) {
                map.put(SGST, cell.getColumnIndex());
                System.out.println("Purchase - Found SGST at column: " + cell.getColumnIndex());
            }
            if (h.contains("PARTICULAR") || h.contains("NAME")) {
                map.put(PARTICULARS, cell.getColumnIndex());
                System.out.println("Purchase - Found PARTICULARS at column: " + cell.getColumnIndex());
            }
            if (h.contains("GROSS") || h.contains("TOTAL")) {
                map.put(GROSS_TOTAL, cell.getColumnIndex());
                System.out.println("Purchase - Found GROSS_TOTAL at column: " + cell.getColumnIndex());
            }
        }
        return map;
    }

    private void validateColumns(Map<String, Integer> colIndex) {
        String[] required = {INVOICE_NO, SUPPLIER_GSTIN, INVOICE_DATE, IGST, CGST, SGST};
        for (String col : required) {
            if (!colIndex.containsKey(col)) {
                throw new RuntimeException("Missing required column in purchase file: " + col);
            }
        }
    }

    private String getStringValue(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                return BigDecimal.valueOf(cell.getNumericCellValue())
                        .stripTrailingZeros().toPlainString();
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    try {
                        return String.valueOf(cell.getNumericCellValue());
                    } catch (Exception ex) {
                        return "";
                    }
                }
            default:
                return "";
        }
    }

    private BigDecimal parseBigDecimal(Cell cell) {
        try {
            if (cell == null) return BigDecimal.ZERO;

            switch (cell.getCellType()) {
                case NUMERIC:
                    return BigDecimal.valueOf(cell.getNumericCellValue());
                case STRING:
                    String val = cell.getStringCellValue().trim();
                    if (val.isEmpty()) return BigDecimal.ZERO;
                    return new BigDecimal(val);
                case FORMULA:
                    try {
                        return BigDecimal.valueOf(cell.getNumericCellValue());
                    } catch (Exception e) {
                        String val2 = cell.getStringCellValue().trim();
                        if (val2.isEmpty()) return BigDecimal.ZERO;
                        return new BigDecimal(val2);
                    }
                default:
                    return BigDecimal.ZERO;
            }
        } catch (Exception e) {
            return BigDecimal.ZERO;
        }
    }

    private LocalDate parseDate(Cell cell) {
        if (cell == null) return null;

        try {
            // Try Excel date format first
            if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue().toInstant()
                        .atZone(ZoneId.systemDefault())
                        .toLocalDate();
            }

            // Try string parsing
            String dateStr = getStringValue(cell).trim();
            if (dateStr.isEmpty()) return null;

            // Remove time part if present (for "2025-09-01 00:00:00")
            if (dateStr.contains(" ")) {
                dateStr = dateStr.split(" ")[0];
            }

            // Try different date formats
            DateTimeFormatter[] formatters = {
                    DateTimeFormatter.ofPattern("yyyy-MM-dd"),
                    DateTimeFormatter.ofPattern("dd/MM/yyyy"),
                    DateTimeFormatter.ofPattern("dd-MM-yyyy"),
                    DateTimeFormatter.ofPattern("dd-MMM-yyyy"),
                    DateTimeFormatter.ofPattern("yyyy/MM/dd")
            };

            for (DateTimeFormatter fmt : formatters) {
                try {
                    return LocalDate.parse(dateStr, fmt);
                } catch (Exception ignored) {}
            }

            System.err.println("Could not parse purchase date: " + dateStr);
            return null;

        } catch (Exception e) {
            System.err.println("Error parsing purchase date: " + e.getMessage());
            return null;
        }
    }

    private String normalizeGstin(String g) {
        if (g == null || g.trim().isEmpty()) return "";
        return g.trim().toUpperCase().replaceAll("[^A-Z0-9]", "");
    }

    // Add this method to both parsers
    private String extractNumericInvoice(String invoice) {
        if (invoice == null) return "";

        // Extract the last sequence of digits
        String normalized = invoice.replaceAll("[^0-9]", " ");
        String[] parts = normalized.trim().split("\\s+");

        if (parts.length > 0) {
            // Return the last numeric part (usually the invoice number)
            return parts[parts.length - 1];
        }

        return "";
    }

    // Update the normalizeInvoice method
    private String normalizeInvoice(String i) {
        if (i == null) return "";

        String base = i.trim().toUpperCase();

        // Common pattern fixes
        base = base.replace("FY25-26/", "")
                .replace("GST-25-26/", "")
                .replace("EP/2025-26/", "")
                .replace("JE/2025-26/", "")
                .replace("TIA/T/", "")
                .replace("/24-25", "")
                .replace("/25-26", "")
                .replaceAll("\\s+", "");

        // Character normalization (O/0, I/1, etc.)
        base = base.replace('O', '0')
                .replace('I', '1')
                .replace('L', '1');

        return base;
    }
}





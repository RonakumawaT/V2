package com.RK8.V2.Parser;

import com.RK8.V2.DTO.Gstr2BDTO;
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
public class Gstr2BExcelParser {
    private static final String INVOICE_NO = "INVOICE_NO";
    private static final String SUPPLIER_GSTIN = "SUPPLIER_GSTIN";
    private static final String INVOICE_DATE = "INVOICE_DATE";
    private static final String IGST = "IGST";
    private static final String CGST = "CGST";
    private static final String SGST = "SGST";
    private static final String TAXABLE_VALUE = "TAXABLE_VALUE";
    private static final String INVOICE_VALUE = "INVOICE_VALUE";
    private static final String PARTICULARS = "PARTICULARS";

    public List<Gstr2BDTO> parse(InputStream is) throws Exception {
        try (Workbook wb = WorkbookFactory.create(is)) {
            Sheet sheet = wb.getSheetAt(0);
            int headerRow = findHeaderRow(sheet);
            Row header = sheet.getRow(headerRow);
            Map<String, Integer> col = buildColumnMap(header);
            validate(col);

            List<Gstr2BDTO> out = new ArrayList<>();
            for (int i = headerRow + 1; i <= sheet.getLastRowNum(); i++) {
                Row r = sheet.getRow(i);
                if (r == null || isEmptyRow(r)) continue;

                // Check for total row
                Cell firstCell = r.getCell(0);
                if (firstCell != null && "Total".equalsIgnoreCase(getStringValue(firstCell).trim())) {
                    continue;
                }

                String invoice = getStringValue(r.getCell(col.get(INVOICE_NO)));
                if (invoice == null || invoice.trim().isEmpty() || invoice.equalsIgnoreCase("Total")) {
                    continue;
                }

                // Parse all values
                String supplierGstin = getStringValue(r.getCell(col.get(SUPPLIER_GSTIN)));
                LocalDate invoiceDate = parseDate(r.getCell(col.get(INVOICE_DATE)));

                // IMPORTANT: Don't skip zero tax rows! Include ALL invoices
                BigDecimal igst = parseBigDecimal(r.getCell(col.get(IGST)));
                BigDecimal cgst = parseBigDecimal(r.getCell(col.get(CGST)));
                BigDecimal sgst = parseBigDecimal(r.getCell(col.get(SGST)));

                Gstr2BDTO d = new Gstr2BDTO();
                d.setInvoiceNo(normalizeInvoice(invoice));
                d.setSupplierGstin(normalizeGstin(supplierGstin));
                d.setInvoiceDate(invoiceDate);
                d.setIgst(igst);
                d.setCgst(cgst);
                d.setSgst(sgst);

                // Get taxable value if available
                Integer taxableValueIdx = col.get(TAXABLE_VALUE);
                if (taxableValueIdx != null) {
                    d.setTaxableValue(parseBigDecimal(r.getCell(taxableValueIdx)));
                }

                // Get invoice value if available
                Integer invoiceValueIdx = col.get(INVOICE_VALUE);
                if (invoiceValueIdx != null) {
                    d.setInvoiceValue(parseBigDecimal(r.getCell(invoiceValueIdx)));
                } else {
                    // Calculate from taxable value + taxes
                    BigDecimal taxable = d.getTaxableValue() != null ? d.getTaxableValue() : BigDecimal.ZERO;
                    BigDecimal totalTax = igst.add(cgst).add(sgst);
                    d.setInvoiceValue(taxable.add(totalTax));
                }

                // Get particulars/legal name
                Integer particularsIdx = col.get(PARTICULARS);
                if (particularsIdx != null) {
                    d.setLegalName(getStringValue(r.getCell(particularsIdx)));
                }

                out.add(d);
            }

            System.out.println("GSTR-2B Parser: Loaded " + out.size() + " invoices");
            return out;
        }
    }

    private boolean isEmptyRow(Row r) {
        if (r == null) return true;
        for (Cell c : r) {
            if (c != null && c.getCellType() != CellType.BLANK) {
                String val = getStringValue(c).trim();
                if (!val.isEmpty()) {
                    return false;
                }
            }
        }
        return true;
    }

    private int findHeaderRow(Sheet s) {
        for (int i = 0; i < 20; i++) {
            Row r = s.getRow(i);
            if (r == null) continue;

            int hit = 0;
            for (Cell c : r) {
                String v = getStringValue(c).toUpperCase();
                if (v.contains("INVOICE") && v.contains("NO")) hit++;
                if (v.contains("SUPPLIER") && v.contains("GST")) hit++;
                if (v.contains("INVOICE") && v.contains("DATE")) hit++;
            }
            if (hit >= 2) {
                System.out.println("Found GSTR-2B header at row: " + i);
                return i;
            }
        }
        throw new RuntimeException("2B header row not found");
    }

    private Map<String, Integer> buildColumnMap(Row h) {
        Map<String, Integer> m = new HashMap<>();
        for (Cell c : h) {
            String v = getStringValue(c).toUpperCase().trim();

            if (v.contains("INVOICE") && v.contains("NO")) {
                m.put(INVOICE_NO, c.getColumnIndex());
                System.out.println("Found INVOICE_NO at column: " + c.getColumnIndex());
            }
            if (v.contains("SUPPLIER") && v.contains("GST")) {
                m.put(SUPPLIER_GSTIN, c.getColumnIndex());
                System.out.println("Found SUPPLIER_GSTIN at column: " + c.getColumnIndex());
            }
            if (v.contains("INVOICE") && v.contains("DATE")) {
                m.put(INVOICE_DATE, c.getColumnIndex());
                System.out.println("Found INVOICE_DATE at column: " + c.getColumnIndex());
            }
            if (v.contains("IGST")) {
                m.put(IGST, c.getColumnIndex());
                System.out.println("Found IGST at column: " + c.getColumnIndex());
            }
            if (v.contains("CGST")) {
                m.put(CGST, c.getColumnIndex());
                System.out.println("Found CGST at column: " + c.getColumnIndex());
            }
            if (v.contains("SGST")) {
                m.put(SGST, c.getColumnIndex());
                System.out.println("Found SGST at column: " + c.getColumnIndex());
            }
            if (v.contains("TAXABLE") && v.contains("VALUE")) {
                m.put(TAXABLE_VALUE, c.getColumnIndex());
                System.out.println("Found TAXABLE_VALUE at column: " + c.getColumnIndex());
            }
            if (v.contains("INVOICE") && v.contains("VALUE")) {
                m.put(INVOICE_VALUE, c.getColumnIndex());
                System.out.println("Found INVOICE_VALUE at column: " + c.getColumnIndex());
            }
            if (v.contains("PARTICULARS") || v.contains("LEGAL") || v.contains("NAME")) {
                m.put(PARTICULARS, c.getColumnIndex());
                System.out.println("Found PARTICULARS at column: " + c.getColumnIndex());
            }
        }
        return m;
    }

    private void validate(Map<String, Integer> m) {
        String[] required = {INVOICE_NO, SUPPLIER_GSTIN, INVOICE_DATE, IGST, CGST, SGST};
        for (String k : required) {
            if (!m.containsKey(k)) {
                throw new RuntimeException("Missing 2B column: " + k);
            }
        }
    }

    private String getStringValue(Cell c) {
        if (c == null) return "";

        switch (c.getCellType()) {
            case STRING:
                return c.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(c)) {
                    return c.getDateCellValue().toString();
                }
                return BigDecimal.valueOf(c.getNumericCellValue())
                        .stripTrailingZeros().toPlainString();
            case BOOLEAN:
                return String.valueOf(c.getBooleanCellValue());
            case FORMULA:
                try {
                    return c.getStringCellValue();
                } catch (Exception e) {
                    try {
                        return String.valueOf(c.getNumericCellValue());
                    } catch (Exception ex) {
                        return "";
                    }
                }
            default:
                return "";
        }
    }

    private BigDecimal parseBigDecimal(Cell c) {
        try {
            if (c == null) return BigDecimal.ZERO;

            switch (c.getCellType()) {
                case NUMERIC:
                    return BigDecimal.valueOf(c.getNumericCellValue());
                case STRING:
                    String val = c.getStringCellValue().trim();
                    if (val.isEmpty()) return BigDecimal.ZERO;
                    return new BigDecimal(val);
                case FORMULA:
                    try {
                        return BigDecimal.valueOf(c.getNumericCellValue());
                    } catch (Exception e) {
                        String val2 = c.getStringCellValue().trim();
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

    private LocalDate parseDate(Cell c) {
        if (c == null) return null;

        try {
            // Try Excel date format first
            if (c.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(c)) {
                return c.getDateCellValue().toInstant()
                        .atZone(ZoneId.systemDefault())
                        .toLocalDate();
            }

            // Try string parsing
            String dateStr = getStringValue(c).trim();
            if (dateStr.isEmpty()) return null;

            // Remove time part if present
            if (dateStr.contains(" ")) {
                dateStr = dateStr.split(" ")[0];
            }

            // Try different date formats
            DateTimeFormatter[] formatters = {
                    DateTimeFormatter.ofPattern("dd/MM/yyyy"),
                    DateTimeFormatter.ofPattern("dd-MM-yyyy"),
                    DateTimeFormatter.ofPattern("yyyy-MM-dd"),
                    DateTimeFormatter.ofPattern("dd-MMM-yyyy"),
                    DateTimeFormatter.ofPattern("d/M/yyyy"),
                    DateTimeFormatter.ofPattern("d-M-yyyy")
            };

            for (DateTimeFormatter fmt : formatters) {
                try {
                    return LocalDate.parse(dateStr, fmt);
                } catch (Exception ignored) {}
            }

            System.err.println("Could not parse date: " + dateStr);
            return null;

        } catch (Exception e) {
            System.err.println("Error parsing date: " + e.getMessage());
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






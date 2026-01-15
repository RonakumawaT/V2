package com.RK8.V2.Service;

import com.RK8.V2.DTO.Gstr2BDTO;
import com.RK8.V2.DTO.PurchaseInvoiceDTO;
import com.RK8.V2.DTO.ReconciliationResult;
import org.springframework.stereotype.Service;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.*;
import java.util.stream.Collectors;
import java.time.YearMonth;
import java.time.LocalDate;


@Service
public class Purchase2BReconciliationService {
    private static final BigDecimal TOLERANCE = new BigDecimal("1.00"); // ₹1 tolerance

    public List<ReconciliationResult> reconcile(
            List<PurchaseInvoiceDTO> purchases,
            List<Gstr2BDTO> gstr2bList
    ) {
        List<ReconciliationResult> results = new ArrayList<>();

        // Build lookup structures
        Map<String, List<Gstr2BDTO>> gstr2bMap = new HashMap<>();
        for (Gstr2BDTO g : gstr2bList) {
            // Multiple keys for fuzzy matching
            String exactKey = createExactKey(g.getSupplierGstin(), g.getInvoiceNo());
            String invoiceOnlyKey = createInvoiceOnlyKey(g.getInvoiceNo());
            String numericKey = createNumericKey(g.getInvoiceNo());

            gstr2bMap.computeIfAbsent(exactKey, k -> new ArrayList<>()).add(g);
            gstr2bMap.computeIfAbsent(invoiceOnlyKey, k -> new ArrayList<>()).add(g);
            gstr2bMap.computeIfAbsent(numericKey, k -> new ArrayList<>()).add(g);
        }

        // Phase 1: Match purchase invoices with 2B
        Set<String> matched2BKeys = new HashSet<>();

        for (PurchaseInvoiceDTO p : purchases) {
            BigDecimal purchaseTax = calculateTax(p.getIgst(), p.getCgst(), p.getSgst());
            YearMonth month = YearMonth.from(p.getInvoiceDate());

            // Try multiple matching strategies in order of confidence
            Gstr2BDTO match = findBestMatch(p, gstr2bMap, purchaseTax);

            if (match != null) {
                BigDecimal gstr2bTax = calculateTax(match.getIgst(), match.getCgst(), match.getSgst());
                matched2BKeys.add(createExactKey(match.getSupplierGstin(), match.getInvoiceNo()));

                if (isTaxMatch(purchaseTax, gstr2bTax)) {
                    results.add(new ReconciliationResult(
                            p.getSupplierGstin(),
                            p.getInvoiceNo(),
                            "MATCHED",
                            purchaseTax,
                            gstr2bTax,
                            BigDecimal.ZERO,
                            buildRemarks(p, match, "Matched"),
                            month
                    ));
                } else {
                    BigDecimal diff = purchaseTax.subtract(gstr2bTax).abs();
                    String status = diff.compareTo(TOLERANCE) <= 0 ?
                            "MATCHED_WITH_TOLERANCE" : "MISMATCH";

                    results.add(new ReconciliationResult(
                            p.getSupplierGstin(),
                            p.getInvoiceNo(),
                            status,
                            purchaseTax,
                            gstr2bTax,
                            purchaseTax.subtract(gstr2bTax).max(BigDecimal.ZERO),
                            buildRemarks(p, match, "Tax amount differs by " + diff),
                            month
                    ));
                }
            } else {
                // No match found
                results.add(new ReconciliationResult(
                        p.getSupplierGstin(),
                        p.getInvoiceNo(),
                        "MISSING_IN_2B",
                        purchaseTax,
                        BigDecimal.ZERO,
                        purchaseTax,
                        "No matching invoice found in GSTR-2B",
                        month
                ));
            }
        }

        // Phase 2: Find invoices in 2B not matched to purchase
        for (Gstr2BDTO g : gstr2bList) {
            String key = createExactKey(g.getSupplierGstin(), g.getInvoiceNo());
            if (!matched2BKeys.contains(key)) {
                BigDecimal gstr2bTax = calculateTax(g.getIgst(), g.getCgst(), g.getSgst());

                results.add(new ReconciliationResult(
                        g.getSupplierGstin(),
                        g.getInvoiceNo(),
                        "MISSING_IN_PURCHASE",
                        BigDecimal.ZERO,
                        gstr2bTax,
                        gstr2bTax,
                        "Invoice present in GSTR-2B but not in purchase register",
                        YearMonth.from(g.getInvoiceDate())
                ));
            }
        }

        return results;
    }

    private Gstr2BDTO findBestMatch(PurchaseInvoiceDTO purchase,
                                    Map<String, List<Gstr2BDTO>> gstr2bMap,
                                    BigDecimal purchaseTax) {
        String purchaseGstin = normalizeGstin(purchase.getSupplierGstin());
        String purchaseInvoice = normalizeInvoice(purchase.getInvoiceNo());
        LocalDate purchaseDate = purchase.getInvoiceDate();

        // Strategy 1: Exact match (GSTIN + Invoice)
        String exactKey = createExactKey(purchaseGstin, purchaseInvoice);
        List<Gstr2BDTO> exactMatches = gstr2bMap.getOrDefault(exactKey, new ArrayList<>());
        if (!exactMatches.isEmpty()) {
            return findClosestTaxMatch(exactMatches, purchaseTax);
        }

        // Strategy 2: Same GSTIN, fuzzy invoice match
        String gstinOnlyKey = purchaseGstin + "|*";
        List<Gstr2BDTO> sameGstin = gstr2bMap.entrySet().stream()
                .filter(e -> e.getKey().startsWith(purchaseGstin + "|"))
                .flatMap(e -> e.getValue().stream())
                .collect(Collectors.toList());

        for (Gstr2BDTO candidate : sameGstin) {
            if (isInvoiceFuzzyMatch(purchaseInvoice, normalizeInvoice(candidate.getInvoiceNo()))) {
                // Also check date proximity (±30 days)
                if (isDateClose(purchaseDate, candidate.getInvoiceDate(), 30)) {
                    return candidate;
                }
            }
        }

        // Strategy 3: Invoice only match (ignore GSTIN)
        String invoiceOnlyKey = createInvoiceOnlyKey(purchaseInvoice);
        List<Gstr2BDTO> invoiceMatches = gstr2bMap.getOrDefault(invoiceOnlyKey, new ArrayList<>());
        if (!invoiceMatches.isEmpty()) {
            return findClosestTaxMatch(invoiceMatches, purchaseTax);
        }

        // Strategy 4: Numeric invoice match
        String numericKey = createNumericKey(purchaseInvoice);
        List<Gstr2BDTO> numericMatches = gstr2bMap.getOrDefault(numericKey, new ArrayList<>());
        if (!numericMatches.isEmpty()) {
            return findClosestTaxMatch(numericMatches, purchaseTax);
        }

        return null;
    }

    private boolean isInvoiceFuzzyMatch(String inv1, String inv2) {
        if (inv1.equals(inv2)) return true;

        // Try with common substitutions
        String norm1 = inv1.replace("O", "0").replace("I", "1").replace("L", "1");
        String norm2 = inv2.replace("O", "0").replace("I", "1").replace("L", "1");
        if (norm1.equals(norm2)) return true;

        // Try removing prefixes/suffixes
        String base1 = extractBaseInvoice(inv1);
        String base2 = extractBaseInvoice(inv2);
        return base1.equals(base2);
    }

    private String extractBaseInvoice(String invoice) {
        // Remove common prefixes/suffixes
        return invoice.replaceAll("^(FY25-26/|GST-25-26/|EP/2025-26/|JE/2025-26/|TIA/T/\\d+/24-25/)", "")
                .replaceAll("/24-25$", "")
                .replaceAll("/25-26$", "");
    }

    private boolean isDateClose(LocalDate date1, LocalDate date2, int daysTolerance) {
        if (date1 == null || date2 == null) return false;
        return Math.abs(date1.toEpochDay() - date2.toEpochDay()) <= daysTolerance;
    }

    private Gstr2BDTO findClosestTaxMatch(List<Gstr2BDTO> candidates, BigDecimal targetTax) {
        return candidates.stream()
                .min((c1, c2) -> {
                    BigDecimal tax1 = calculateTax(c1.getIgst(), c1.getCgst(), c1.getSgst());
                    BigDecimal tax2 = calculateTax(c2.getIgst(), c2.getCgst(), c2.getSgst());
                    BigDecimal diff1 = tax1.subtract(targetTax).abs();
                    BigDecimal diff2 = tax2.subtract(targetTax).abs();
                    return diff1.compareTo(diff2);
                })
                .orElse(null);
    }

    private String buildRemarks(PurchaseInvoiceDTO p, Gstr2BDTO g, String reason) {
        StringBuilder remarks = new StringBuilder(reason);
        remarks.append(" | Purchase GSTIN: ").append(p.getSupplierGstin());
        remarks.append(" | 2B GSTIN: ").append(g.getSupplierGstin());
        remarks.append(" | Purchase Inv: ").append(p.getInvoiceNo());
        remarks.append(" | 2B Inv: ").append(g.getInvoiceNo());
        return remarks.toString();
    }

    // Key creation methods
    private String createExactKey(String gstin, String invoice) {
        return normalizeGstin(gstin) + "|" + normalizeInvoice(invoice);
    }

    private String createInvoiceOnlyKey(String invoice) {
        return "INV|" + normalizeInvoice(invoice);
    }

    private String createNumericKey(String invoice) {
        String numeric = invoice.replaceAll("[^0-9]", "");
        return "NUM|" + (numeric.isEmpty() ? "0" : numeric);
    }

    private BigDecimal calculateTax(BigDecimal igst, BigDecimal cgst, BigDecimal sgst) {
        return safeAdd(safeAdd(igst, cgst), sgst);
    }

    private boolean isTaxMatch(BigDecimal tax1, BigDecimal tax2) {
        return tax1.compareTo(tax2) == 0;
    }

    private BigDecimal safeAdd(BigDecimal a, BigDecimal b) {
        BigDecimal safeA = a == null ? BigDecimal.ZERO : a;
        BigDecimal safeB = b == null ? BigDecimal.ZERO : b;
        return safeA.add(safeB).setScale(2, RoundingMode.HALF_UP);
    }

    private String normalizeGstin(String g) {
        if (g == null) return "";
        return g.trim().toUpperCase().replaceAll("[^A-Z0-9]", "");
    }

    private String normalizeInvoice(String i) {
        if (i == null) return "";
        return i.trim().toUpperCase().replaceAll("\\s+", "");
    }
}



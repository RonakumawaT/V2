package com.RK8.V2.Service;
import com.RK8.V2.DTO.ReconciliationResult;
import org.springframework.stereotype.Service;

import java.math.BigDecimal;
import java.util.*;
import java.util.stream.Collectors;

@Service
public class ReconciliationReportService {

    public Map<String, Object> generateActionReport(List<ReconciliationResult> results) {
        Map<String, Object> report = new HashMap<>();

        // 1. Summary Statistics
        Map<String, Long> statusSummary = results.stream()
                .collect(Collectors.groupingBy(
                        ReconciliationResult::getStatus,
                        Collectors.counting()
                ));

        // 2. ITC Analysis
        BigDecimal totalPurchaseTax = results.stream()
                .map(ReconciliationResult::getPurchaseTax)
                .reduce(BigDecimal.ZERO, BigDecimal::add);

        BigDecimal totalGstr2bTax = results.stream()
                .map(ReconciliationResult::getGstr2bTax)
                .reduce(BigDecimal.ZERO, BigDecimal::add);

        BigDecimal totalItcAtRisk = results.stream()
                .map(ReconciliationResult::getItcAtRisk)
                .reduce(BigDecimal.ZERO, BigDecimal::add);

        BigDecimal totalItcAvailable = results.stream()
                .filter(r -> "MATCHED".equals(r.getStatus()) ||
                        "MATCHED_WITH_TOLERANCE".equals(r.getStatus()) ||
                        "MISSING_IN_PURCHASE".equals(r.getStatus()))
                .map(r -> r.getGstr2bTax().max(r.getPurchaseTax()))
                .reduce(BigDecimal.ZERO, BigDecimal::add);

        // 3. Action Items
        List<Map<String, Object>> actionItems = new ArrayList<>();

        // Action: Add missing purchase invoices
        results.stream()
                .filter(r -> "MISSING_IN_PURCHASE".equals(r.getStatus()))
                .sorted((a, b) -> b.getGstr2bTax().compareTo(a.getGstr2bTax()))
                .forEach(r -> {
                    Map<String, Object> action = new HashMap<>();
                    action.put("action", "ADD_TO_PURCHASE_REGISTER");
                    action.put("priority", "HIGH");
                    action.put("supplierGstin", r.getSupplierGstin());
                    action.put("invoiceNo", r.getInvoiceNo());
                    action.put("invoiceMonth", r.getInvoiceMonth());
                    action.put("taxAmount", r.getGstr2bTax());
                    action.put("reason", "Invoice exists in GSTR-2B but not in purchase register");
                    actionItems.add(action);
                });

        // Action: Review GSTIN mismatches
        results.stream()
                .filter(r -> r.getRemarks().contains("GSTIN mismatch"))
                .forEach(r -> {
                    Map<String, Object> action = new HashMap<>();
                    action.put("action", "VERIFY_SUPPLIER_GSTIN");
                    action.put("priority", "MEDIUM");
                    action.put("purchaseGstin", r.getSupplierGstin());
                    action.put("invoiceNo", r.getInvoiceNo());
                    action.put("purchaseTax", r.getPurchaseTax());
                    action.put("gstr2bTax", r.getGstr2bTax());
                    action.put("remarks", r.getRemarks());
                    actionItems.add(action);
                });

        // 4. Supplier-wise Analysis
        Map<String, Map<String, Object>> supplierAnalysis = new HashMap<>();

        results.forEach(r -> {
            String gstin = r.getSupplierGstin();
            supplierAnalysis.computeIfAbsent(gstin, k -> {
                Map<String, Object> supplier = new HashMap<>();
                supplier.put("gstin", gstin);
                supplier.put("totalInvoices", 0L);
                supplier.put("matched", 0L);
                supplier.put("missingInPurchase", 0L);
                supplier.put("totalTax", BigDecimal.ZERO);
                supplier.put("itcAtRisk", BigDecimal.ZERO);
                return supplier;
            });

            Map<String, Object> supplier = supplierAnalysis.get(gstin);
            supplier.put("totalInvoices", (Long) supplier.get("totalInvoices") + 1);

            if ("MATCHED".equals(r.getStatus()) || "MATCHED_WITH_TOLERANCE".equals(r.getStatus())) {
                supplier.put("matched", (Long) supplier.get("matched") + 1);
            } else if ("MISSING_IN_PURCHASE".equals(r.getStatus())) {
                supplier.put("missingInPurchase", (Long) supplier.get("missingInPurchase") + 1);
            }

            BigDecimal totalTax = ((BigDecimal) supplier.get("totalTax")).add(r.getGstr2bTax());
            supplier.put("totalTax", totalTax);

            BigDecimal itcAtRisk = ((BigDecimal) supplier.get("itcAtRisk")).add(r.getItcAtRisk());
            supplier.put("itcAtRisk", itcAtRisk);
        });

        // 5. Monthly Analysis
        Map<String, Map<String, Object>> monthlyAnalysis = new HashMap<>();

        results.forEach(r -> {
            String month = r.getInvoiceMonth().toString();
            monthlyAnalysis.computeIfAbsent(month, k -> {
                Map<String, Object> monthly = new HashMap<>();
                monthly.put("month", month);
                monthly.put("totalInvoices", 0L);
                monthly.put("totalTax", BigDecimal.ZERO);
                monthly.put("itcClaimed", BigDecimal.ZERO);
                monthly.put("itcAtRisk", BigDecimal.ZERO);
                return monthly;
            });

            Map<String, Object> monthly = monthlyAnalysis.get(month);
            monthly.put("totalInvoices", (Long) monthly.get("totalInvoices") + 1);

            BigDecimal totalTax = ((BigDecimal) monthly.get("totalTax")).add(r.getGstr2bTax());
            monthly.put("totalTax", totalTax);

            if ("MISSING_IN_PURCHASE".equals(r.getStatus())) {
                BigDecimal itcAtRisk = ((BigDecimal) monthly.get("itcAtRisk")).add(r.getItcAtRisk());
                monthly.put("itcAtRisk", itcAtRisk);
            } else {
                BigDecimal itcClaimed = ((BigDecimal) monthly.get("itcClaimed")).add(r.getPurchaseTax());
                monthly.put("itcClaimed", itcClaimed);
            }
        });

        // 6. Top Missing Invoices by Value
        List<Map<String, Object>> topMissingByValue = results.stream()
                .filter(r -> "MISSING_IN_PURCHASE".equals(r.getStatus()))
                .sorted((a, b) -> b.getGstr2bTax().compareTo(a.getGstr2bTax()))
                .limit(10)
                .map(r -> {
                    Map<String, Object> invoice = new HashMap<>();
                    invoice.put("supplierGstin", r.getSupplierGstin());
                    invoice.put("invoiceNo", r.getInvoiceNo());
                    invoice.put("invoiceMonth", r.getInvoiceMonth());
                    invoice.put("taxAmount", r.getGstr2bTax());
                    return invoice;
                })
                .collect(Collectors.toList());

        // 7. Compliance Score
        long totalInvoices = results.size();
        long matchedInvoices = statusSummary.getOrDefault("MATCHED", 0L) +
                statusSummary.getOrDefault("MATCHED_WITH_TOLERANCE", 0L);

        double complianceScore = totalInvoices > 0 ?
                (matchedInvoices * 100.0) / totalInvoices : 100.0;

        // Build final report
        report.put("summary", Map.of(
                "totalInvoices", totalInvoices,
                "matchedInvoices", matchedInvoices,
                "complianceScore", String.format("%.2f%%", complianceScore),
                "itcAvailable", totalItcAvailable,
                "itcClaimed", totalPurchaseTax,
                "itcAtRisk", totalItcAtRisk,
                "itcUnclaimed", totalItcAvailable.subtract(totalPurchaseTax)
        ));

        report.put("statusBreakdown", statusSummary);
        report.put("actionItems", actionItems);
        report.put("supplierAnalysis", new ArrayList<>(supplierAnalysis.values()));
        report.put("monthlyAnalysis", new ArrayList<>(monthlyAnalysis.values()));
        report.put("topMissingByValue", topMissingByValue);

        return report;
    }
}

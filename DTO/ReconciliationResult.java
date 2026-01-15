package com.RK8.V2.DTO;

import lombok.AllArgsConstructor;
import lombok.Data;

import java.math.BigDecimal;
import java.time.YearMonth;

@Data
@AllArgsConstructor
public class ReconciliationResult {
    private String supplierGstin;
    private String invoiceNo;
    private String status;
    private BigDecimal purchaseTax;
    private BigDecimal gstr2bTax;
    private BigDecimal itcAtRisk;
    private String remarks;
    private YearMonth invoiceMonth;
}



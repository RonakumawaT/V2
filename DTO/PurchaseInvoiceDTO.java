package com.RK8.V2.DTO;

import lombok.Data;

import java.math.BigDecimal;
import java.time.LocalDate;

@Data
public class PurchaseInvoiceDTO {

    private String supplierGstin;
    private String invoiceNo;
    private LocalDate invoiceDate;

    private BigDecimal igst = BigDecimal.ZERO;
    private BigDecimal cgst = BigDecimal.ZERO;
    private BigDecimal sgst = BigDecimal.ZERO;

    private String particulars;
    private BigDecimal grossTotal = BigDecimal.ZERO;
}




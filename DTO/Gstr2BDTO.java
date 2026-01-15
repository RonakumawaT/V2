package com.RK8.V2.DTO;

import lombok.Data;

import java.math.BigDecimal;
import java.time.LocalDate;

@Data
public class Gstr2BDTO {
    private BigDecimal taxableValue;
    private BigDecimal invoiceValue;
    private String supplierGstin;
    private String invoiceNo;
    private LocalDate invoiceDate;
    private BigDecimal igst;
    private BigDecimal cgst;
    private BigDecimal sgst;
    private String legalName;
}


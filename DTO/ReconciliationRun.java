package com.RK8.V2.DTO;

import lombok.Data;

import java.math.BigDecimal;
import java.time.LocalDateTime;

@Data
public class ReconciliationRun {


    private Long id;

    private String clientGstin;
    private String period; // 2024-12

    private LocalDateTime runAt;

    private Long totalInvoices;
    private Long matchedCount;
    private Long mismatchCount;
    private Long missingCount;

    private BigDecimal itcAtRisk;

    // getters & setters
}


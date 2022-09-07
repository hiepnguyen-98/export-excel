package com.example.demo.dto;

import com.opencsv.bean.CsvBindByPosition;
import lombok.Data;

@Data
public class KpiEmployee {
    @CsvBindByPosition(position = 0)
    private String nameEmp;

    @CsvBindByPosition(position = 1)
    private Double kpi;

    @CsvBindByPosition(position = 2)
    private String position;
}

package com.example.demo.excel;

import com.example.demo.dto.KpiEmployee;
import com.opencsv.bean.CsvToBeanBuilder;
import lombok.extern.slf4j.Slf4j;

import java.io.FileReader;
import java.util.ArrayList;
import java.util.List;

@Slf4j
public class CSVUtils {

    public static List<KpiEmployee> importSymbolCSV(String filePath) {
        try {
            return (List<KpiEmployee>) new CsvToBeanBuilder(new FileReader(filePath))
                    .withSkipLines(1)
                    .withType(KpiEmployee.class)
                    .build().parse();

        } catch (Exception e) {
            log.error("CSV Utils throws exception: {}", e.getMessage());
        }

        return new ArrayList<>();
    }
}

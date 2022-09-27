package com.example.demo.service;

import java.io.IOException;

public interface ExportExcelService {

    void createExcel(String empKpiPath, String folderPath) throws IOException;
}

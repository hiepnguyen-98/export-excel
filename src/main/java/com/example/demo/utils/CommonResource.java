package com.example.demo.utils;

import lombok.Getter;
import lombok.RequiredArgsConstructor;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

@Component
@RequiredArgsConstructor
public class CommonResource {

    @Getter
    @Value("${folder.template.excel.path}")
    private String templateExcelFolderPath;

    @Getter
    @Value("${folder.export.path}")
    private String exportFolderPath;
}

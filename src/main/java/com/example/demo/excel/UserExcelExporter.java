package com.example.demo.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Component;

import javax.annotation.PostConstruct;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

@Component
public class UserExcelExporter {

    @PostConstruct
    public void createExcel() {
        try {
            // pre-process
            FileInputStream originFile = new FileInputStream(new File("C:\\Users\\MY PC\\Downloads\\KPI DEV.xlsx"));

            FileOutputStream out = new FileOutputStream(new File("C:\\Users\\MY PC\\Downloads\\test.xlsx"));

            XSSFWorkbook workbookOrigin = new XSSFWorkbook(originFile);

            XSSFWorkbook workbook = new XSSFWorkbook();

            XSSFSheet sourceSheet = workbookOrigin.getSheetAt(0);

            //Create a blank sheet
            XSSFSheet destinationWorksheet = workbook.createSheet(sourceSheet.getSheetName());

            //This data needs to be written (Object[])

            Integer lastRow = sourceSheet.getLastRowNum();
            Integer firstRow = sourceSheet.getFirstRowNum();

            for (int i = firstRow; i < lastRow; i++) {
                copyRow(sourceSheet.getRow(i), destinationWorksheet, workbook, sourceSheet);
            }
//            for(int colNum = 0; colNum<row.getLastCellNum();colNum++)
//                workbook.getSheetAt(0).autoSizeColumn(colNum);

            //Write the workbook in file system
            workbook.write(out);
            out.close();
            System.out.println("howtodoinjava_demo.xlsx written successfully on disk.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void copyRow(XSSFRow sourceRow, XSSFSheet destinationWorksheet, XSSFWorkbook workbook, XSSFSheet sourceSheet) {
        XSSFRow newRow = destinationWorksheet.createRow(sourceRow.getRowNum());
        // Loop through source columns to add to new row
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            // Grab a copy of the old/new cell
            XSSFCell oldCell = sourceRow.getCell(i);
            XSSFCell newCell = newRow.createCell(i);
            // If the old cell is null jump to next cell
            if (oldCell == null) {
                continue;
            }

            // Copy style from old cell and apply to new cell
            XSSFCellStyle newCellStyle = workbook.createCellStyle();
            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
            newCell.setCellStyle(newCellStyle);

            // If there is a cell comment, copy
            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }

            // Set the cell data type
            if (oldCell.getCellType() == CellType.FORMULA) {
                newCell.setCellFormula(null);
            } else {
                newCell.setCellType(oldCell.getCellType());

            }

            // Set the cell data value
            switch (oldCell.getCellType()) {
                case BLANK:// Cell.CELL_TYPE_BLANK:
                    newCell.setCellValue(oldCell.getStringCellValue());
                    break;
                case BOOLEAN:
                    newCell.setCellValue(oldCell.getBooleanCellValue());
                    break;
                case NUMERIC:
                    newCell.setCellValue(oldCell.getNumericCellValue());
                    break;
                case STRING:
                    newCell.setCellValue(oldCell.getRichStringCellValue());
                    break;
                default:
                    break;
            }

            // If there are are any merged regions in the source row, copy to new row

        }
        for (int index = 0; index < sourceSheet.getNumMergedRegions(); index++) {
            if (index == 52){
                int a = 0;
            }
            CellRangeAddress cellRangeAddress = sourceSheet.getMergedRegion(index);
            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
                CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
                        (newRow.getRowNum() + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())),
                        cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastColumn());
                destinationWorksheet.addMergedRegion(newCellRangeAddress);
            }
        }
    }


}
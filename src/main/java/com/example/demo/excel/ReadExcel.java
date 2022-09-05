//package com.example.demo.excel;
//
//
//import org.apache.poi.ss.usermodel.CellType;
//import org.apache.poi.xssf.usermodel.*;
//import org.springframework.stereotype.Component;
//
//import javax.annotation.PostConstruct;
//import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileOutputStream;
//import java.io.IOException;
//
//@Component
//public class ReadExcel {
//
//    @PostConstruct
//    public void readExcel() {
//        try {
//            FileInputStream originFile = new FileInputStream(new File("C:\\Users\\MY PC\\Downloads\\KPI DEV.xlsx"));
//            File testXlsx = createNewFileExcel();
//            FileInputStream desFile = new FileInputStream(new File("C:\\Users\\MY PC\\Downloads\\test.xlsx"));
//
//            //Create Workbook instance holding reference to .xlsx file
//            XSSFWorkbook workbook = new XSSFWorkbook(originFile);
//            XSSFWorkbook workbookDes = new XSSFWorkbook(desFile);
//            //Get first/desired sheet from the workbook
//
//            copyData(workbook, workbookDes);
//
//            originFile.close();
//            desFile.close();
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//    }
//
//    private File createNewFileExcel() {
//        String filename = "C:\\Users\\MY PC\\Downloads\\test.xlsx";
//        try {
//            XSSFWorkbook workbook = new XSSFWorkbook();
//
//            FileOutputStream fileOut = new FileOutputStream(filename);
//            workbook.write(fileOut);
//            fileOut.close();
//            workbook.close();
//            System.out.println("Your excel file has been generated!");
//
//        } catch (Exception ex) {
//            System.out.println(ex);
//        }
//        return new File(filename);
//    }
//
//    private void copyData(XSSFWorkbook sourceWorksheet, XSSFWorkbook destinationworkbook) throws IOException {
//
//
//        XSSFSheet sourceSheet = sourceWorksheet.getSheetAt(0);
//        XSSFSheet destinationWorksheet = destinationworkbook.createSheet();
//
//        Integer lastRow = sourceSheet.getLastRowNum();
//        Integer firstRow = sourceSheet.getFirstRowNum();
//
//        for (int i = firstRow; i < lastRow; i++) {
//            copyRow(sourceSheet.getRow(i), destinationWorksheet, destinationworkbook);
//        }
//
//
//        XSSFRow xssfRow = destinationworkbook.getSheetAt(0).getRow(0);
//        xssfRow.getCell(0).setCellValue("Hiepdz");
//        sourceWorksheet.close();
//        destinationworkbook.close();
//
//    }
//
//    public void copyRow(XSSFRow sourceRow, XSSFSheet destinationWorksheet, XSSFWorkbook workbook) {
//        XSSFRow newRow = destinationWorksheet.createRow(sourceRow.getRowNum());
//        // Loop through source columns to add to new row
//        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
//            // Grab a copy of the old/new cell
//            XSSFCell oldCell = sourceRow.getCell(i);
//            XSSFCell newCell = newRow.createCell(i);
//            // If the old cell is null jump to next cell
//            if (oldCell == null) {
//                continue;
//            }
//
//            // Copy style from old cell and apply to new cell
//            XSSFCellStyle newCellStyle = workbook.createCellStyle();
//            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
//            newCell.setCellStyle(newCellStyle);
//
//            // If there is a cell comment, copy
//            if (oldCell.getCellComment() != null) {
//                newCell.setCellComment(oldCell.getCellComment());
//            }
//
//            // If there is a cell hyperlink, copy
//            if (oldCell.getHyperlink() != null) {
//                newCell.setHyperlink(oldCell.getHyperlink());
//            }
//
//            // Set the cell data type
//            if (oldCell.getCellType() == CellType.FORMULA){
//                newCell.setCellFormula(null);
//            } else {
//                newCell.setCellType(oldCell.getCellType());
//
//            }
//            // Set the cell data value
//            switch (oldCell.getCellType()) {
//                case BLANK:// Cell.CELL_TYPE_BLANK:
//                    newCell.setCellValue(oldCell.getStringCellValue());
//                    break;
//                case BOOLEAN:
//                    newCell.setCellValue(oldCell.getBooleanCellValue());
//                    break;
//                case NUMERIC:
//                    newCell.setCellValue(oldCell.getNumericCellValue());
//                    break;
//                case STRING:
//                    newCell.setCellValue(oldCell.getRichStringCellValue());
//                    break;
//                default:
//                    break;
//            }
//
//            // If there are are any merged regions in the source row, copy to new row
////        for (int i = 0; i < sourceWorksheet.getNumMergedRegions(); i++) {
////            CellRangeAddress cellRangeAddress = sourceWorksheet.getMergedRegion(i);
////            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
////                CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
////                        (newRow.getRowNum() + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())),
////                        cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastColumn());
////                destinationWorksheet.addMergedRegion(newCellRangeAddress);
////            }
////        }
//        }
//
//    }
//
//    public void func(){
//        XSSFRow sourceRow = sourceWorksheet.getRow(sourceRowNum);
//        XSSFRow newRow = destinationWorksheet.createRow(destinationRowNum);
//        // Loop through source columns to add to new row
//        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
//            // Grab a copy of the old/new cell
//            XSSFCell oldCell = sourceRow.getCell(i);
//            XSSFCell newCell = newRow.createCell(i);
//            // If the old cell is null jump to next cell
//            if (oldCell == null) {
//                continue;
//            }
//
//            // Copy style from old cell and apply to new cell
//            XSSFCellStyle newCellStyle = workbook.createCellStyle();
//            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
//            newCell.setCellStyle(newCellStyle);
//
//            // If there is a cell comment, copy
//            if (oldCell.getCellComment() != null) {
//                newCell.setCellComment(oldCell.getCellComment());
//            }
//
//            // If there is a cell hyperlink, copy
//            if (oldCell.getHyperlink() != null) {
//                newCell.setHyperlink(oldCell.getHyperlink());
//            }
//
//            // Set the cell data type
//            newCell.setCellType(oldCell.getCellTypeEnum());
//
//            // Set the cell data value
//            switch (oldCell.getCellTypeEnum()) {
//                case BLANK:// Cell.CELL_TYPE_BLANK:
//                    newCell.setCellValue(oldCell.getStringCellValue());
//                    break;
//                case BOOLEAN:
//                    newCell.setCellValue(oldCell.getBooleanCellValue());
//                    break;
//                case FORMULA:
//                    newCell.setCellFormula(oldCell.getCellFormula());
//                    break;
//                case NUMERIC:
//                    newCell.setCellValue(oldCell.getNumericCellValue());
//                    break;
//                case STRING:
//                    newCell.setCellValue(oldCell.getRichStringCellValue());
//                    break;
//                default:
//                    break;
//            }
//        }
//
//        // If there are are any merged regions in the source row, copy to new row
//        for (int i = 0; i < sourceWorksheet.getNumMergedRegions(); i++) {
//            CellRangeAddress cellRangeAddress = sourceWorksheet.getMergedRegion(i);
//            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
//                CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
//                        (newRow.getRowNum() + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())),
//                        cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastColumn());
//                destinationWorksheet.addMergedRegion(newCellRangeAddress);
//            }
//        }
//    }
//}

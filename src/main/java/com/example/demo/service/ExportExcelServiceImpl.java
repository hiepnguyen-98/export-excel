package com.example.demo.service;

import com.example.demo.dto.KpiEmployee;
import com.example.demo.utils.Constant;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
@Slf4j
public class ExportExcelServiceImpl implements ExportExcelService {

    HashMap<Integer, XSSFCellStyle> styleExcelDev = new HashMap<>();
    HashMap<Integer, XSSFCellStyle> styleExcelTest = new HashMap<>();

    public void createExcel(String employeeKpiFilePath) throws IOException {
        List<KpiEmployee> kpiEmployees = new ArrayList<>();
        List<KpiEmployee> kpiTest = new ArrayList<>();

        loadKpiEmployees(kpiEmployees, kpiTest, employeeKpiFilePath);
        createKpi(kpiEmployees, Constant.PositionEmployee.DEV.getType());
        createKpi(kpiTest, Constant.PositionEmployee.TESTER.getType());
        log.info("Generate excel successfully");
    }

    private void loadKpiEmployees(List<KpiEmployee> kpiEmployees, List<KpiEmployee> kpiTest, String employeeKpiFilePath) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(new File(employeeKpiFilePath)));
        XSSFSheet worksheet = workbook.getSheetAt(0);

        for (int i = 1; i < worksheet.getPhysicalNumberOfRows(); i++) {
            KpiEmployee tempStudent = new KpiEmployee();

            XSSFRow row = worksheet.getRow(i);

            tempStudent.setNameEmp(row.getCell(0).getStringCellValue());
            tempStudent.setKpi(row.getCell(1).getNumericCellValue());
            tempStudent.setPosition(row.getCell(2).getStringCellValue());
            if (StringUtils.equals(Constant.PositionEmployee.DEV.getValue(), tempStudent.getPosition())) {
                kpiEmployees.add(tempStudent);
            }
            if (StringUtils.equals(Constant.PositionEmployee.TESTER.getValue(), tempStudent.getPosition())) {
                kpiTest.add(tempStudent);
            }
        }
    }

    private void createKpi(List<KpiEmployee> kpiEmployees, Integer typePosition) {
        kpiEmployees = kpiEmployees.stream()
                .sorted(Comparator.comparing(KpiEmployee::getKpi))
                .collect(Collectors.toList());
        InputStream in = null;
        FileOutputStream out = null;
        try {
            // pre-process
            if (Constant.PositionEmployee.DEV.getType().equals(typePosition)) {
                in = this.getClass().getResourceAsStream("/excel/KPI-DEV.xlsx");
                out = new FileOutputStream(new File("C:\\Users\\admin\\Downloads\\KPI_DEV.xlsx"));
            } else if (Constant.PositionEmployee.TESTER.getType().equals(typePosition)) {
                in = this.getClass().getResourceAsStream("/excel/KPI-TEST.xlsx");
                out = new FileOutputStream(new File("C:\\Users\\admin\\Downloads\\KPI_TEST.xlsx"));
            }

            XSSFWorkbook sourceWorkbook = new XSSFWorkbook(in);

            XSSFWorkbook destinationWorkbook = new XSSFWorkbook();

            int version = 1;
            Double oldKpi = 0d;
            for (KpiEmployee employee : kpiEmployees) {

                // reset version if new kpi
                if (!oldKpi.equals(employee.getKpi())) {
                    oldKpi = employee.getKpi();
                    version = 1;
                }

                // Process create new sheet
                XSSFSheet sourceSheet = sourceWorkbook.getSheet(employee.getKpi() + Constant.DOT + version);
                while (Objects.isNull(sourceSheet)) {
                    if (version == 8) {
                        version = 1;
                    } else {
                        version = version + 1;
                    }
                    sourceSheet = sourceWorkbook.getSheet(employee.getKpi() + Constant.DOT + version);

                }
                XSSFSheet destWorksheet = destinationWorkbook.createSheet(employee.getNameEmp());

                for (int i = sourceSheet.getFirstRowNum(); i < sourceSheet.getLastRowNum(); i++) {
                    copyRow(sourceSheet.getRow(i), destWorksheet, destinationWorkbook, sourceSheet, employee, typePosition);
                }

                // Each employee will correspond to one version sheet
                version++;
            }

            //Write the workbook in file system
            destinationWorkbook.write(out);
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void copyRow(XSSFRow sourceRow, XSSFSheet destinationSheet, XSSFWorkbook destinationWorkbook, XSSFSheet sourceSheet, KpiEmployee employee, Integer typePosition) {

        // Create new row
        XSSFRow newRow = destinationSheet.createRow(sourceRow.getRowNum());

        // copy all cell from source row into new row
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            // Grab a copy of the old/new cell
            XSSFCell oldCell = sourceRow.getCell(i);
            XSSFCell newCell = newRow.createCell(i);

            if (oldCell == null) {
                continue;
            }

            // performance optimization: don't create new style for every cell. if it's already available, use it
            getCellStyle(oldCell, destinationWorkbook, newCell, typePosition);

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
                newCell.setCellFormula(oldCell.getCellFormula());
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
            if (CellType.STRING.equals(oldCell.getCellType()) && StringUtils.equals(oldCell.getStringCellValue(), "time")) {
                newCell.setCellValue("Tháng 07/2021");
            }
            if (CellType.STRING.equals(oldCell.getCellType()) && StringUtils.equals(oldCell.getStringCellValue(), "employee")) {
                newCell.setCellValue(employee.getNameEmp());
            }
        }


        // If there are are any merged regions in the source row, copy to new row
        for (int index = 0; index < sourceSheet.getNumMergedRegions(); index++) {
            CellRangeAddress cellRangeAddress = sourceSheet.getMergedRegion(index);
            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
                CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
                        (newRow.getRowNum() + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())),
                        cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastColumn());
                destinationSheet.addMergedRegion(newCellRangeAddress);
            }
        }
    }

    private void getCellStyle(XSSFCell oldCell, XSSFWorkbook destinationWorkbook, XSSFCell newCell, Integer typePosition) {
        if (Constant.PositionEmployee.DEV.getType().equals(typePosition)) {
            int styleHashCode = oldCell.getCellStyle().hashCode();
            XSSFCellStyle newCellStyle = styleExcelDev.get(styleHashCode);
            if (newCellStyle == null) {
                newCellStyle = destinationWorkbook.createCellStyle();
                newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
                styleExcelDev.put(styleHashCode, newCellStyle);
            }
            newCell.setCellStyle(newCellStyle);
        } else if (Constant.PositionEmployee.TESTER.getType().equals(typePosition)) {
            int styleHashCode = oldCell.getCellStyle().hashCode();
            XSSFCellStyle newCellStyle = styleExcelTest.get(styleHashCode);
            if (newCellStyle == null) {
                newCellStyle = destinationWorkbook.createCellStyle();
                newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
                styleExcelTest.put(styleHashCode, newCellStyle);
            }
            newCell.setCellStyle(newCellStyle);
        }
    }
}

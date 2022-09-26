package com.example.demo.utils;

import lombok.AllArgsConstructor;

public class Constant {


    @AllArgsConstructor
    public enum PositionEmployee {
        DEV("DEV", 1),
        TESTER("Tester", 2),
        BOM("BOM", 3),
        ACC("ACC", 4),
        BA("BA", 5),
        IT_SUPPORT("IT Support", 6),
        BO("BO", 7),
        HR("HR", 8);

        private String value;
        private Integer type;

        public Integer getType() {
            return type;
        }

        public void setType(Integer type) {
            this.type = type;
        }

        public String getValue() {
            return value;
        }

        public void setValue(String value) {
            this.value = value;
        }
    }


    public static String DOT = ".";
    public static String KPI_DEV_FILE_NAME = "KPI-DEV.xlsx";
    public static String KPI_TEST_FILE_NAME = "KPI-TEST.xlsx";
    public static String EXPORT_KPI_DEV_FILE_NAME = "EXPORT_KPI_DEV.xlsx";
    public static String EXPORT_KPI_TEST_FILE_NAME = "EXPORT_KPI_TEST.xlsx";
}

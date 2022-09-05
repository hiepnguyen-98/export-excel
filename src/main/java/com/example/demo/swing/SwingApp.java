//package com.example.demo.swing;
//
//import org.springframework.boot.autoconfigure.SpringBootApplication;
//import org.springframework.boot.builder.SpringApplicationBuilder;
//import org.springframework.context.ConfigurableApplicationContext;
//
//import javax.swing.*;
//
//@SpringBootApplication
//public class SwingApp extends JFrame {
//
//    public SwingApp() {
//        initUI();
//    }
//
//    private void initUI() {
//        String[] columns = new String[]{
//                "Id", "Name", "Hourly Rate", "Part Time"
//        };
//        Object[][] data = new Object[][]{
//                {1, "John", 40.0, false},
//                {2, "Rambo", 70.0, false},
//                {3, "Zorro", 60.0, true},
//        };
//        //create table with data
//        JTable table = new JTable(data, columns);
//        displayGui dg = new displayGui(table);
//        dg.setVisible(true);
//    }
//
//    public static void main(String[] args) {
//        ConfigurableApplicationContext ctx = new SpringApplicationBuilder(SwingApp.class)
//                .headless(false).run(args);
//    }
//}

//package com.example.demo;
//
//import javax.swing.*;
//import java.awt.*;
//import java.awt.event.ActionEvent;
//
//public class CustomerTable extends JFrame {
//
//    public CustomerTable() {
//        //headers for the table
//        String[] columns = new String[]{
//                "Id", "Name", "Hourly Rate", "Part Time"
//        };
//
//        //actual data for the table in a 2d array
//        Object[][] data = new Object[][]{
//                {1, "John", 40.0, false},
//                {2, "Rambo", 70.0, false},
//                {3, "Zorro", 60.0, true},
//        };
//        //create table with data
//        JTable table = new JTable(data, columns);
//
//        JButton quitButton = new JButton("Quita");
//
//        quitButton.addActionListener((ActionEvent event) -> {
//            System.exit(0);
//        });
//
//        createLayout(quitButton);
//        setTitle("Quita button");
//        setSize(500, 200);
//        setLocationRelativeTo(null);
//        setDefaultCloseOperation(EXIT_ON_CLOSE);
//
//
//        //add the table to the frame
//        this.add(new JScrollPane(table));
//        this.add(new JScrollPane(quitButton));
//
//        this.setTitle("Table Example");
//        this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//        this.pack();
//        this.setVisible(true);
//
//
//    }
//
//    private void createLayout(JComponent... arg) {
//
//        Container pane = getContentPane();
//        GroupLayout gl = new GroupLayout(pane);
//        pane.setLayout(gl);
//
//        gl.setAutoCreateContainerGaps(true);
//
//        gl.setHorizontalGroup(gl.createSequentialGroup()
//                .addComponent(arg[0])
//        );
//
//        gl.setVerticalGroup(gl.createSequentialGroup()
//                .addComponent(arg[0])
//        );
//    }
//}

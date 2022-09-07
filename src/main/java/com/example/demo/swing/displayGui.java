package com.example.demo.swing;

import javax.swing.*;
import java.awt.*;

public class displayGui extends JFrame {
    private JPanel topPanel;
    private JPanel btnPanel;
    private JScrollPane scrollPane;

    public displayGui(JTable tbl) {
        setTitle("Company Record Application");
        setSize(300, 200);
        setBackground(Color.gray);


        topPanel = new JPanel();
        btnPanel = new JPanel();

        topPanel.setLayout(new BorderLayout());
        getContentPane().add(topPanel, BorderLayout.CENTER);
        getContentPane().add(btnPanel, BorderLayout.SOUTH);
        scrollPane = new JScrollPane(tbl);
        topPanel.add(scrollPane, BorderLayout.CENTER);
        JButton addButton = new JButton("ADD");
        JButton delButton = new JButton("DELETE");
        JButton saveButton = new JButton("SAVE");

        btnPanel.add(addButton);
        btnPanel.add(delButton);

    }
}
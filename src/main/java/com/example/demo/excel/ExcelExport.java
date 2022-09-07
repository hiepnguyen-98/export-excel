package com.example.demo.excel;

import com.example.demo.service.ExportExcelService;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import javax.annotation.PostConstruct;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.util.Objects;

@Component
@RequiredArgsConstructor
@Slf4j
public class ExcelExport {

    private final ExportExcelService exportExcelService;

    @PostConstruct
    private void createWindow() {
        JFrame frame = new JFrame("Export KPI excel");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        createUI(frame);
        frame.setSize(560, 200);
        frame.setLocationRelativeTo(null);
        frame.setVisible(true);
    }

    private void createUI(final JFrame frame) {
        JPanel panel = new JPanel();
        LayoutManager layout = new FlowLayout();
        panel.setLayout(layout);

        JButton button = new JButton("Import file");
        final JLabel label = new JLabel();

        button.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setMultiSelectionEnabled(false);
            FileNameExtensionFilter excelFilter = new FileNameExtensionFilter("excel", "xlsx");
            fileChooser.setFileFilter(excelFilter);

            int option = fileChooser.showOpenDialog(frame);
            if (option == JFileChooser.APPROVE_OPTION) {
                File file = fileChooser.getSelectedFile();
                if (Objects.isNull(file)) {
                    label.setText("File selected is empty");
                } else {
                    String pathFile = file.getAbsolutePath();
                    try {
                        exportExcelService.createExcel(pathFile);
                    } catch (IOException ioException) {
                        log.error("Create excel failed: {}", ioException.getMessage());
                    }
                    label.setText("File name: " + file.getName() + "  , File path: " + pathFile);
                }

            } else {
                label.setText("Open command canceled");
            }
        });

        panel.add(button);
        panel.add(label);
        frame.getContentPane().add(panel, BorderLayout.CENTER);
    }
}

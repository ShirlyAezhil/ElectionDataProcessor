package com.example.demo;

import org.apache.poi.ss.usermodel.*;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.util.Vector;

public class ExcelViewerFrame extends JFrame {

    private JTable table;

    public ExcelViewerFrame() {
        setTitle("Excel Viewer");
        setSize(800, 600);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        // Initialize empty table
        table = new JTable();
        JScrollPane scrollPane = new JScrollPane(table);
        add(scrollPane, BorderLayout.CENTER);

        // Menu Bar
        JMenuBar menuBar = new JMenuBar();
        JMenu fileMenu = new JMenu("File");
        JMenuItem openItem = new JMenuItem("Open Excel File");
        openItem.addActionListener(e -> {
            try {
                openExcelFile();
            } catch (Exception e1) {
                e1.printStackTrace();
            }
        });
        fileMenu.add(openItem);
        menuBar.add(fileMenu);
        setJMenuBar(menuBar);
    }

    private void openExcelFile() throws Exception {
        JFileChooser fileChooser = new JFileChooser();
        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            displayExcelFile(selectedFile);
        }
    }

    private void displayExcelFile(File file) throws Exception {
        FileInputStream fis = new FileInputStream(file);
        Workbook workbook = WorkbookFactory.create(fis);
        Sheet sheet = workbook.getSheetAt(0);

        // Extract column headers
        Row headerRow = sheet.getRow(0);
        Vector<String> columnNames = new Vector<>();
        columnNames.add("Select"); // Checkbox column
        for (Cell cell : headerRow) {
            columnNames.add(cell.toString());
        }

        // Extract data rows (start from row 1 to skip header row)
        Vector<Vector<Object>> data = new Vector<>();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            Vector<Object> rowData = new Vector<>();
            rowData.add(false); // Checkbox for each data row

            for (int j = 0; j < headerRow.getLastCellNum(); j++) {
                Cell cell = row.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                switch (cell.getCellType()) {
                    case STRING -> rowData.add(cell.getStringCellValue());
                    case NUMERIC -> rowData.add(cell.getNumericCellValue());
                    case BOOLEAN -> rowData.add(cell.getBooleanCellValue());
                    case FORMULA -> rowData.add(cell.getCellFormula());
                    default -> rowData.add("");
                }
            }
            data.add(rowData);
        }

        workbook.close();
        fis.close();

        // Create a model where only the first column is editable and rendered as checkbox
        DefaultTableModel model = new DefaultTableModel(data, columnNames) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return column == 0; // Only checkbox column editable
            }

            @Override
            public Class<?> getColumnClass(int columnIndex) {
                if (columnIndex == 0) {
                    return Boolean.class; // Render as checkbox
                }
                return super.getColumnClass(columnIndex);
            }
        };

        // Set the model to the existing table
        table.setModel(model);
    }
}

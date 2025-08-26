package com.example.demo;

import org.apache.poi.ss.usermodel.*;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.util.Vector;

public class ExcelViewerFrame extends JFrame {

    private JTable excelTable;
    private JTable electionTable;

    public ExcelViewerFrame() {
        setTitle("Excel Viewer");
        setSize(1200, 600);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        // Initialize empty Excel table
        excelTable = new JTable();

        // Initialize second table with fixed headers
        String[] electionHeaders = {"Election", "Assembly", "Party", "Name", "Votes"};
        DefaultTableModel electionModel = new DefaultTableModel(electionHeaders, 0);
        electionTable = new JTable(electionModel);

        // Create scroll panes
        JScrollPane excelScrollPane = new JScrollPane(excelTable);
        JScrollPane electionScrollPane = new JScrollPane(electionTable);

        // Add both tables side by side using JSplitPane
        JSplitPane splitPane = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT, excelScrollPane, electionScrollPane);
        splitPane.setDividerLocation(600);
        add(splitPane, BorderLayout.CENTER);

        // Bottom button panel
        JPanel bottomPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 10, 5));
        JButton removeColumnBtn = new JButton("Remove Column");
        JButton removeRowBtn = new JButton("Remove Row");
        JButton addPartyBtn = new JButton("Add Party");
        JButton reverseCellBtn = new JButton("Reverse Cell");
        JButton addElectionBtn = new JButton("Add Election");
        JButton addConstituencyBtn = new JButton("Add Constituency");

        // Add buttons to panel
        bottomPanel.add(removeColumnBtn);
        bottomPanel.add(removeRowBtn);
        bottomPanel.add(addPartyBtn);
        bottomPanel.add(reverseCellBtn);
        bottomPanel.add(addElectionBtn);
        bottomPanel.add(addConstituencyBtn);

        add(bottomPanel, BorderLayout.SOUTH);

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

        // Extract data rows
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

        // Create a model for the Excel table with checkbox in first column
        DefaultTableModel excelModel = new DefaultTableModel(data, columnNames) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return column == 0; // Only checkbox column editable
            }

            @Override
            public Class<?> getColumnClass(int columnIndex) {
                if (columnIndex == 0) {
                    return Boolean.class;
                }
                return super.getColumnClass(columnIndex);
            }
        };

        excelTable.setModel(excelModel);
    }
}

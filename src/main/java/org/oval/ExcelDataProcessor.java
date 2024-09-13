package org.oval;

import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.io.*;
import java.util.*;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class ExcelDataProcessor extends JFrame {

    // Data structures to hold the data
    private List<List<String>> data = new ArrayList<>();
    private List<List<String>> processedData = new ArrayList<>();

    // Constructor to set up the GUI
    public ExcelDataProcessor() {
        setTitle("Excel Data Processor");
        setSize(400, 200);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new FlowLayout());

        JButton loadButton = new JButton("Load Excel File");
        JButton processButton = new JButton("Process Data");
        JButton saveButton = new JButton("Save Processed File");

        add(loadButton);
        add(processButton);
        add(saveButton);

        // Add action listeners
        loadButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                loadFile();
            }
        });

        processButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                processData();
            }
        });

        saveButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                saveFile();
            }
        });

        setLocationRelativeTo(null);
        setVisible(true);
    }

    private void loadFile() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Select Excel File");
        int userSelection = fileChooser.showOpenDialog(this);
        if (userSelection == JFileChooser.APPROVE_OPTION) {
            File fileToLoad = fileChooser.getSelectedFile();
            try {
                FileInputStream fis = new FileInputStream(fileToLoad);
                Workbook workbook = new XSSFWorkbook(fis);
                Sheet sheet = workbook.getSheetAt(0);
                data.clear(); // Clear previous data

                // Iterate through rows and columns
                for (Row row : sheet) {
                    List<String> rowData = new ArrayList<>();
                    for (Cell cell : row) {
                        cell.setCellType(CellType.STRING);
                        rowData.add(cell.getStringCellValue());
                    }
                    data.add(rowData);
                }
                workbook.close();
                fis.close();
                JOptionPane.showMessageDialog(this, "File loaded successfully!");
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(this, "Failed to load the file: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private void processData() {
        if (data.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Please load a file first.", "Warning", JOptionPane.WARNING_MESSAGE);
            return;
        }
        try {
            processedData.clear();
            for (List<String> row : data) {
                List<String> newRow = new ArrayList<>(row);
                if (!row.isEmpty()) {
                    String firstCell = row.get(0);
                    String[] splitValues = firstCell.split("\\.");
                    newRow.addAll(Arrays.asList(splitValues));
                }
                processedData.add(newRow);
            }
            JOptionPane.showMessageDialog(this, "Data processed successfully!");
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Failed to process the data: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void saveFile() {
        if (processedData.isEmpty()) {
            JOptionPane.showMessageDialog(this, "No data to save. Please process the data first.", "Warning", JOptionPane.WARNING_MESSAGE);
            return;
        }

        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Save Processed File");
        int userSelection = fileChooser.showSaveDialog(this);

        if (userSelection == JFileChooser.APPROVE_OPTION) {
            File fileToSave = fileChooser.getSelectedFile();
            try {
                // Save the processed data
                Workbook workbook = new XSSFWorkbook();
                Sheet sheet = workbook.createSheet("Processed Data");

                for (int i = 0; i < processedData.size(); i++) {
                    Row row = sheet.createRow(i);
                    List<String> rowData = processedData.get(i);
                    for (int j = 0; j < rowData.size(); j++) {
                        Cell cell = row.createCell(j);
                        cell.setCellValue(rowData.get(j));
                    }
                }

                FileOutputStream fos = new FileOutputStream(fileToSave);
                workbook.write(fos);
                fos.close();
                workbook.close();

                JOptionPane.showMessageDialog(this, "File saved successfully!");

                // Read back the saved file
                FileInputStream fis = new FileInputStream(fileToSave);
                Workbook newWorkbook = new XSSFWorkbook(fis);
                Sheet newSheet = newWorkbook.getSheetAt(0);

                // Prepare to count occurrences starting from column B (index 1)
                Map<String, Integer> elementCounts = new HashMap<>();

                for (Row row : newSheet) {
                    for (int i = 1; i < row.getLastCellNum(); i++) {
                        Cell cell = row.getCell(i);
                        if (cell != null) {
                            cell.setCellType(CellType.STRING);
                            String value = cell.getStringCellValue();
                            elementCounts.put(value, elementCounts.getOrDefault(value, 0) + 1);
                        }
                    }
                }

                fis.close();
                newWorkbook.close();

                // Display the counts
                StringBuilder countsMessage = new StringBuilder("Element counts:\n");
                for (Map.Entry<String, Integer> entry : elementCounts.entrySet()) {
                    countsMessage.append(entry.getKey()).append(": ").append(entry.getValue()).append("\n");
                }
                System.out.println(countsMessage.toString());

                // Optionally save the counts to a new Excel file
                userSelection = fileChooser.showSaveDialog(this);
                if (userSelection == JFileChooser.APPROVE_OPTION) {
                    File countsFile = fileChooser.getSelectedFile();
                    Workbook countsWorkbook = new XSSFWorkbook();
                    Sheet countsSheet = countsWorkbook.createSheet("Element Counts");

                    int rowIndex = 0;
                    for (Map.Entry<String, Integer> entry : elementCounts.entrySet()) {
                        Row row = countsSheet.createRow(rowIndex++);
                        Cell cellKey = row.createCell(0);
                        Cell cellValue = row.createCell(1);
                        cellKey.setCellValue(entry.getKey());
                        cellValue.setCellValue(entry.getValue());
                    }

                    FileOutputStream countsFos = new FileOutputStream(countsFile);
                    countsWorkbook.write(countsFos);
                    countsFos.close();
                    countsWorkbook.close();
                }

            } catch (Exception ex) {
                JOptionPane.showMessageDialog(this, "Failed to save the file: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    public static void main(String[] args) {
        // Set the look and feel to system default
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        new ExcelDataProcessor();
    }
}

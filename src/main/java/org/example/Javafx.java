package org.example;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.RandomStringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class Javafx extends Application {

    private int rowCount = 0;
    private Workbook workbook;
    private Sheet sheet;
    private String fileName = "output.xlsx";
    private String excelFilePath = System.getProperty("user.home") + File.separator +
             "Documents" + File.separator + "WORK_FILES" + File.separator;

    @Override
    public void start(Stage primaryStage) {
        // Initialize Excel file
        initializeExcel();

        // Create the input form
        GridPane grid = new GridPane();
        grid.setPadding(new Insets(10, 10, 10, 10));
        grid.setVgap(10);
        grid.setHgap(10);

        Label filePathLabel = new Label("File saved Location: " +excelFilePath + fileName);

        // Create labels and input fields
        Label label1 = new Label("Enter Date Of Birth:");
        TextField input1 = new TextField();
        Label label2 = new Label("Enter Mother Name:");
        TextField input2 = new TextField();
        Label label3 = new Label("Enter Father Name");
        TextField input3 = new TextField();

        Button saveButton = new Button("Save to Excel");

        // Add components to the grid
        grid.add(label1, 0, 0);
        grid.add(input1, 1, 0);
        grid.add(label2, 0, 1);
        grid.add(input2, 1, 1);
        grid.add(label3, 0, 2);
        grid.add(input3, 1, 2);
        grid.add(saveButton, 1, 3);

        // VBox to hold the filePathLabel with padding
        VBox vbox = new VBox(filePathLabel);
        vbox.setPadding(new Insets(10, 0, 0, 0));

        grid.add(filePathLabel, 1, 4);


        // Set button action to save data progressively
        saveButton.setOnAction(e -> {
            String val1 = input1.getText();
            String val2 = input2.getText();
            String val3 = input3.getText();

            // Add the data to Excel
            saveToExcel(val1, val2, val3);

            // Clear fields after saving
            input1.clear();
            input2.clear();
            input3.clear();
        });

        Scene scene = new Scene(grid, 400, 200);
        primaryStage.setTitle("Excel Saver");
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    private void initializeExcel() {
        System.out.println("Excel initialised");

        try {
            File file = new File(excelFilePath + fileName);
            if(file.exists()) {
                fileName = RandomStringUtils.randomAlphabetic(5) + ".xlsx";
                file = new File(excelFilePath + fileName);
                FileUtils.forceMkdirParent(file);
            } else {
                FileUtils.forceMkdirParent(file);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Create a new workbook and a sheet
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Data");

        // Create header row
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("DOB");
        header.createCell(1).setCellValue("Mother Name");
        header.createCell(2).setCellValue("Father Name");

        rowCount++;
    }

    private void saveToExcel(String val1, String val2, String val3) {
        // Create a new row in the sheet
        Row row = sheet.createRow(rowCount++);

        // Add values to the row
        row.createCell(0).setCellValue(val1);
        row.createCell(1).setCellValue(val2);
        row.createCell(2).setCellValue(val3);

        // Write the updated workbook to a file
        try (FileOutputStream fileOut = new FileOutputStream(new File(excelFilePath + fileName))) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public void stop() throws Exception {
        // Close the workbook when the application stops
        if (workbook != null) {
            try (FileOutputStream fileOut = new FileOutputStream(new File(excelFilePath + fileName))) {
                workbook.write(fileOut);
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                workbook.close();
            }
        }
    }

    public static void main(String[] args) {
        launch(args);
    }

}
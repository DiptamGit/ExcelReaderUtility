package com.diptam;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;

public class ExcelReader {

    public static final String FILE_PATH = "C:/Users/dipta/Desktop/Sample.xlsx";

    public static void main(String[] args) {
        try {
            read();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void read() throws Exception{

        //Creating a workbook from an Excel file
        Workbook workbook = WorkbookFactory.create(new File(FILE_PATH));
        System.out.println("Workbook found with number of sheets "+workbook.getNumberOfSheets());

        //This will give you a list of Sheet objects which you can iterate
        workbook.sheetIterator().forEachRemaining(sheet -> {
            System.out.println(sheet.getSheetName());
        });

        DataFormatter formatter = new DataFormatter(); // use dataformatter to get cell values as String

        //now we will use same lambda function above for sheet iterator and use each sheet object to get row by row values
        //then iterate every row for all cell values
        workbook.sheetIterator().forEachRemaining(sheet -> {
            sheet.forEach(row -> {
                row.forEach(cell -> {
                    String cellValue = formatter.formatCellValue(cell);
                    System.out.println(cellValue + "\t");
                });
            });
        });


        //finally close the workbook
        workbook.close();
    }
}

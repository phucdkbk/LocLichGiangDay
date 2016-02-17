/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.phucdk.lichhoc.main;

import com.phucdk.lichhoc.util.ExcelUtil;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Administrator
 */
public class ExtractInvidualSchedule {

    public static void main(String[] args) throws Exception {
        ExcelUtil.readData("D:\\20150831\\Projects\\LoclichCongtac\\1. General file.xlsx");
    }
    
    private void testReadFile() throws FileNotFoundException, IOException{
        File myFile = new File("D:\\20150831\\Projects\\LoclichCongtac\\TestPOI\\test1.xlsx");
        FileInputStream fis = new FileInputStream(myFile);
// Finds the workbook instance for XLSX file
        XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);
// Return first sheet from the XLSX workbook
        XSSFSheet mySheet = myWorkBook.getSheetAt(0);
// Get iterator to all the rows in current sheet
        Iterator<Row> rowIterator = mySheet.iterator();
// Traversing over each row of XLSX file
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
// For each row, iterate through each columns
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        System.out.print(cell.getStringCellValue() + "\t");
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.print(cell.getNumericCellValue() + "\t");
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        System.out.print(cell.getBooleanCellValue() + "\t");
                        break;
                    default:
                }
            }
            System.out.println("");
        }
        for (int i = 0; i < mySheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = mySheet.getMergedRegion(i);
            // Just add it to the sheet on the new workbook.
            //newSheet.addMergedRegion(mergedRegion);
            System.out.println(mergedRegion.formatAsString());
        }
    }

}

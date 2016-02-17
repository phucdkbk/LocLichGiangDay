/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.phucdk.lichhoc.util;

import com.phucdk.lichhoc.object.GeneralData;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Administrator
 */
public class ExcelUtil {

    public static void readAndRewrite() throws FileNotFoundException, IOException, InvalidFormatException {
        InputStream inp = new FileInputStream("D:\\testPOI\\test1.xlsx");
        //InputStream inp = new FileInputStream("workbook.xlsx");

        Workbook wb = WorkbookFactory.create(inp);
        Sheet sheet = wb.getSheetAt(0);
        Row row = sheet.getRow(2);
        Cell cell = row.getCell(3);
        if (cell == null) {
            cell = row.createCell(3);
        }
        cell.setCellType(Cell.CELL_TYPE_STRING);
        cell.setCellValue("a test");

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("test1.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    public static GeneralData readData(String fileName) throws IOException {
        GeneralData generalData = new GeneralData();
        File myFile = new File(fileName);
        FileInputStream fis = new FileInputStream(myFile);
        // Finds the workbook instance for XLSX file
        XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);
        // Return first sheet from the XLSX workbook
        XSSFSheet mySheet = myWorkBook.getSheetAt(0);
        // Get iterator to all the rows in current sheet
        Iterator<Row> rowIterator = mySheet.iterator();
        // Traversing over each row of XLSX file
        int numberOfRow = mySheet.getPhysicalNumberOfRows();
        System.out.println("numberOfRow: " + numberOfRow);

        List<MergeRegion> listMergedRegion = new ArrayList<>();
        for (int i = 0; i < mySheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = mySheet.getMergedRegion(i);
            // Just add it to the sheet on the new workbook.
            //newSheet.addMergedRegion(mergedRegion);
            listMergedRegion.add(new MergeRegion(mergedRegion));
            System.out.println(mergedRegion.formatAsString());
        }

        for (int i = 3; i < numberOfRow; i++) {

        }

//        while (rowIterator.hasNext()) {
//            Row row = rowIterator.next();
//            Iterator<Cell> cellIterator = row.cellIterator();
//            while (cellIterator.hasNext()) {
//                Cell cell = cellIterator.next();
//                switch (cell.getCellType()) {
//                    case Cell.CELL_TYPE_STRING:
//                        System.out.print(cell.getStringCellValue() + "\t");
//                        break;
//                    case Cell.CELL_TYPE_NUMERIC:
//                        System.out.print(cell.getNumericCellValue() + "\t");
//                        break;
//                    case Cell.CELL_TYPE_BOOLEAN:
//                        System.out.print(cell.getBooleanCellValue() + "\t");
//                        break;
//                    default:
//                }
//            }
//            System.out.println("");
//        }


        return generalData;
    }

}

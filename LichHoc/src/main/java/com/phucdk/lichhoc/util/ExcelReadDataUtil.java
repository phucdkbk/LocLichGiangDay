/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.phucdk.lichhoc.util;

import com.phucdk.lichhoc.object.GeneralData;
import com.phucdk.lichhoc.object.LectureSchedule;
import com.phucdk.lichhoc.object.SchoolClass;
import com.phucdk.lichhoc.object.Teacher;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Administrator
 */
public class ExcelReadDataUtil {    

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

        int startRow = Constants.ROW.START_ROW;
        Date startDateOfWeek = getCell(1, Constants.COLUMN.MONDAY_COLUMN, mySheet).getDateCellValue();
        generalData.setStartDateOfWeek(startDateOfWeek);
        while (haveNextClass(startRow, mySheet)) {
            ClassRowPair classRowPair = getNextClassRowPair(startRow, mySheet);
            if (classRowPair != null) {
                for (int i = Constants.COLUMN.MONDAY_COLUMN; i <= Constants.COLUMN.SUNDAY_COLUMN; i++) {
                    if (!isEmptyCell(classRowPair.getFromRow(), i, mySheet)) {
                        String teacherName = getStringCellValue(classRowPair.getFromRow(), i, mySheet);
                        Teacher teacher = getTeacher(teacherName, generalData);
                        LectureSchedule lectureSchedule = new LectureSchedule();
                        lectureSchedule.setTeacher(teacher);
                        String campus = getStringCellValue(classRowPair.getFromRow(), Constants.COLUMN.CAMPUS_COLUMN, mySheet);
                        String schoolClassName = getStringCellValue(classRowPair.getFromRow(), Constants.COLUMN.SCHOOL_CLASS_COLUMN, mySheet);
                        String time = getStringCellValue(classRowPair.getFromRow(), Constants.COLUMN.TIME_COLUMN, mySheet);
                        SchoolClass schoolClass = new SchoolClass(schoolClassName);
                        Date date = getDateOfColumn(startDateOfWeek, i);
                        lectureSchedule.setCampus(campus);
                        lectureSchedule.setSchoolClass(schoolClass);
                        lectureSchedule.setDate(date);
                        lectureSchedule.setHour(time);
                        String lession = null;
                        for (int j = classRowPair.getFromRow() + 1; j <= classRowPair.getToRow(); j++) {
                            lession = getStringCellValue(j, i, mySheet);
                            if (lession != null && !"".equals(lession.trim())) {
                                break;
                            }
                        }
                        lectureSchedule.setLession(lession);
                        generalData.getListLectureSchedules().add(lectureSchedule);
                    }
                }
                startRow = classRowPair.getToRow();
            }

        }
        return generalData;
    }

    private static Teacher getTeacher(String teacherName, GeneralData generalData) {
        Teacher teacher = null;
        if (!existTeacher(teacherName, generalData)) {
            teacher = new Teacher();
            teacher.setFullName(teacherName);
            generalData.getListTeachers().add(teacher);
        } else {
            teacher = getTeacherByName(teacherName, generalData);
        }
        return teacher;
    }

    private static boolean haveNextClass(int startRow, XSSFSheet sheet) {
        ClassRowPair classRowPair = getNextClassRowPair(startRow, sheet);
        return classRowPair != null;
    }

    private static ClassRowPair getNextClassRowPair(int fromRow, XSSFSheet sheet) {
        int numberOfRow = sheet.getPhysicalNumberOfRows();
        if (fromRow < numberOfRow) {
            ClassRowPair classRowPair = new ClassRowPair();
            classRowPair.setFromRow(fromRow + 1);
            int startRow = fromRow + 1;
            XSSFRow row;
            XSSFCell cell;
            String cellValue = null;
            System.out.println(fromRow);
            while (StringUtils.isEmpty(cellValue) && startRow < numberOfRow) {
                startRow++;
                row = sheet.getRow(startRow);
                if (row != null) {
                    cell = row.getCell(Constants.COLUMN.TIME_COLUMN);
                    cellValue = cell.getStringCellValue();
                } else {
                    return null;
                }
            }
            classRowPair.setToRow(startRow - 1);
            return classRowPair;
        } else {
            return null;
        }
    }

    private static Cell getCell(int row, int column, XSSFSheet sheet) {
        return CellUtil.getRow(row, sheet).getCell(column);
    }

    private static String getStringCellValue(int row, int column, XSSFSheet sheet) {
        Cell cell = getCell(row, column, sheet);
        if (cell != null) {
            return cell.getStringCellValue();
        } else {
            return null;
        }
    }

    private static boolean isEmptyCell(int row, int column, XSSFSheet sheet) {
        String cellValue = getStringCellValue(row, column, sheet);
        boolean isEmpty = true;
        if (cellValue != null && !"".equals(cellValue.trim())) {
            isEmpty = false;
        }
        return isEmpty;
    }

    private static boolean existTeacher(String teacherName, GeneralData generalData) {
        Teacher teacher = getTeacherByName(teacherName, generalData);
        return teacher != null;
    }

    private static Teacher getTeacherByName(String teacherName, GeneralData generalData) {
        for (int i = 0; i < generalData.getListTeachers().size(); i++) {
            if (teacherName.equals(generalData.getListTeachers().get(i).getFullName())) {
                return generalData.getListTeachers().get(i);
            }
        }
        return null;
    }

    private static Date getDateOfColumn(Date startDateOfWeek, int dateColumn) {
        Calendar cal = Calendar.getInstance();
        cal.setTime(startDateOfWeek);
        cal.add(Calendar.DATE, dateColumn - Constants.COLUMN.MONDAY_COLUMN);
        return cal.getTime();
    }

}

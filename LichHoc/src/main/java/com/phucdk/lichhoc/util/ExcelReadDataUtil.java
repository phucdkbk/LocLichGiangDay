/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.phucdk.lichhoc.util;

import com.phucdk.lichhoc.object.BusySchedule;
import com.phucdk.lichhoc.object.GeneralData;
import com.phucdk.lichhoc.object.LectureSchedule;
import com.phucdk.lichhoc.object.SchoolClass;
import com.phucdk.lichhoc.object.Teacher;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Administrator
 */
public class ExcelReadDataUtil {

    public static GeneralData readData(String fileName, String busySchedule) throws IOException {
        GeneralData generalData = new GeneralData();
        getScheduleData(fileName, generalData);
        getBusyScheduleData(busySchedule, generalData);
        return generalData;
    }

    private static void getScheduleData(String fileName, GeneralData generalData) throws FileNotFoundException, IOException {
        File myFile = new File(fileName);
        FileInputStream fis = new FileInputStream(myFile);
        // Finds the workbook instance for XLSX file
        XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);
        // Return first sheet from the XLSX workbook
        XSSFSheet mySheet = myWorkBook.getSheetAt(0);
        int startRow = Constants.ROW.START_ROW;
        Date startDateOfWeek = getCell(1, Constants.COLUMN.MONDAY_COLUMN, mySheet).getDateCellValue();
        generalData.setStartDateOfWeek(startDateOfWeek);
        while (haveNextClass(startRow, mySheet)) {
            ClassRowPair classRowPair = getNextClassRowPair(startRow, mySheet);
            if (classRowPair != null) {
                for (int i = Constants.COLUMN.MONDAY_COLUMN; i <= Constants.COLUMN.SUNDAY_COLUMN; i++) {
                    if (!isEmptyCell(classRowPair.getFromRow(), i, mySheet)) {
                        String teacherName = getStringCellValue(classRowPair.getFromRow(), i, mySheet);
                        if ("Carmel".equals(teacherName)) {
                        } else {
                            //continue;
                        }
                        Teacher teacher = getTeacher(teacherName, generalData.getListTeachers());
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
        markConfictLectureSchedule(generalData);
    }

    private static void getBusyScheduleData(String fileName, GeneralData generalData) throws FileNotFoundException, IOException {
        File myFile = new File(fileName);
        FileInputStream fis = new FileInputStream(myFile);
        // Finds the workbook instance for XLSX file
        XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);

        for (int i = 0; i < myWorkBook.getNumberOfSheets(); i++) {
            //for (int i = 0; i < 1; i++) {
            XSSFSheet mySheet = myWorkBook.getSheetAt(i);
            String teacherName = mySheet.getSheetName().trim();
            if ("Carmel".equals(teacherName)) {
            } else {
                //continue;
            }
            Teacher teacher = getTeacher(teacherName, generalData.getListBusyTeachers());
            List<String> listTimes = new ArrayList<>();
            int numberOfRows = mySheet.getPhysicalNumberOfRows();
            Date startDateOfWeek = getCell(5, Constants.BUSY_SCHEDULE.COLUMN.MONDAY_COLUMN, mySheet).getDateCellValue();
            int startRow = Constants.BUSY_SCHEDULE.ROW.START_ROW_FRANCE;
            if (startDateOfWeek == null) {
                startDateOfWeek = getCell(4, Constants.BUSY_SCHEDULE.COLUMN.MONDAY_COLUMN, mySheet).getDateCellValue();
                startRow = Constants.BUSY_SCHEDULE.ROW.START_ROW_VI;
            }
            for (int j = startRow; j <= numberOfRows; j++) {
                XSSFRow row;
                row = mySheet.getRow(j);
                if (row != null) {
                    XSSFCell row_0 = row.getCell(0);
                    if (row_0 != null && !StringUtils.isEmpty(row_0.getStringCellValue())) {
                        for (int k = Constants.BUSY_SCHEDULE.COLUMN.MONDAY_COLUMN; k <= Constants.BUSY_SCHEDULE.COLUMN.SUNDAY_COLUMN; k++) {
                            XSSFCell cell_dayOfWeek = row.getCell(k);
                            if (cell_dayOfWeek != null) {
                                XSSFCellStyle cellStyle = cell_dayOfWeek.getCellStyle();
//                                System.out.println(cellStyle.getFillBackgroundColor());
//                                System.out.println(cellStyle.getFillPattern());
                                if (cellStyle.getFillPattern() != (int) HSSFCellStyle.NO_FILL) {
                                    BusySchedule busySchedule = new BusySchedule();
                                    busySchedule.setTeacher(teacher);
                                    busySchedule.setHour(row_0.getStringCellValue().trim());
                                    busySchedule.setDate(getDateOfColumnBusy(startDateOfWeek, k));
                                    generalData.getListBusySchedules().add(busySchedule);
                                }
                            }
                        }
                        if(!listTimes.contains(row_0.getStringCellValue().trim())){
                            listTimes.add(row_0.getStringCellValue().trim());
                        }                        
                    }
                }
            }
            generalData.getMapTeacherTimes().put(teacher, listTimes);
        }
    }

    private static Teacher getTeacher(String teacherName, List<Teacher> listTeachers) {
        Teacher teacher = null;
        if (!existTeacher(teacherName, listTeachers)) {
            teacher = new Teacher();
            teacher.setFullName(teacherName);
            listTeachers.add(teacher);
        } else {
            teacher = getTeacherByName(teacherName, listTeachers);
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
            //System.out.println(fromRow);
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

    public static String getStringCellValue(int row, int column, XSSFSheet sheet) {
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

    private static boolean existTeacher(String teacherName, List<Teacher> listTeachers) {
        Teacher teacher = getTeacherByName(teacherName, listTeachers);
        return teacher != null;
    }

    private static Teacher getTeacherByName(String teacherName, List<Teacher> listTeachers) {
        for (int i = 0; i < listTeachers.size(); i++) {
            if (teacherName.equals(listTeachers.get(i).getFullName())) {
                return listTeachers.get(i);
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

    private static Date getDateOfColumnBusy(Date startDateOfWeek, int dateColumn) {
        Calendar cal = Calendar.getInstance();
        cal.setTime(startDateOfWeek);
        cal.add(Calendar.DATE, dateColumn - Constants.BUSY_SCHEDULE.COLUMN.MONDAY_COLUMN);
        return cal.getTime();
    }

    private static void markConfictLectureSchedule(GeneralData generalData) {
        for (int i = 0; i < generalData.getListLectureSchedules().size() - 1; i++) {
            for (int j = i + 1; j < generalData.getListLectureSchedules().size(); j++) {
                if (confictLectureSchedule(generalData.getListLectureSchedules().get(i), generalData.getListLectureSchedules().get(j))) {
                    generalData.getListLectureSchedules().get(i).setIsConfict(true);
                    generalData.getListLectureSchedules().get(j).setIsConfict(true);
                }
            }
        }
    }

    private static boolean confictLectureSchedule(LectureSchedule lectureSchedule1, LectureSchedule lectureSchedule2) {
        if (!lectureSchedule1.getTeacher().equals(lectureSchedule2.getTeacher())) {
            return false;
        }
        if (!DateTimeUtils.equalDate(lectureSchedule1.getDate(), lectureSchedule2.getDate())) {
            return false;
        }
        if (!lectureSchedule1.getHour().equals(lectureSchedule2.getHour())) {
            return false;
        }
        return true;
    }

}

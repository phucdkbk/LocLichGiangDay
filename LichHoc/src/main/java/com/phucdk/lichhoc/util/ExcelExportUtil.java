/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.phucdk.lichhoc.util;

import com.phucdk.lichhoc.object.GeneralData;
import com.phucdk.lichhoc.object.LectureSchedule;
import com.phucdk.lichhoc.object.Teacher;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Administrator
 */
public class ExcelExportUtil {

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

    public static void exportFile(GeneralData generalData) throws FileNotFoundException, IOException, Exception {
        //for (int i = 0; i < generalData.getListTeachers().size(); i++) {
        for (int i = 0; i < 1; i++) {
            Teacher teacher = generalData.getListTeachers().get(i);

            Workbook wb = new XSSFWorkbook();
            CreationHelper createHelper = wb.getCreationHelper();
            Sheet sheet = wb.createSheet(teacher.getFullName());

            //-----------------------  row 0 --------------------
            sheet.createRow((short) 0);
            //-----------------------  row 1 --------------------
            Row row1 = sheet.createRow((short) 1);
            Cell cell_10 = row1.createCell(0);
            cell_10.setCellValue("Teacher");

            Cell cell_11 = row1.createCell(1);
            cell_11.setCellValue(teacher.getFullName());
            //-----------------------  row 2 --------------------
            Row row2 = sheet.createRow((short) 2);
            row2.createCell(0).setCellValue("Tel:");
            row2.createCell(3).setCellValue("Email");
            //-----------------------  row 3 --------------------
            Row row3 = sheet.createRow((short) 3);
            row3.createCell(0).setCellValue(getWeekTitle(generalData));
            //-----------------------  row 4 --------------------
            Row row4 = sheet.createRow((short) 4);
            Cell cell_40 = row4.createCell(0);
            cell_40.setCellValue("Time");
            Cell cell_41 = row4.createCell(1);
            cell_41.setCellValue("Mon");
            Cell cell_42 = row4.createCell(2);
            cell_42.setCellValue("Tue");
            Cell cell_43 = row4.createCell(3);
            cell_43.setCellValue("Wed");
            Cell cell_44 = row4.createCell(4);
            cell_44.setCellValue("Thu");
            Cell cell_45 = row4.createCell(5);
            cell_45.setCellValue("Fri");
            Cell cell_46 = row4.createCell(6);
            cell_46.setCellValue("Sat");
            Cell cell_47 = row4.createCell(7);
            cell_47.setCellValue("Sun");
            //-----------------------  row 5 --------------------
            Row row5 = sheet.createRow((short) 5);
            Cell cell_51 = row5.createCell(1);
            cell_51.setCellValue("Mon");
            Cell cell_52 = row5.createCell(2);
            cell_52.setCellValue("Tue");
            Cell cell_53 = row5.createCell(3);
            cell_53.setCellValue("Wed");
            Cell cell_54 = row5.createCell(4);
            cell_54.setCellValue("Thu");
            Cell cell_55 = row5.createCell(5);
            cell_55.setCellValue("Fri");
            Cell cell_56 = row5.createCell(6);
            cell_56.setCellValue("Sat");
            Cell cell_57 = row5.createCell(7);
            cell_57.setCellValue("Sun");
            //---------------------------------------------------
            List<LectureSchedule> listLectureSchedules = getListLectureScheduleByTeacher(generalData, teacher);
            List<String> listTimes = getListTimes(listLectureSchedules);
            for (int j = 0; j < listTimes.size(); j++) {
                String time = listTimes.get(j);
                Row loopRow_0 = sheet.createRow((short) (6 + j * 4));
                Row loopRow_1 = sheet.createRow((short) (6 + j * 4 + 1));
                Row loopRow_2 = sheet.createRow((short) (6 + j * 4 + 2));
                Row loopRow_3 = sheet.createRow((short) (6 + j * 4 + 3));
                Cell loopRow_01 = loopRow_0.createCell(0);
                loopRow_01.setCellValue(time);

                List<LectureSchedule> listLectureSchedulesByTime = getListLectureScheduleByTime(listLectureSchedules, time);
                for (int k = 0; k < listLectureSchedulesByTime.size(); k++) {
                    LectureSchedule lectureSchedule = listLectureSchedulesByTime.get(k);
                    switch (getDayOfWeek(lectureSchedule)) {
                        case Calendar.MONDAY:
                            setLectureValue(loopRow_0, lectureSchedule, loopRow_1, loopRow_3, 1);
                            break;
                        case Calendar.TUESDAY:
                            setLectureValue(loopRow_0, lectureSchedule, loopRow_1, loopRow_3, 2);
                            break;
                        case Calendar.WEDNESDAY:
                            setLectureValue(loopRow_0, lectureSchedule, loopRow_1, loopRow_3, 3);
                            break;
                        case Calendar.THURSDAY:
                            setLectureValue(loopRow_0, lectureSchedule, loopRow_1, loopRow_3, 4);
                            break;
                        case Calendar.FRIDAY:
                            setLectureValue(loopRow_0, lectureSchedule, loopRow_1, loopRow_3, 5);
                            break;
                        case Calendar.SATURDAY:
                            setLectureValue(loopRow_0, lectureSchedule, loopRow_1, loopRow_3, 6);
                            break;
                        case Calendar.SUNDAY:
                            setLectureValue(loopRow_0, lectureSchedule, loopRow_1, loopRow_3, 7);
                            break;
                    }
                }
            }

//            Cell cell3 = row.createCell(3);
//
//            CellStyle cellStyle = wb.createCellStyle();
//            cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/MM/yyyy"));
//            cell3.setCellStyle(cellStyle);
//            cell3.setCellValue(new Date());
            // Write the output to a file
            FileOutputStream fileOut = new FileOutputStream("D:\\20150831\\Projects\\LoclichCongtac\\output\\test.xlsx");
            wb.write(fileOut);
            fileOut.close();
        }
    }

    private static void setLectureValue(Row loopRow_0, LectureSchedule lectureSchedule, Row loopRow_1, Row loopRow_3, int column) {
        Cell loopRow_0_monday = loopRow_0.createCell(column);
        loopRow_0_monday.setCellValue(lectureSchedule.getSchoolClass().getSchoolClassName());
        Cell loopRow_1_monday = loopRow_1.createCell(column);
        loopRow_1_monday.setCellValue(lectureSchedule.getLession());
        Cell loopRow_3_monday = loopRow_3.createCell(column);
        loopRow_3_monday.setCellValue(lectureSchedule.getCampus());
    }

    private static String getWeekTitle(GeneralData generalData) throws Exception {
        StringBuilder weekTitle = new StringBuilder();
        weekTitle.append("Week from ");
        weekTitle.append(DateTimeUtils.convertDateToString(generalData.getStartDateOfWeek(), "dd-MM"));
        weekTitle.append(" to ");
        Date getSundayOfWeek = DateTimeUtils.getSundayOfWeek(generalData.getStartDateOfWeek());
        weekTitle.append(DateTimeUtils.convertDateToString(getSundayOfWeek, "dd-MM"));
        return weekTitle.toString();
    }

    private static List<LectureSchedule> getListLectureScheduleByTeacher(GeneralData generalData, Teacher teacher) {
        List<LectureSchedule> listLectureSchedules = new ArrayList<>();
        for(int i = 0; i<generalData.getListLectureSchedules().size(); i++){
            LectureSchedule lectureSchedule = generalData.getListLectureSchedules().get(i);
            if(lectureSchedule.getTeacher().equals(teacher)){
                listLectureSchedules.add(lectureSchedule);
            }
        }
        return listLectureSchedules;
    }

    private static List<String> getListTimes(List<LectureSchedule> listLectureSchedules) {
        List<String> listTimes = new ArrayList<>();
        for (LectureSchedule lectureSchedule : listLectureSchedules) {
            if (lectureSchedule.getHour() != null) {
                if (!listTimes.contains(lectureSchedule.getHour().trim())) {
                    listTimes.add(lectureSchedule.getHour().trim());
                }
            }
        }
        return listTimes;
    }

    private static List<LectureSchedule> getListLectureScheduleByTime(List<LectureSchedule> listLectureSchedules, String time) {
        List<LectureSchedule> listLectureSchedulesByTime = new ArrayList<>();
        for (LectureSchedule lectureSchedule : listLectureSchedules) {
            if (lectureSchedule.getHour() != null && lectureSchedule.getHour().equals(time)) {
                listLectureSchedulesByTime.add(lectureSchedule);
            }
        }
        return listLectureSchedulesByTime;
    }

    private static int getDayOfWeek(LectureSchedule lectureSchedule) {
        return DateTimeUtils.getDayOfWeek(lectureSchedule.getDate());
    }

}

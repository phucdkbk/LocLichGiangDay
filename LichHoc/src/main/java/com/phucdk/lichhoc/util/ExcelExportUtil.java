/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.phucdk.lichhoc.util;

import com.phucdk.lichhoc.object.BusySchedule;
import com.phucdk.lichhoc.object.GeneralData;
import com.phucdk.lichhoc.object.LectureSchedule;
import com.phucdk.lichhoc.object.Teacher;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Administrator
 */
public class ExcelExportUtil {

    public static void exportFile(GeneralData generalData, String outputFolder) throws FileNotFoundException, IOException, Exception {
        outputFolder = outputFolder + "\\" + DateTimeUtils.convertDateToString(new Date(), "yyyyMMdd_HHmmss");
        for (int i = 0; i < generalData.getListTeachers().size(); i++) {
            //for (int i = 0; i < 1; i++) {
            Teacher teacher = generalData.getListTeachers().get(i);

            if (teacher.getFullName().toLowerCase().equals("carmel")) {
                int debug = 1;
            } else {
                //continue;
            }

            XSSFWorkbook wb = new XSSFWorkbook();
            CreationHelper createHelper = wb.getCreationHelper();
            XSSFSheet sheet = wb.createSheet(teacher.getFullName());
            sheet.setDisplayGridlines(false);
            sheet.setColumnWidth(0, 1500);
            sheet.setColumnWidth(1, 4200);
            sheet.setColumnWidth(2, 4200);
            sheet.setColumnWidth(3, 4200);
            sheet.setColumnWidth(4, 4200);
            sheet.setColumnWidth(5, 4200);
            sheet.setColumnWidth(6, 4200);
            sheet.setColumnWidth(7, 4200);

//            for (int j = Constants.COLUMN.INVIDUAL.SUNDAY_COLUMN + 2; j < 16384; j++) {
//                sheet.setColumnHidden(j, true);
//            }
            //-----------------------  row 0 --------------------
            sheet.createRow((short) 0);
            //-----------------------  row 1 --------------------
            XSSFCellStyle cellStyleBold = wb.createCellStyle();
            Font font = wb.createFont();//Create font
            font.setBoldweight(Font.BOLDWEIGHT_BOLD);//Make font bold
            cellStyleBold.setFont(font);//set it to bold

            XSSFRow row1 = sheet.createRow((short) 1);
            XSSFCell cell_10 = row1.createCell(0);
            cell_10.setCellStyle(cellStyleBold);
            cell_10.setCellValue("Teacher:");

            XSSFCell cell_11 = row1.createCell(1);
            cell_11.setCellStyle(cellStyleBold);
            cell_11.setCellValue(teacher.getFullName());
            //-----------------------  row 2 --------------------
            XSSFRow row2 = sheet.createRow((short) 2);
            XSSFCell cell_20 = row2.createCell(0);
            cell_20.setCellStyle(cellStyleBold);
            cell_20.setCellValue("Tel:");
            XSSFCell cell_23 = row2.createCell(3);
            cell_23.setCellStyle(cellStyleBold);
            cell_23.setCellValue("Email");
            //-----------------------  row 3 --------------------
            XSSFCellStyle cellStyleTitle = wb.createCellStyle();
            XSSFFont fontTitle = wb.createFont();//Create font
            fontTitle.setBoldweight(Font.BOLDWEIGHT_BOLD);//Make font bold
            fontTitle.setFontHeightInPoints((short) 14);
            fontTitle.setUnderline(FontUnderline.SINGLE);
            cellStyleTitle.setFont(fontTitle);//set it to bold
            XSSFRow row3 = sheet.createRow((short) 3);
            XSSFCell cell_30 = row3.createCell(0);
            cell_30.setCellStyle(cellStyleTitle);
            cell_30.setCellValue(getWeekTitle(generalData));
            //-----------------------  row 4 --------------------
            XSSFCellStyle cellStyleDateLabel = wb.createCellStyle();
            cellStyleDateLabel.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            cellStyleDateLabel.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            cellStyleDateLabel.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            cellStyleDateLabel.setBorderTop(HSSFCellStyle.BORDER_THIN);
            cellStyleDateLabel.setBorderRight(HSSFCellStyle.BORDER_THIN);
            cellStyleDateLabel.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            cellStyleDateLabel.setAlignment(HorizontalAlignment.CENTER);

            XSSFCellStyle cellStyleCell_40 = wb.createCellStyle();
            cellStyleCell_40.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            cellStyleCell_40.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            cellStyleCell_40.setBorderTop(HSSFCellStyle.BORDER_THIN);
            cellStyleCell_40.setBorderRight(HSSFCellStyle.BORDER_THIN);
            cellStyleCell_40.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            cellStyleCell_40.setAlignment(HorizontalAlignment.CENTER);

            XSSFRow row4 = sheet.createRow((short) 4);
            XSSFCell cell_40 = row4.createCell(0);
            cell_40.setCellStyle(cellStyleCell_40);
            cell_40.setCellValue("Time");
            XSSFCell cell_41 = row4.createCell(1);
            cell_41.setCellStyle(cellStyleDateLabel);
            cell_41.setCellValue("Mon");
            XSSFCell cell_42 = row4.createCell(2);
            cell_42.setCellStyle(cellStyleDateLabel);
            cell_42.setCellValue("Tue");
            XSSFCell cell_43 = row4.createCell(3);
            cell_43.setCellStyle(cellStyleDateLabel);
            cell_43.setCellValue("Wed");
            XSSFCell cell_44 = row4.createCell(4);
            cell_44.setCellStyle(cellStyleDateLabel);
            cell_44.setCellValue("Thu");
            XSSFCell cell_45 = row4.createCell(5);
            cell_45.setCellStyle(cellStyleDateLabel);
            cell_45.setCellValue("Fri");
            XSSFCell cell_46 = row4.createCell(6);
            cell_46.setCellStyle(cellStyleDateLabel);
            cell_46.setCellValue("Sat");
            XSSFCell cell_47 = row4.createCell(7);
            cell_47.setCellStyle(cellStyleDateLabel);
            cell_47.setCellValue("Sun");
            //-----------------------  row 5 --------------------
            XSSFCellStyle cellStyleDate = wb.createCellStyle();
            cellStyleDate.setDataFormat(createHelper.createDataFormat().getFormat("dd-MMM"));
            cellStyleDate.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            cellStyleDate.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            cellStyleDate.setAlignment(HorizontalAlignment.CENTER);
            cellStyleDate.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            cellStyleDate.setBorderTop(HSSFCellStyle.BORDER_THIN);
            cellStyleDate.setBorderRight(HSSFCellStyle.BORDER_THIN);
            cellStyleDate.setBorderLeft(HSSFCellStyle.BORDER_THIN);

            XSSFCellStyle cellStyleCell_50 = wb.createCellStyle();
            cellStyleCell_50.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            cellStyleCell_50.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            cellStyleCell_50.setAlignment(HorizontalAlignment.CENTER);
            //cellStyleCell_50.setBorderTop(HSSFCellStyle.BORDER_THIN);
            cellStyleCell_50.setBorderRight(HSSFCellStyle.BORDER_THIN);
            cellStyleCell_50.setBorderLeft(HSSFCellStyle.BORDER_THIN);

            XSSFRow row5 = sheet.createRow((short) 5);
            XSSFCell cell_50 = row5.createCell(0);
            cell_50.setCellStyle(cellStyleCell_50);
            XSSFCell cell_51 = row5.createCell(1);
            cell_51.setCellStyle(cellStyleDate);
            cell_51.setCellValue(generalData.getStartDateOfWeek());
            XSSFCell cell_52 = row5.createCell(2);
            cell_52.setCellStyle(cellStyleDate);
            cell_52.setCellValue(DateTimeUtils.addDate(generalData.getStartDateOfWeek(), 1));
            XSSFCell cell_53 = row5.createCell(3);
            cell_53.setCellStyle(cellStyleDate);
            cell_53.setCellValue(DateTimeUtils.addDate(generalData.getStartDateOfWeek(), 2));
            XSSFCell cell_54 = row5.createCell(4);
            cell_54.setCellStyle(cellStyleDate);
            cell_54.setCellValue(DateTimeUtils.addDate(generalData.getStartDateOfWeek(), 3));
            XSSFCell cell_55 = row5.createCell(5);
            cell_55.setCellStyle(cellStyleDate);
            cell_55.setCellValue(DateTimeUtils.addDate(generalData.getStartDateOfWeek(), 4));
            XSSFCell cell_56 = row5.createCell(6);
            cell_56.setCellStyle(cellStyleDate);
            cell_56.setCellValue(DateTimeUtils.addDate(generalData.getStartDateOfWeek(), 5));
            XSSFCell cell_57 = row5.createCell(7);
            cell_57.setCellStyle(cellStyleDate);
            cell_57.setCellValue(DateTimeUtils.addDate(generalData.getStartDateOfWeek(), 6));
            //---------------------------------------------------
            List<LectureSchedule> listLectureSchedules = getListLectureScheduleByTeacher(generalData, teacher);
            List<String> listTimes = getListTimes(listLectureSchedules);
            List<BusySchedule> listBusySchedules = getListBusyScheduleByTeacher(generalData, teacher);

            //--------------- row style for each time----------
            //--------------- style row 0 ---------------------
            Font fontBold = wb.createFont();//Create font
            fontBold.setBoldweight(Font.BOLDWEIGHT_BOLD);//Make font bold

            Font fontCampus = wb.createFont();//Create font
            fontCampus.setBoldweight(Font.BOLDWEIGHT_BOLD);//Make font bold
            fontCampus.setColor(HSSFColor.RED.index);

            XSSFCellStyle cellStyleRow0 = wb.createCellStyle();
            cellStyleRow0.setAlignment(HorizontalAlignment.CENTER);
            cellStyleRow0.setBorderTop(HSSFCellStyle.BORDER_THIN);
            cellStyleRow0.setBorderRight(HSSFCellStyle.BORDER_THIN);
            cellStyleRow0.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            cellStyleRow0.setFont(fontBold);
            cellStyleRow0.setWrapText(true);
            //--------------- style row 1 ---------------
            XSSFCellStyle cellStyleRow1 = wb.createCellStyle();
            cellStyleRow1.setAlignment(HorizontalAlignment.CENTER);
            cellStyleRow1.setBorderRight(HSSFCellStyle.BORDER_THIN);
            cellStyleRow1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            cellStyleRow1.setFont(fontBold);
            cellStyleRow1.setWrapText(true);
            //--------------- style row 3 ---------------
            XSSFCellStyle cellStyleRow3 = wb.createCellStyle();
            cellStyleRow3.setAlignment(HorizontalAlignment.CENTER);
            cellStyleRow3.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            cellStyleRow3.setBorderRight(HSSFCellStyle.BORDER_THIN);
            cellStyleRow3.setBorderLeft(HSSFCellStyle.BORDER_THIN);
            cellStyleRow3.setFont(fontCampus);
            cellStyleRow3.setWrapText(true);

            for (int j = 0; j < listTimes.size(); j++) {
                String time = listTimes.get(j);
                XSSFRow loopRow_0 = sheet.createRow((short) (6 + j * 4));
                XSSFRow loopRow_1 = sheet.createRow((short) (6 + j * 4 + 1));
                XSSFRow loopRow_2 = sheet.createRow((short) (6 + j * 4 + 2));
                XSSFRow loopRow_3 = sheet.createRow((short) (6 + j * 4 + 3));
                XSSFCell loopRow_01 = loopRow_0.createCell(0);
                loopRow_01.setCellValue(time);

                List<LectureSchedule> listLectureSchedulesByTime = getListLectureScheduleByTime(listLectureSchedules, time);
                List<BusySchedule> listBusySchedulesByTime = getListBusyScheduleByTime(listBusySchedules, time);
                for (LectureSchedule lectureSchedule : listLectureSchedulesByTime) {
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

                for (int k = 0; k <= Constants.COLUMN.INVIDUAL.SUNDAY_COLUMN; k++) {
                    XSSFCell loopCell_0 = loopRow_0.getCell(k);
                    if (loopCell_0 == null) {
                        loopCell_0 = loopRow_0.createCell(k);
                    }
                    XSSFCell loopCell_1 = loopRow_1.getCell(k);
                    if (loopCell_1 == null) {
                        loopCell_1 = loopRow_1.createCell(k);
                    }
                    XSSFCell loopCell_2 = loopRow_2.getCell(k);
                    if (loopCell_2 == null) {
                        loopCell_2 = loopRow_2.createCell(k);
                    }
                    XSSFCell loopCell_3 = loopRow_3.getCell(k);
                    if (loopCell_3 == null) {
                        loopCell_3 = loopRow_3.createCell(k);
                    }
                    loopCell_0.setCellStyle(cellStyleRow0);
                    loopCell_1.setCellStyle(cellStyleRow1);
                    loopCell_2.setCellStyle(cellStyleRow1);
                    loopCell_3.setCellStyle(cellStyleRow3);
                }

                for (BusySchedule busySchedule : listBusySchedulesByTime) {
                    switch (DateTimeUtils.getDayOfWeek(busySchedule.getDate())) {
                        case Calendar.MONDAY:
                            fillBusy(loopRow_0, loopRow_1, loopRow_2, loopRow_3, 1);
                            break;
                        case Calendar.TUESDAY:
                            fillBusy(loopRow_0, loopRow_1, loopRow_2, loopRow_3, 2);
                            break;
                        case Calendar.WEDNESDAY:
                            fillBusy(loopRow_0, loopRow_1, loopRow_2, loopRow_3, 3);
                            break;
                        case Calendar.THURSDAY:
                            fillBusy(loopRow_0, loopRow_1, loopRow_2, loopRow_3, 4);
                            break;
                        case Calendar.FRIDAY:
                            fillBusy(loopRow_0, loopRow_1, loopRow_2, loopRow_3, 5);
                            break;
                        case Calendar.SATURDAY:
                            fillBusy(loopRow_0, loopRow_1, loopRow_2, loopRow_3, 6);
                            break;
                        case Calendar.SUNDAY:
                            fillBusy(loopRow_0, loopRow_1, loopRow_2, loopRow_3, 7);
                            break;
                    }
                }
            }

            File folderFile = new File(outputFolder);
            if (!folderFile.exists()) {
                folderFile.mkdirs();
            }
            FileOutputStream fileOut = new FileOutputStream(outputFolder + "\\" + teacher.getFullName() + ".xlsx");
            wb.write(fileOut);
            fileOut.close();
        }
    }

    private static void setLectureValue(XSSFRow loopRow_0, LectureSchedule lectureSchedule, XSSFRow loopRow_1, XSSFRow loopRow_3, int column) {
        XSSFCell loopRow_0_monday = loopRow_0.createCell(column);
        loopRow_0_monday.setCellValue(lectureSchedule.getSchoolClass().getSchoolClassName());
        XSSFCell loopRow_1_monday = loopRow_1.createCell(column);
        loopRow_1_monday.setCellValue(lectureSchedule.getLession());
        XSSFCell loopRow_3_monday = loopRow_3.createCell(column);
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
        for (int i = 0; i < generalData.getListLectureSchedules().size(); i++) {
            LectureSchedule lectureSchedule = generalData.getListLectureSchedules().get(i);
            if (lectureSchedule.getTeacher().equals(teacher)) {
                listLectureSchedules.add(lectureSchedule);
            }
        }
        return listLectureSchedules;
    }

    private static List<BusySchedule> getListBusyScheduleByTeacher(GeneralData generalData, Teacher teacher) {
        List<BusySchedule> listBusySchedules = new ArrayList<>();
        for (int i = 0; i < generalData.getListBusySchedules().size(); i++) {
            BusySchedule busySchedule = generalData.getListBusySchedules().get(i);
            if (busySchedule.getTeacher().equals(teacher)) {
                listBusySchedules.add(busySchedule);
            }
        }
        return listBusySchedules;
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

        for (int i = 0; i < listTimes.size() - 1; i++) {
            for (int j = i + 1; j < listTimes.size(); j++) {
                if (StringUtils.compareTime(listTimes.get(i), listTimes.get(j)) < 0) {
                    String tmp = listTimes.get(i);
                    listTimes.set(i, listTimes.get(j));
                    listTimes.set(j, tmp);
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

    private static List<BusySchedule> getListBusyScheduleByTime(List<BusySchedule> listBusySchedules, String time) {
        List<BusySchedule> listBusySchedulesByTime = new ArrayList<>();
        for (BusySchedule lectureSchedule : listBusySchedules) {
            if (lectureSchedule.getHour() != null && lectureSchedule.getHour().equals(time)) {
                listBusySchedulesByTime.add(lectureSchedule);
            }
        }
        return listBusySchedulesByTime;
    }

    private static int getDayOfWeek(LectureSchedule lectureSchedule) {
        return DateTimeUtils.getDayOfWeek(lectureSchedule.getDate());
    }

    private static void fillBusy(XSSFRow loopRow_0, XSSFRow loopRow_1, XSSFRow loopRow_2, XSSFRow loopRow_3, int k) {
        XSSFCell loopCell_0 = loopRow_0.getCell(k);
        if (loopCell_0 == null) {
            loopCell_0 = loopRow_0.createCell(k);
        }
        XSSFCell loopCell_1 = loopRow_1.getCell(k);
        if (loopCell_1 == null) {
            loopCell_1 = loopRow_1.createCell(k);
        }
        XSSFCell loopCell_2 = loopRow_2.getCell(k);
        if (loopCell_2 == null) {
            loopCell_2 = loopRow_2.createCell(k);
        }
        XSSFCell loopCell_3 = loopRow_3.getCell(k);
        if (loopCell_3 == null) {
            loopCell_3 = loopRow_3.createCell(k);
        }
        XSSFCellStyle cellStyleRow0 = loopCell_0.getCellStyle();
        XSSFCellStyle cellStyleRow1 = loopCell_1.getCellStyle();
        XSSFCellStyle cellStyleRow3 = loopCell_3.getCellStyle();

        XSSFCellStyle cellStyleRow0_fillBusy = (XSSFCellStyle) cellStyleRow0.clone();
        XSSFCellStyle cellStyleRow1_fillBusy = (XSSFCellStyle) cellStyleRow1.clone();
        XSSFCellStyle cellStyleRow3_fillBusy = (XSSFCellStyle) cellStyleRow3.clone();

        cellStyleRow0_fillBusy.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        cellStyleRow0_fillBusy.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        cellStyleRow1_fillBusy.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        cellStyleRow1_fillBusy.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        cellStyleRow3_fillBusy.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        cellStyleRow3_fillBusy.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

        loopCell_0.setCellStyle(cellStyleRow0_fillBusy);
        loopCell_1.setCellStyle(cellStyleRow1_fillBusy);
        loopCell_2.setCellStyle(cellStyleRow1_fillBusy);
        loopCell_3.setCellStyle(cellStyleRow3_fillBusy);
    }

}

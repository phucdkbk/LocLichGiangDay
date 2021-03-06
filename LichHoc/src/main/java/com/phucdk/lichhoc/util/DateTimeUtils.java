/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.phucdk.lichhoc.util;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

/**
 *
 * @author Administrator
 */
public class DateTimeUtils {

    public static String convertDateToString(Date date, String format) throws Exception {
        SimpleDateFormat dateFormat = new SimpleDateFormat(format);
        try {
            if (date != null) {
                return dateFormat.format(date);
            } else {
                return "";
            }
        } catch (Exception e) {
            throw e;
        }
    }

    public static int getDayOfWeek(Date date) {
        Calendar cal = Calendar.getInstance();
        cal.setTime(date);
        return cal.get(Calendar.DAY_OF_WEEK);
    }

    public static Date getSundayOfWeek(Date startDateOfWeek) {
        Calendar cal = Calendar.getInstance();
        cal.setTime(startDateOfWeek);
        cal.setFirstDayOfWeek(Calendar.MONDAY);
        cal.set(Calendar.DAY_OF_WEEK, Calendar.MONDAY);
        cal.add(Calendar.DATE, 6);
        return cal.getTime();
    }

    public static Date addDate(Date date, int increament) {
        Calendar cal = Calendar.getInstance();
        cal.setTime(date);
        cal.add(Calendar.DATE, increament);
        return cal.getTime();
    }

    public static boolean equalDate(Date date1, Date date2) {
        return getZeroTimeDate(date1).compareTo(getZeroTimeDate(date2)) == 0;
    }

    public static Date getZeroTimeDate(Date fecha) {
        Date res = fecha;
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(fecha);
        calendar.set(Calendar.HOUR_OF_DAY, 0);
        calendar.set(Calendar.MINUTE, 0);
        calendar.set(Calendar.SECOND, 0);
        calendar.set(Calendar.MILLISECOND, 0);
        res = calendar.getTime();
        return res;
    }

}

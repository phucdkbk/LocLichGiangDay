/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.phucdk.lichhoc.util;

/**
 *
 * @author Administrator
 */
public class StringUtils {

    public static boolean isEmpty(String str) {
        if (str == null) {
            return true;
        } else {
            if ("".equals(str.trim())) {
                return true;
            }
        }
        return false;
    }

    /**
     * return 1 if time1 before time2 return -1 if time1 after time2
     *
     * @param time1
     * @param time2
     * @return
     */
    public static int compareTime(String time1, String time2) {
        int compare = 1;
        if (time1 != null && time2 != null) {
            String[] splitTime1_hour = time1.split(" ");
            String[] splitTime1_startHour = splitTime1_hour[0].split("-");

            int hourTime1 = 0;
            int hourTime2 = 0;

            if (StringUtils.isNumeric(splitTime1_startHour[0].trim())) {
                hourTime1 = Integer.parseInt(splitTime1_startHour[0].trim());
            } else {
                String[] splitTime1_startHour_tmp = splitTime1_hour[0].split(":");
                if (StringUtils.isNumeric(splitTime1_startHour_tmp[0].trim())) {
                    hourTime1 = Integer.parseInt(splitTime1_startHour_tmp[0].trim());
                }
            }

            String[] splitTime2_hour = time2.split(" ");
            String[] splitTime2_startHour = splitTime2_hour[0].split("-");

            if (StringUtils.isNumeric(splitTime2_startHour[0].trim())) {
                hourTime2 = Integer.parseInt(splitTime2_startHour[0].trim());
            } else {
                String[] splitTime2_startHour_tmp = splitTime2_hour[0].split(":");
                if (StringUtils.isNumeric(splitTime2_startHour_tmp[0].trim())) {
                    hourTime2 = Integer.parseInt(splitTime2_startHour_tmp[0].trim());
                }
            }

            if ("pm".equals(splitTime1_hour[splitTime1_hour.length - 1].trim())
                    && "am".equals(splitTime2_hour[splitTime2_hour.length - 1].trim())) {
                compare = -1;
            } else {
                if (hourTime1 > hourTime2) {
                    compare = -1;
                }
            }
        }
        return compare;
    }

    public static boolean isNumeric(String str) {
        if (str == null || "".equals(str.trim())) {
            return false;
        }
        int sz = str.length();
        for (int i = 0; i < sz; i++) {
            if (!Character.isDigit(str.charAt(i))) {
                return false;
            }
        }
        return true;
    }
}

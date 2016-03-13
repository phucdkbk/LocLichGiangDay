/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.phucdk.lichhoc.main;

import com.phucdk.lichhoc.object.GeneralData;
import com.phucdk.lichhoc.object.HeaderFooter;
import com.phucdk.lichhoc.util.ExcelExportUtil;
import com.phucdk.lichhoc.util.ExcelReadDataUtil;

/**
 *
 * @author Administrator
 */
public class Test {

    public static void main(String[] args) throws Exception {
//        GeneralData generalData = ExcelReadDataUtil.readData("D:\\20150831\\Projects\\LoclichCongtac\\Run app 29.02-06.03.xlsx",
//                "D:\\20150831\\Projects\\LoclichCongtac\\LocLichGiangDay\\Busy schedule - General 20160307.xlsx");
//        HeaderFooter headerFooter = ExcelReadDataUtil.readHeaderFooter("D:\\20150831\\Projects\\LoclichCongtac\\LocLichGiangDay\\header_footer.xlsx");
//        ExcelExportUtil.exportFile(generalData, headerFooter , "D:\\20150831\\Projects\\LoclichCongtac\\output\\");
        GeneralData generalData = ExcelReadDataUtil.readData("D:\\Projects\\LichGiangDay\\LocLichGiangDay\\HN - LICH TONG TUAN 07.03 to 14.03_Run app_v3.xlsx",
                "D:\\Projects\\LichGiangDay\\LocLichGiangDay\\Busy schedule - General 20160307.xlsx");
        HeaderFooter headerFooter = ExcelReadDataUtil.readHeaderFooter("D:\\Projects\\LichGiangDay\\LocLichGiangDay\\header_footer.xlsx");
        ExcelExportUtil.exportFile(generalData, headerFooter , "D:\\Projects\\LichGiangDay\\output\\");
    }

}

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.phucdk.lichhoc.main;

import com.phucdk.lichhoc.object.GeneralData;
import com.phucdk.lichhoc.util.ExcelExportUtil;
import com.phucdk.lichhoc.util.ExcelReadDataUtil;

/**
 *
 * @author Administrator
 */
public class Test {

    public static void main(String[] args) throws Exception {
        GeneralData generalData = ExcelReadDataUtil.readData("D:\\20150831\\Projects\\LoclichCongtac\\1. General file.xlsx");
        ExcelExportUtil.exportFile(generalData, "D:\\20150831\\Projects\\LoclichCongtac\\output\\");
    }

}
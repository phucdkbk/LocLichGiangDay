/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.phucdk.lichhoc.object;

import java.util.ArrayList;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Administrator
 */
public class HeaderFooter {

    XSSFWorkbook workbook;
    private List<XSSFRow> listHeaderRows;
    private List<XSSFRow> listFooterRows;

    public XSSFWorkbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    public List<XSSFRow> getListHeaderRows() {
        return listHeaderRows;
    }

    public void setListHeaderRows(List<XSSFRow> listHeaderRows) {
        this.listHeaderRows = listHeaderRows;
    }

    public List<XSSFRow> getListFooterRows() {
        return listFooterRows;
    }

    public void setListFooterRows(List<XSSFRow> listFooterRows) {
        this.listFooterRows = listFooterRows;
    }

    public HeaderFooter() {
        this.listFooterRows = new ArrayList<>();
        this.listHeaderRows = new ArrayList<>();
    }

}

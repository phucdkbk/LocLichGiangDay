/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.phucdk.lichhoc.util;

import org.apache.poi.ss.util.CellRangeAddress;

/**
 *
 * @author Administrator
 */
public class MergeRegion {

    private int fromRow;
    private int toRow;
    private int fromColumn;
    private int toColumn;
    private CellRangeAddress cellRangeAddress;

    public int getFromRow() {
        return fromRow;
    }

    public void setFromRow(int fromRow) {
        this.fromRow = fromRow;
    }

    public int getToRow() {
        return toRow;
    }

    public void setToRow(int toRow) {
        this.toRow = toRow;
    }

    public int getFromColumn() {
        return fromColumn;
    }

    public void setFromColumn(int fromColumn) {
        this.fromColumn = fromColumn;
    }

    public int getToColumn() {
        return toColumn;
    }

    public void setToColumn(int toColumn) {
        this.toColumn = toColumn;
    }

    public CellRangeAddress getCellRangeAddress() {
        return cellRangeAddress;
    }

    public void setCellRangeAddress(CellRangeAddress cellRangeAddress) {
        this.cellRangeAddress = cellRangeAddress;
    }

    public MergeRegion(CellRangeAddress cellRangeAddress) {
        this.cellRangeAddress = cellRangeAddress;
        String cellRangeString = this.cellRangeAddress.formatAsString();
        String[] arrRangeString = cellRangeString.split(":");
    }

}

package com.jg.poiet;

import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 对原生poi的扩展
 */
public class NiceXSSFWorkbook extends XSSFWorkbook {

    private static Logger logger = LoggerFactory.getLogger(NiceXSSFWorkbook.class);

    protected Map<String, CellRangeAddress> cellRangeAddressMap = new HashMap<>();
    protected Map<CellRangeAddress, Integer> cellRangeAddressIntegerMap = new HashMap<>();
    protected List<CellRangeAddress> allCellRangeAddress = new ArrayList<>();

    public NiceXSSFWorkbook() {
        super();
    }

    public NiceXSSFWorkbook(InputStream in) throws IOException {
        super(in);
        buildAllCellRangeAddress();
    }

    private void buildAllCellRangeAddress() {
        for (int i = 0; i < this.getNumberOfSheets(); i++) {
            XSSFSheet xssfSheet = this.getSheetAt(i);
            List<CellRangeAddress> cellRangeAddresses = xssfSheet.getMergedRegions();
            if (null == cellRangeAddresses) continue;
            allCellRangeAddress.addAll(cellRangeAddresses);
            for (int j = 0; j < xssfSheet.getNumMergedRegions(); j++) {
                CellRangeAddress cellAddress = xssfSheet.getMergedRegion(j);
                cellRangeAddressIntegerMap.put(cellAddress, j);
                int firstRow = cellAddress.getFirstRow();
                int firstColumn = cellAddress.getFirstColumn();
                for (int ri = firstRow; ri <= cellAddress.getLastRow(); ri++) {
                    for (int cj = firstColumn; cj <= cellAddress.getLastColumn(); cj++) {
                        String key = this.getCellRangeAddressMapKey(i, ri, cj);
                        cellRangeAddressMap.put(key, cellAddress);
                    }
                }
            }
        }
    }

    public void updateCellRangeAddress() {
        cellRangeAddressMap = new HashMap<>();
        allCellRangeAddress = new ArrayList<>();
        cellRangeAddressIntegerMap = new HashMap<>();
        buildAllCellRangeAddress();
    }

    private String getCellRangeAddressMapKey(int sheetIndex, int rowIndex, int columnIndex) {
        return sheetIndex + "_" + rowIndex + "_" + columnIndex;
    }

    /**
     * 获取合并单元格
     * @param sheetIndex    sheet index
     * @param rowIndex  row index
     * @param columnIndex   column index
     * @return  CellRangeAddress
     */
    public CellRangeAddress getCellRangeAddress(int sheetIndex, int rowIndex, int columnIndex) {
        return cellRangeAddressMap.get(getCellRangeAddressMapKey(sheetIndex, rowIndex, columnIndex));
    }

    /**
     * 获取合并单元格
     * @param cell  单元格
     * @return  CellRangeAddress
     */
    public CellRangeAddress getCellRangeAddress(XSSFCell cell) {
        int sheetIndex = this.getSheetIndex(cell.getSheet());
        int rowIndex = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();
        return getCellRangeAddress(sheetIndex, rowIndex, columnIndex);
    }

    /**
     * 在第sheetIndex个sheet中的第insertRowIndex行之前插入insertNum行
     * @param sheetIndex    sheetIndex
     * @param insertRowIndex    insertRowIndex
     * @param insertNum insertNum
     */
    public void insertRowsBefore(int sheetIndex, int insertRowIndex, int insertNum) {
        if (insertNum <= 0) return;
        //TO DO 插入行
        XSSFSheet sheet = this.getSheetAt(sheetIndex);
        if (null == sheet) return ;
        sheet.createRow(sheet.getLastRowNum() + insertNum);
        for (int i = sheet.getLastRowNum(); i >= insertRowIndex + insertNum; i--) {
            XSSFRow targetRow = sheet.createRow(i);
            XSSFRow sourceRow = sheet.getRow(i - insertNum);
            this.copyRow(targetRow, sourceRow);
        }
        for (int i = insertRowIndex + insertNum - 1; i >= insertRowIndex; i--) {
            sheet.createRow(i);
        }
        updateCellRangeAddress();
    }

    /**
     * 在第sheetIndex个sheet中的第insertRowIndex行之后插入insertNum行
     * @param sheetIndex    sheetIndex
     * @param insertRowIndex    insertRowIndex
     * @param insertNum insertNum
     */
    public void insertRowsAfter(int sheetIndex, int insertRowIndex, int insertNum) {
        if (insertNum <= 0) return;
        //TO DO 插入行
        XSSFSheet sheet = this.getSheetAt(sheetIndex);
        if (null == sheet) return ;
        sheet.createRow(sheet.getLastRowNum() + insertNum);
        for (int i = sheet.getLastRowNum(); i > insertRowIndex + insertNum; i--) {
            XSSFRow targetRow = sheet.createRow(i);
            XSSFRow sourceRow = sheet.getRow(i - insertNum);
            this.copyRow(targetRow, sourceRow);
        }
        for (int i = insertRowIndex + insertNum; i > insertRowIndex; i--) {
            sheet.createRow(i);
        }
        updateCellRangeAddress();
    }

    /**
     * 复制行
     * @param targetRow 目标行
     * @param sourceRow 源行
     */
    private void copyRow(XSSFRow targetRow, XSSFRow sourceRow) {
        targetRow.copyRowFrom(sourceRow, new CellCopyPolicy());
        XSSFSheet sheet = targetRow.getSheet();
        for (int i = 0; i <= sourceRow.getLastCellNum(); i++) {
            XSSFCell xssfCell = sourceRow.getCell(i);
            if (null == xssfCell) continue;
            CellRangeAddress cellAddresses = getCellRangeAddress(xssfCell);
            if (null == cellAddresses) continue;
            if (xssfCell.getRowIndex() != cellAddresses.getFirstRow()
                    || xssfCell.getColumnIndex() != cellAddresses.getFirstColumn()) continue;
            int firstRow = cellAddresses.getFirstRow();
            int firstColumn = cellAddresses.getFirstColumn();
            int lastRow = cellAddresses.getLastRow();
            int lastColumn = cellAddresses.getLastColumn();
            int regionIndex = cellRangeAddressIntegerMap.get(cellAddresses);
            sheet.removeMergedRegion(cellRangeAddressIntegerMap.get(cellAddresses));
            addMergedRegion(this.getSheetIndex(sheet),
                    targetRow.getRowNum(),
                    targetRow.getRowNum() + (lastRow - firstRow),
                    firstColumn, lastColumn);
        }
    }

    /**
     * 删除行
     * @param sheetIndex    sheetIndex
     * @param rowIndex  rowIndex
     */
    public void removeRow(int sheetIndex, int rowIndex) {
        if (rowIndex < 0) return;
        XSSFSheet sheet = this.getSheetAt(sheetIndex);
        if (null == sheet) return ;
        for (int i = rowIndex; i < sheet.getLastRowNum(); i++) {
            XSSFRow targetRow = sheet.createRow(i);
            XSSFRow sourceRow = sheet.getRow(i + 1);
            targetRow.copyRowFrom(sourceRow, new CellCopyPolicy());
        }
        sheet.removeRow(sheet.getRow(sheet.getLastRowNum()));
        updateCellRangeAddress();
    }

    /**
     * 合并单元格
     * @param sheetIndex    sheetIndex
     * @param firstRowIndex firstRowIndex
     * @param lastRowIndex  lastRowIndex
     * @param firstColumnIndex  firstColumnIndex
     * @param lastColumnIndex   lastColumnIndex
     * @param isUpdate  是否更新
     */
    public void addMergedRegion(int sheetIndex, int firstRowIndex, int lastRowIndex, int firstColumnIndex, int lastColumnIndex, boolean isUpdate) {
        if (firstRowIndex > lastRowIndex || firstColumnIndex > lastColumnIndex) return ;
        if (firstRowIndex == lastRowIndex && firstColumnIndex == lastColumnIndex) return ;
        XSSFSheet sheet = this.getSheetAt(sheetIndex);
        CellRangeAddress region = new CellRangeAddress(firstRowIndex, lastRowIndex, firstColumnIndex, lastColumnIndex);
        sheet.addMergedRegion(region);
        if (isUpdate) {
            updateCellRangeAddress();
        } else {
            allCellRangeAddress.add(region);
            for (int ri = firstRowIndex; ri <= lastRowIndex; ri++) {
                for (int cj = firstColumnIndex; cj <= lastColumnIndex; cj++) {
                    String key = this.getCellRangeAddressMapKey(sheetIndex, ri, cj);
                    cellRangeAddressMap.put(key, region);
                }
            }
        }
    }

    /**
     * 合并单元格
     * @param sheetIndex    sheetIndex
     * @param firstRowIndex firstRowIndex
     * @param lastRowIndex  lastRowIndex
     * @param firstColumnIndex  firstColumnIndex
     * @param lastColumnIndex   lastColumnIndex
     */
    public void addMergedRegion(int sheetIndex, int firstRowIndex, int lastRowIndex, int firstColumnIndex, int lastColumnIndex) {
        this.addMergedRegion(sheetIndex, firstRowIndex, lastRowIndex, firstColumnIndex, lastColumnIndex, true);
    }

    /**
     * 拆分单元格
     * @param cell   单元格
     * @param isUpdate 是否更新合并单元格信息
     */
    public void removeMergedRegion(XSSFCell cell, boolean isUpdate) {
        CellRangeAddress cellAddress = this.getCellRangeAddress(cell);
        if (null == cellAddress) return;
        cell.getSheet().removeMergedRegion(cellRangeAddressIntegerMap.get(cellAddress));
        if (isUpdate) {
            updateCellRangeAddress();
        }
    }
}

package com.jg.poiet.template.cell;

import com.jg.poiet.template.ElementTemplate;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class CellTemplate extends ElementTemplate {

    protected XSSFCell cell;

    public CellTemplate() {}

    public CellTemplate(String tagName, XSSFCell cell) {
        this.tagName = tagName;
        this.cell = cell;
    }

    public XSSFRow getRow() {
        return cell.getRow();
    }

    public Integer getCellColumnIndex() {
        return cell.getColumnIndex();
    }

    public XSSFCell getBeforeCell() {
        Integer cellColumnIndex = getCellColumnIndex();
        if (null == cellColumnIndex) return null;
        return cellColumnIndex == 0 ? null : getRow().getCell(cellColumnIndex - 1);
    }

    public XSSFCell getAfterCell() {
        Integer cellColumnIndex = getCellColumnIndex();
        if (null == cellColumnIndex) return null;
        return cellColumnIndex == (getRow().getLastCellNum()) ? null
                : getRow().getCell(cellColumnIndex + 1);
    }

    public XSSFCell getCell() {
        return cell;
    }

    public void setCell(XSSFCell cell) {
        this.cell = cell;
    }
}

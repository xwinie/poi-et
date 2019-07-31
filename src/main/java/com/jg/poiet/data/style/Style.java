package com.jg.poiet.data.style;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.HashMap;
import java.util.Map;

/**
 * 样式
 */
public class Style {

    private Map<XSSFCellStyle, XSSFCellStyle> cellStyleMap = new HashMap<>();
    /**
     * 字体颜色
     */
    private int[] color;
    /**
     * 字体
     */
    private String fontName;

    /**
     * 字体大小
     */
    private int fontSize = -1;
    /**
     * 粗体
     */
    private Boolean isBold;
    /**
     * 斜体
     */
    private Boolean isItalic;
    /**
     *  字体下划线样式
     */
    private FontUnderline fontUnderline;
    /**
     * 前景色
     */
    private int[] foregroundColor;
    /**
     * 前景色填充方式
     */
    private FillPatternType fillPatternType = FillPatternType.NO_FILL;
    /**
     * 水平对齐方式
     */
    private HorizontalAlignment horizontalAlignment;
    /**
     * 垂直对齐方式
     */
    private VerticalAlignment verticalAlignment;
    /**
     * 是否自动换行
     */
    private Boolean wrapText;
    /**
     * 单元格左边框样式
     */
    private BorderStyle leftBorderStyle;
    /**
     * 单元格左边框颜色
     */
    private int[] leftBorderColor;
    /**
     * 单元格右边框样式
     */
    private BorderStyle rightBorderStyle;
    /**
     * 单元格右边框颜色
     */
    private int[] rightBorderColor;
    /**
     * 单元格上边框样式
     */
    private BorderStyle topBorderStyle;
    /**
     * 单元格上边框颜色
     */
    private int[] topBorderColor;
    /**
     * 单元格下边框样式
     */
    private BorderStyle bottomBorderStyle;
    /**
     * 单元格下边框颜色
     */
    private int[] bottomBorderColor;
    /**
     * 行高
     */
    private float rowHeight = -1f;
    /**
     * 列宽
     */
    private int columnWidth = -1;

    public int getColumnWidth() {
        return columnWidth;
    }

    public void setColumnWidth(int columnWidth) {
        this.columnWidth = columnWidth;
    }

    public float getRowHeight() {
        return rowHeight;
    }

    public void setRowHeight(float rowHeight) {
        this.rowHeight = rowHeight;
    }

    public BorderStyle getLeftBorderStyle() {
        return leftBorderStyle;
    }

    public void setLeftBorderStyle(BorderStyle leftBorderStyle) {
        this.leftBorderStyle = leftBorderStyle;
    }

    public int[] getLeftBorderColor() {
        return leftBorderColor;
    }

    public void setLeftBorderColor(int[] leftBorderColor) {
        this.leftBorderColor = leftBorderColor;
    }

    public BorderStyle getRightBorderStyle() {
        return rightBorderStyle;
    }

    public void setRightBorderStyle(BorderStyle rightBorderStyle) {
        this.rightBorderStyle = rightBorderStyle;
    }

    public int[] getRightBorderColor() {
        return rightBorderColor;
    }

    public void setRightBorderColor(int[] rightBorderColor) {
        this.rightBorderColor = rightBorderColor;
    }

    public BorderStyle getTopBorderStyle() {
        return topBorderStyle;
    }

    public void setTopBorderStyle(BorderStyle topBorderStyle) {
        this.topBorderStyle = topBorderStyle;
    }

    public int[] getTopBorderColor() {
        return topBorderColor;
    }

    public void setTopBorderColor(int[] topBorderColor) {
        this.topBorderColor = topBorderColor;
    }

    public BorderStyle getBottomBorderStyle() {
        return bottomBorderStyle;
    }

    public void setBottomBorderStyle(BorderStyle bottomBorderStyle) {
        this.bottomBorderStyle = bottomBorderStyle;
    }

    public int[] getBottomBorderColor() {
        return bottomBorderColor;
    }

    public void setBottomBorderColor(int[] bottomBorderColor) {
        this.bottomBorderColor = bottomBorderColor;
    }

    public Boolean getWrapText() {
        return wrapText;
    }

    public void setWrapText(Boolean wrapText) {
        this.wrapText = wrapText;
    }

    public HorizontalAlignment getHorizontalAlignment() {
        return horizontalAlignment;
    }

    public void setHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
    }

    public VerticalAlignment getVerticalAlignment() {
        return verticalAlignment;
    }

    public void setVerticalAlignment(VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
    }

    public FillPatternType getFillPatternType() {
        return fillPatternType;
    }

    public void setFillPatternType(FillPatternType fillPatternType) {
        this.fillPatternType = fillPatternType;
    }

    public int[] getForegroundColor() {
        return foregroundColor;
    }

    public void setForegroundColor(int[] foregroundColor) {
        this.foregroundColor = foregroundColor;
    }

    public int getFontSize() {
        return fontSize;
    }

    public void setFontSize(int fontSize) {
        this.fontSize = fontSize;
    }

    public Boolean getBold() {
        return isBold;
    }

    public void setBold(Boolean bold) {
        isBold = bold;
    }

    public Boolean getItalic() {
        return isItalic;
    }

    public void setItalic(Boolean italic) {
        isItalic = italic;
    }

    public FontUnderline getFontUnderline() {
        return fontUnderline;
    }

    public void setFontUnderline(FontUnderline fontUnderline) {
        this.fontUnderline = fontUnderline;
    }

    public String getFontName() {
        return fontName;
    }

    public void setFontName(String fontName) {
        this.fontName = fontName;
    }

    public int[] getColor() {
        return color;
    }

    public void setColor(int[] color) {
        this.color = color;
    }

    public XSSFCellStyle getCellStyle(XSSFCellStyle key) {
        return cellStyleMap.get(key);
    }

    public void setCellStyle(XSSFCellStyle key, XSSFCellStyle value) {
        this.cellStyleMap.put(key, value);
    }

}

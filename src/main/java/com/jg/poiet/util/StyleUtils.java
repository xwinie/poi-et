package com.jg.poiet.util;

import com.jg.poiet.XSSFTemplate;
import com.jg.poiet.exception.RenderException;
import com.jg.poiet.data.style.Style;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.awt.*;

/**
 * 样式工具类
 */
public class StyleUtils {

    /**
     * 设置cell的样式
     *
     * @param cell  设置的单元格
     * @param style 样式对象
     */
    public static void styleCell(XSSFTemplate template, XSSFCell cell, Style style) {
        if (null == cell || null == style) return;
        XSSFCellStyle oldCellStyle = cell.getCellStyle();
        XSSFCellStyle cellStyle = style.getCellStyle(oldCellStyle);
        XSSFSheet sheet = cell.getSheet();
        if (null == cellStyle) {
            cellStyle = sheet.getWorkbook().createCellStyle();

            //设置前景色
            if (null == style.getForegroundColor()) {
                cellStyle.setFillForegroundColor(oldCellStyle.getFillForegroundXSSFColor());
                cellStyle.setFillPattern(oldCellStyle.getFillPattern());
            } else {
                cellStyle.setFillForegroundColor(getXSSFColorFromRGB(style.getForegroundColor()));// 设置背景色
                cellStyle.setFillPattern(style.getFillPatternType());
            }
            //设置文字对齐方式和自动换行
            if (null == style.getHorizontalAlignment()) {
                cellStyle.setAlignment(oldCellStyle.getAlignment());
            } else {
                cellStyle.setAlignment(style.getHorizontalAlignment());
            }
            if (null == style.getVerticalAlignment()) {
                cellStyle.setVerticalAlignment(oldCellStyle.getVerticalAlignment());
            } else {
                cellStyle.setVerticalAlignment(style.getVerticalAlignment());
            }
            if (null == style.getWrapText()) {
                cellStyle.setWrapText(oldCellStyle.getWrapText());
            } else {
                cellStyle.setWrapText(style.getWrapText());
            }
            //设置边框
            if (null == style.getLeftBorderStyle()) {
                cellStyle.setBorderLeft(oldCellStyle.getBorderLeft());
            } else {
                cellStyle.setBorderLeft(style.getLeftBorderStyle());
            }
            if (null == style.getLeftBorderColor()) {
                if (null != oldCellStyle.getLeftBorderXSSFColor()) {
                    cellStyle.setLeftBorderColor(oldCellStyle.getLeftBorderXSSFColor());
                }
            } else {
                cellStyle.setLeftBorderColor(getXSSFColorFromRGB(style.getLeftBorderColor()));
            }
            if (null == style.getTopBorderStyle()) {
                cellStyle.setBorderTop(oldCellStyle.getBorderTop());
            } else {
                cellStyle.setBorderTop(style.getTopBorderStyle());
            }
            if (null == style.getTopBorderColor()) {
                if (null != oldCellStyle.getTopBorderXSSFColor()) {
                    cellStyle.setTopBorderColor(oldCellStyle.getTopBorderXSSFColor());
                }
            } else {
                cellStyle.setTopBorderColor(getXSSFColorFromRGB(style.getTopBorderColor()));
            }
            if (null == style.getRightBorderStyle()) {
                cellStyle.setBorderRight(oldCellStyle.getBorderRight());
            } else {
                cellStyle.setBorderRight(style.getRightBorderStyle());
            }
            if (null == style.getRightBorderColor()) {
                if (null != oldCellStyle.getRightBorderXSSFColor()) {
                    cellStyle.setRightBorderColor(oldCellStyle.getRightBorderXSSFColor());
                }
            } else {
                cellStyle.setRightBorderColor(getXSSFColorFromRGB(style.getRightBorderColor()));
            }
            if (null == style.getBottomBorderStyle()) {
                cellStyle.setBorderBottom(oldCellStyle.getBorderBottom());
            } else {
                cellStyle.setBorderBottom(style.getBottomBorderStyle());
            }
            if (null == style.getBottomBorderColor()) {
                if (null != oldCellStyle.getBottomBorderXSSFColor()) {
                    cellStyle.setBottomBorderColor(oldCellStyle.getBottomBorderXSSFColor());
                }
            } else {
                cellStyle.setBottomBorderColor(getXSSFColorFromRGB(style.getBottomBorderColor()));
            }
            //设置字体
            XSSFFont font = sheet.getWorkbook().createFont();
            XSSFFont oldFont = oldCellStyle.getFont();
            if (null == style.getColor()) {
                font.setColor(oldFont.getXSSFColor());
            } else {
                font.setColor(getXSSFColorFromRGB(style.getColor()));
            }
            if (null == style.getFontName()) {
                font.setFontName(oldFont.getFontName());
            } else {
                font.setFontName(style.getFontName());
            }
            if (style.getFontSize() < 0) {
                font.setFontHeightInPoints(oldFont.getFontHeightInPoints());
            } else {
                font.setFontHeightInPoints((short)style.getFontSize());
            }
            if (null == style.getBold()) {
                font.setBold(oldFont.getBold());
            } else {
                font.setBold(style.getBold());
            }
            if (null == style.getItalic()) {
                font.setItalic(oldFont.getItalic());
            } else {
                font.setItalic(style.getItalic());
            }
            if (null == style.getFontUnderline()) {
                font.setUnderline(oldFont.getUnderline());
            } else {
                font.setUnderline(style.getFontUnderline());
            }
            cellStyle.setFont(font);
            style.setCellStyle(oldCellStyle, cellStyle);
        }
        //设置样式
        CellRangeAddress cellAddresses = template.getXSSFWorkbook().getCellRangeAddress(cell);
        if (null == cellAddresses) {    //非合并单元格
            cell.setCellStyle(cellStyle);
            //设置列宽（注：只有非合并单元格可以设置列宽）
            if (style.getColumnWidth() >= 0) {
                sheet.setColumnWidth(cell.getColumnIndex(), style.getColumnWidth());
            }
        } else {    //合并单元格
            //只有当前单元格为合并单元格的第一个单元格可以设置样式
            if (cellAddresses.getFirstRow() == cell.getRowIndex()
                    && cellAddresses.getFirstColumn() == cell.getColumnIndex()) {
                int firstRow = cellAddresses.getFirstRow();
                int lastRow = cellAddresses.getLastRow();
                int firstColumn = cellAddresses.getFirstColumn();
                int lastColumn = cellAddresses.getLastColumn();
                for (int i = firstRow; i <= lastRow; i++) {
                    XSSFRow row = sheet.getRow(i);
                    if (null == row) {
                        row = sheet.createRow(i);
                    }
                    for (int j = firstColumn; j <= lastColumn; j++) {
                        XSSFCell xssfCell = row.getCell(j);
                        if (null == xssfCell) {
                            xssfCell = row.createCell(j);
                        }
                        xssfCell.setCellStyle(cellStyle);
                    }
                }
            }
        }
        //设置行高
        if (style.getRowHeight() >= 0) {
            cell.getRow().setHeightInPoints(style.getRowHeight());
        }

    }

    /**
     * 设置样式
     * @param template  template
     * @param cell  cell
     * @param styles    styles
     */
    public static void styleCell(XSSFTemplate template, XSSFCell cell, Style... styles) {
        if (null == cell || null == styles || styles.length <= 0) return;
        for (Style style : styles) {
            styleCell(template, cell, style);
        }
    }

    private static XSSFColor getXSSFColorFromRGB(int[] color) {
        if (null == color || color.length != 3) {
            throw new RenderException("color is null or color is error");
        }
        int r = color[0];
        int g = color[1];
        int b = color[2];
        if (r < 0 || r > 255 || g < 0 || g > 255 || b < 0 || b > 255) {
            throw new RenderException("RGB value must be in 0-255");
        }
        return new XSSFColor(new Color(r, g, b));
    }

}

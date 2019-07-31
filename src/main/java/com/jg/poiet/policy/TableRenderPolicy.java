package com.jg.poiet.policy;

import com.jg.poiet.data.CellRenderData;
import com.jg.poiet.data.RowRenderData;
import com.jg.poiet.data.TableRenderData;
import com.jg.poiet.NiceXSSFWorkbook;
import com.jg.poiet.XSSFTemplate;
import com.jg.poiet.data.style.Style;
import com.jg.poiet.template.cell.CellTemplate;
import com.jg.poiet.util.StyleUtils;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.Iterator;
import java.util.List;

/**
 * 表格处理
 */
public class TableRenderPolicy extends AbstractRenderPolicy<TableRenderData> {

    @Override
    protected boolean validate(TableRenderData data) {
        if (!(data).isSetBody() && !(data).isSetHeader()) {
            logger.debug("Empty MiniTableRenderData datamodel: {}", data);
            return false;
        }
        return true;
    }

    @Override
    public void doRender(CellTemplate cellTemplate, TableRenderData data, XSSFTemplate template) {
        XSSFCell cell = cellTemplate.getCell();
        Helper.renderTable(template, cell, data);
    }

    public static class Helper {
        /**
         * 渲染表格数据
         * @param template template
         * @param cell  cell
         * @param tableData tableData
         */
        public static void renderTable(XSSFTemplate template, XSSFCell cell, TableRenderData tableData) {
            NiceXSSFWorkbook workbook = template.getXSSFWorkbook();
            XSSFSheet sheet = cell.getSheet();
            cell.setCellValue("");  //将该单元格值置空
            insertRow(workbook, cell, tableData);   //插入行
            int rowIndex = cell.getRowIndex();
            if (tableData.isSetHeader()) {
                List<RowRenderData> headerData = tableData.getHeader();
                for (RowRenderData rowRenderData : headerData) {
                    renderRow(template, sheet.getRow(rowIndex), rowRenderData, cell.getColumnIndex(), tableData.getHeaderStyle());
                    rowIndex ++;
                }
            }
            if (tableData.isSetBody()) {
                List<RowRenderData> bodyData = tableData.getRowDatas();
                Iterator<RowRenderData> iterator = bodyData.iterator();
                for (RowRenderData rowRenderData : bodyData) {
                    renderRow(template, sheet.getRow(rowIndex), rowRenderData, cell.getColumnIndex(), tableData.getBodyStyle());
                    rowIndex ++;
                }
            }
        }

        /**
         * 渲染行数据
         * @param template  template
         * @param row   row
         * @param rowRenderData rowRenderData
         * @param columnIndex   columnIndex
         * @param style style
         */
        private static void renderRow(XSSFTemplate template, XSSFRow row, RowRenderData rowRenderData, int columnIndex, Style style) {
            NiceXSSFWorkbook workbook = template.getXSSFWorkbook();
            List<CellRenderData> cellRenderDataList = rowRenderData.getCellDatas();
            for (int i = 0; i < cellRenderDataList.size(); i++) {
                CellRenderData cellRenderData = cellRenderDataList.get(i);
                if (null == cellRenderData || null == cellRenderData.getRenderData()) continue;
                XSSFCell xssfCell = row.getCell(columnIndex);
                if (null == xssfCell) {
                    xssfCell = row.createCell(columnIndex);
                }
                CellRangeAddress cellAddress = workbook.getCellRangeAddress(xssfCell);
                if (null == cellAddress) {
                    if (cellRenderData.getRowspan() > 0 || cellRenderData.getColspan() > 0) {   //需要合并单元格
                        workbook.addMergedRegion(template.getXSSFWorkbook().getSheetIndex(row.getSheet()),
                                xssfCell.getRowIndex(), xssfCell.getRowIndex() + cellRenderData.getRowspan(),
                                xssfCell.getColumnIndex(), xssfCell.getColumnIndex() + cellRenderData.getColspan(), false);
                    }
                    StyleUtils.styleCell(template, xssfCell, style, rowRenderData.getStyle());
                    TextRenderPolicy.Helper.renderTextCell(xssfCell, cellRenderData.getRenderData(), template);
                    columnIndex += cellRenderData.getColspan() + 1;
                } else {
                    if (cellAddress.getFirstRow() == xssfCell.getRowIndex()
                            && cellAddress.getFirstColumn() == xssfCell.getColumnIndex()) {
                        StyleUtils.styleCell(template, xssfCell, style, rowRenderData.getStyle());
                        TextRenderPolicy.Helper.renderTextCell(xssfCell, cellRenderData.getRenderData(), template);
                    } else {
                        i--;
                    }
                    columnIndex = cellAddress.getLastColumn() + 1;
                }
            }
            template.getXSSFWorkbook().updateCellRangeAddress();
        }

        /**
         * 插入行
         * @param workbook  workbook
         * @param cell  cell
         * @param tableData tableData
         */
        private static void insertRow(NiceXSSFWorkbook workbook, XSSFCell cell, TableRenderData tableData) {
            int insertNum = 0;
            if (tableData.isSetHeader()) {
                insertNum += tableData.getHeader().size();
            }
            if (tableData.isSetBody()) {
                insertNum += tableData.getRowDatas().size();
            }
            XSSFSheet sheet = cell.getSheet();
            workbook.insertRowsAfter(workbook.getSheetIndex(sheet), cell.getRowIndex(), insertNum - 1);
            for (int i = cell.getRowIndex() + 1; i <= cell.getRowIndex() + insertNum - 1; i++) {
                XSSFRow row = sheet.createRow(i);
                row.copyRowFrom(cell.getRow(), new CellCopyPolicy());
            }
            XSSFRow xssfRow = cell.getRow();
            for(int i = 0; i <= xssfRow.getLastCellNum(); i++){
                XSSFCell xssfCell = xssfRow.getCell(i);
                if (null == xssfCell) continue;
                CellRangeAddress cellAddress = workbook.getCellRangeAddress(xssfCell);
                if (null == cellAddress) continue;
                if (xssfCell.getColumnIndex() != cellAddress.getFirstColumn()) continue;
                if (cellAddress.getFirstRow() == cellAddress.getLastRow()) continue;
                int firstRow = cellAddress.getFirstRow();
                int lastRow = cellAddress.getLastRow() + insertNum - 1;
                int firstColumn = cellAddress.getFirstColumn();
                int lastColumn = cellAddress.getLastColumn();
                workbook.removeMergedRegion(xssfCell, false);
                workbook.addMergedRegion(workbook.getSheetIndex(sheet), firstRow, lastRow, firstColumn, lastColumn);
            }
        }
    }

}

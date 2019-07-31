package com.jg.poiet.data;

import com.jg.poiet.data.style.Style;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * 表格行数据
 */
public class RowRenderData implements RenderData {

    private List<CellRenderData> cellDatas;

    /**
     * 表格行样式
     */
    private Style style;

    public RowRenderData() {}

    public RowRenderData(List<CellRenderData> cellDatas) {
        this.cellDatas = cellDatas;
    }

    public RowRenderData(List<TextRenderData> rowData, Style style) {
        this.cellDatas = new ArrayList<>();
        if (null != rowData) {
            for (TextRenderData data : rowData) {
                this.cellDatas.add(new CellRenderData(data));
            }
        }
        this.style = style;
    }

    public static RowRenderData build(String... cellStr) {
        List<TextRenderData> cellDatas = new ArrayList<>();
        if (null != cellStr) {
            for (String col : cellStr) {
                cellDatas.add(new TextRenderData(col));
            }
        }
        return new RowRenderData(cellDatas, null);
    }

    public static RowRenderData build(TextRenderData... cellData) {
        return new RowRenderData(null == cellData ? null : Arrays.asList(cellData), null);
    }

    public static RowRenderData build(CellRenderData... cellData) {
        return new RowRenderData(null == cellData ? null : Arrays.asList(cellData));
    }

    public RowRenderData buildStyle(Style style) {
        this.style = style;
        return this;
    }

    public int size() {
        return null == cellDatas ? 0 : cellDatas.size();
    }

    public List<CellRenderData> getCellDatas() {
        return cellDatas;
    }

    public void setCellDatas(List<CellRenderData> cellDatas) {
        this.cellDatas = cellDatas;
    }

    public Style getStyle() {
        return style;
    }

    public RowRenderData setStyle(Style style) {
        this.style = style;
        return this;
    }
}

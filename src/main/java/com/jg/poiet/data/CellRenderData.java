package com.jg.poiet.data;

import com.jg.poiet.data.style.Style;

/**
 * 单元格数据
 */
public class CellRenderData {

    protected TextRenderData renderData;

    /**
     * 单元格合并行数
     */
    protected int rowspan = 0;

    /**
     * 单元格合并列数
     */
    protected int colspan = 0;

    public CellRenderData() {
        renderData = new TextRenderData();
    }

    public CellRenderData(TextRenderData renderData) {
        this.renderData = renderData;
    }

    public CellRenderData(String text) {
        this.renderData = new TextRenderData(text);
    }

    public CellRenderData(TextRenderData renderData, int rowspan, int colspan) {
        this.renderData = renderData;
        this.rowspan = rowspan;
        this.colspan = colspan;
    }

    public CellRenderData(String text, int rowspan, int colspan) {
        this.renderData = new TextRenderData(text);
        this.rowspan = rowspan;
        this.colspan = colspan;
    }

    public TextRenderData getRenderData() {
        return renderData;
    }

    public CellRenderData setStyle(Style style) {
        this.renderData.setStyle(style);
        return this;
    }

    public CellRenderData setText(String text) {
        this.renderData.setText(text);
        return this;
    }

    public CellRenderData setRenderData(TextRenderData renderData) {
        this.renderData = renderData;
        return this;
    }

    public int getRowspan() {
        return rowspan;
    }

    public CellRenderData setRowspan(int rowspan) {
        this.rowspan = rowspan;
        return this;
    }

    public int getColspan() {
        return colspan;
    }

    public CellRenderData setColspan(int colspan) {
        this.colspan = colspan;
        return this;
    }
}

package com.jg.poiet.data;

import com.jg.poiet.data.style.Style;

import java.util.ArrayList;
import java.util.List;

/**
 * 表格数据
 */
public class TableRenderData implements RenderData {

    /**
     * 表格头部数据，可为空
     */
    private List<RowRenderData> header;
    /**
     * 表格数据
     */
    private List<RowRenderData> rowDatas;
    /**
     *  表格头部数据样式
     */
    private Style headerStyle;
    /**
     * 表格行数据样式
     */
    private Style bodyStyle;

    public TableRenderData(List<RowRenderData> datas) {
        this(null, datas);
    }

    public TableRenderData(List<RowRenderData> datas, Style bodyStyle) {
        this.rowDatas = datas;
        this.bodyStyle = bodyStyle;
    }

    public TableRenderData(RowRenderData header, List<RowRenderData> datas) {
        this(header, datas, null, null);
    }

    public TableRenderData(RowRenderData header, List<RowRenderData> datas, Style headerStyle, Style bodyStyle) {
        this.header = new ArrayList<>();
        this.header.add(header);
        this.rowDatas = datas;
        this.headerStyle = headerStyle;
        this.bodyStyle = bodyStyle;
    }

    public TableRenderData(List<RowRenderData> header, List<RowRenderData> datas, Style headerStyle, Style bodyStyle) {
        this.header = header;
        this.rowDatas = datas;
        this.headerStyle = headerStyle;
        this.bodyStyle = bodyStyle;
    }

    public boolean isSetHeader() {
        return null != header && header.size() > 0;
    }

    public boolean isSetBody() {
        return null != rowDatas && rowDatas.size() > 0;
    }

    public List<RowRenderData> getHeader() {
        return header;
    }

    public TableRenderData setHeader(List<RowRenderData> header) {
        this.header = header;
        return this;
    }

    public List<RowRenderData> getRowDatas() {
        return rowDatas;
    }

    public TableRenderData setRowDatas(List<RowRenderData> rowDatas) {
        this.rowDatas = rowDatas;
        return this;
    }

    public Style getHeaderStyle() {
        return headerStyle;
    }

    public TableRenderData setHeaderStyle(Style headerStyle) {
        this.headerStyle = headerStyle;
        return this;
    }

    public Style getBodyStyle() {
        return bodyStyle;
    }

    public TableRenderData setBodyStyle(Style bodyStyle) {
        this.bodyStyle = bodyStyle;
        return this;
    }

}

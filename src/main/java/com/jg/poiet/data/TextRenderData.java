package com.jg.poiet.data;

import com.jg.poiet.data.style.Style;

/**
 * 文本数据
 *
 */
public class TextRenderData implements RenderData {

    protected Style style;

    /**
     * \n 表示换行
     */
    protected String text;

    public TextRenderData() {}

    public TextRenderData(String text) {
        this.text = text;
    }

    public TextRenderData(String text, Style style) {
        this.style = style;
        this.text = text;
    }

    public Style getStyle() {
        return style;
    }

    public TextRenderData setStyle(Style style) {
        this.style = style;
        return this;
    }

    public String getText() {
        return text;
    }

    public TextRenderData setText(String text) {
        this.text = text;
        return this;
    }

}

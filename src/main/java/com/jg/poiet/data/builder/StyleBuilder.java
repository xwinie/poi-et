package com.jg.poiet.data.builder;

import com.jg.poiet.data.style.Style;
import org.apache.poi.ss.usermodel.*;

/**
 * 样式构建器
 */
public class StyleBuilder {

    Style style;

    private StyleBuilder() {
        style = new Style();
    }

    public static StyleBuilder newBuilder() {
        return new StyleBuilder();
    }

    public StyleBuilder buildColor(int r, int g, int b) {
        int[] rgbColor = new int[]{r, g, b};
        style.setColor(rgbColor);
        return this;
    }

    public StyleBuilder buildFontName(String fontName) {
        style.setFontName(fontName);
        return this;
    }

    public StyleBuilder buildFontSize(int fontSize) {
        style.setFontSize(fontSize);
        return this;
    }

    public StyleBuilder buildBold() {
        style.setBold(true);
        return this;
    }

    public StyleBuilder buildNotBold() {
        style.setBold(false);
        return this;
    }

    public StyleBuilder buildItalic() {
        style.setItalic(true);
        return this;
    }

    public StyleBuilder buildNotItalic() {
        style.setItalic(false);
        return this;
    }

    public StyleBuilder buildFontUnderline(FontUnderline fontUnderline) {
        style.setFontUnderline(fontUnderline);
        return this;
    }

    public StyleBuilder buildForegroundColor(int r, int g, int b) {
        int[] rgbColor = new int[]{r, g, b};
        style.setForegroundColor(rgbColor);
        if (FillPatternType.NO_FILL == style.getFillPatternType()) {
            style.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
        }
        return this;
    }

    public StyleBuilder buildFillPatternType(FillPatternType fillPatternType) {
        style.setFillPatternType(fillPatternType);
        return this;
    }

    public StyleBuilder buildHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
        style.setHorizontalAlignment(horizontalAlignment);
        return this;
    }

    public StyleBuilder buildVerticalAlignment(VerticalAlignment verticalAlignment) {
        style.setVerticalAlignment(verticalAlignment);
        return this;
    }

    public StyleBuilder buildWrapText() {
        style.setWrapText(true);
        return this;
    }

    public StyleBuilder buildNotWrapText() {
        style.setWrapText(false);
        return this;
    }

    public StyleBuilder buildLeftBorderStyle(BorderStyle borderStyle) {
        style.setLeftBorderStyle(borderStyle);
        return this;
    }

    public StyleBuilder buildLeftBorderColor(int r, int g, int b) {
        int[] rgbColor = new int[]{r, g, b};
        style.setLeftBorderColor(rgbColor);
        return this;
    }

    public StyleBuilder buildLeftBorder(BorderStyle borderStyle, int r, int g, int b) {
        style.setLeftBorderStyle(borderStyle);
        int[] rgbColor = new int[]{r, g, b};
        style.setLeftBorderColor(rgbColor);
        return this;
    }

    public StyleBuilder buildRightBorderStyle(BorderStyle borderStyle) {
        style.setRightBorderStyle(borderStyle);
        return this;
    }

    public StyleBuilder buildRightBorderColor(int r, int g, int b) {
        int[] rgbColor = new int[]{r, g, b};
        style.setRightBorderColor(rgbColor);
        return this;
    }

    public StyleBuilder buildRightBorder(BorderStyle borderStyle, int r, int g, int b) {
        style.setRightBorderStyle(borderStyle);
        int[] rgbColor = new int[]{r, g, b};
        style.setRightBorderColor(rgbColor);
        return this;
    }

    public StyleBuilder buildTopBorderStyle(BorderStyle borderStyle) {
        style.setTopBorderStyle(borderStyle);
        return this;
    }

    public StyleBuilder buildTopBorderColor(int r, int g, int b) {
        int[] rgbColor = new int[]{r, g, b};
        style.setTopBorderColor(rgbColor);
        return this;
    }

    public StyleBuilder buildTopBorder(BorderStyle borderStyle, int r, int g, int b) {
        style.setTopBorderStyle(borderStyle);
        int[] rgbColor = new int[]{r, g, b};
        style.setTopBorderColor(rgbColor);
        return this;
    }

    public StyleBuilder buildBottomBorderStyle(BorderStyle borderStyle) {
        style.setBottomBorderStyle(borderStyle);
        return this;
    }

    public StyleBuilder buildBottomBorderColor(int r, int g, int b) {
        int[] rgbColor = new int[]{r, g, b};
        style.setBottomBorderColor(rgbColor);
        return this;
    }

    public StyleBuilder buildBottomBorder(BorderStyle borderStyle, int r, int g, int b) {
        style.setBottomBorderStyle(borderStyle);
        int[] rgbColor = new int[]{r, g, b};
        style.setBottomBorderColor(rgbColor);
        return this;
    }

    public StyleBuilder buildBorderStyle(BorderStyle borderStyle) {
        style.setLeftBorderStyle(borderStyle);
        style.setRightBorderStyle(borderStyle);
        style.setTopBorderStyle(borderStyle);
        style.setBottomBorderStyle(borderStyle);
        return this;
    }

    public StyleBuilder buildBorderColor(int r, int g, int b) {
        int[] rgbColor = new int[]{r, g, b};
        style.setLeftBorderColor(rgbColor);
        style.setRightBorderColor(rgbColor);
        style.setTopBorderColor(rgbColor);
        style.setBottomBorderColor(rgbColor);
        return this;
    }

    public StyleBuilder buildBorder(BorderStyle borderStyle, int r, int g, int b) {
        style.setLeftBorderStyle(borderStyle);
        style.setRightBorderStyle(borderStyle);
        style.setTopBorderStyle(borderStyle);
        style.setBottomBorderStyle(borderStyle);
        int[] rgbColor = new int[]{r, g, b};
        style.setLeftBorderColor(rgbColor);
        style.setRightBorderColor(rgbColor);
        style.setTopBorderColor(rgbColor);
        style.setBottomBorderColor(rgbColor);
        return this;
    }

    public StyleBuilder buildRowHeight(float rowHeight) {
        style.setRowHeight(rowHeight);
        return this;
    }

    public StyleBuilder buildColumnWidth(int columnWidth) {
        style.setColumnWidth(columnWidth * 256);
        return this;
    }

    public Style build() {
        return style;
    }

}

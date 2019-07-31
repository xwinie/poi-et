package com.jg.poiet.config;

/**
 * 默认模板语法
 */
public enum GramerSymbol {

    /**
     * 文本
     */
    TEXT('\0'),

    /**
     * 表格
     */
    TABLE('#'),

    /**
     * 图片
     */
    Picture('@')
    ;

    private char symbol;

    GramerSymbol(char symbol) {
        this.symbol = symbol;
    }

    public char getSymbol() {
        return this.symbol;
    }

    public String toString() {
        return String.valueOf(this.symbol);
    }

}

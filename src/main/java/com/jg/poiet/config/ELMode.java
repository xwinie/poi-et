package com.jg.poiet.config;

/**
 * 模板表达式的模式
 */
public enum ELMode {

    /**
     * 标准模式：无法计算表达式时，RenderData默认为null值
     */
    POI_TL_STANDARD_MODE,
    /**
     * 严格模式：无法计算表达式直接抛出异常
     */
    POI_TL_STICT_MODE,

}

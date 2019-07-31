package com.jg.poiet.config;

import com.jg.poiet.policy.PictureRenderPolicy;
import com.jg.poiet.policy.RenderPolicy;
import com.jg.poiet.policy.TableRenderPolicy;
import com.jg.poiet.policy.TextRenderPolicy;

import java.util.HashMap;
import java.util.Map;
import java.util.Set;

/**
 * 插件化配置
 */
public class Configure {

    // defalut expression
    private static final String DEFAULT_GRAMER_REGEX = "[\\w\\u4e00-\\u9fa5]+(\\.[\\w\\u4e00-\\u9fa5]+)*";

    // Highest priority
    private Map<String, RenderPolicy> customPolicys = new HashMap<>();
    // Low priority
    private Map<Character, RenderPolicy> defaultPolicys = new HashMap<>();

    /**
     * 语法前缀
     */
    private String gramerPrefix = "{{";

    /**
     * 语法后缀
     */
    private String gramerSuffix = "}}";

    /**
     * 默认支持中文、字母、数字、下划线的正则
     */
    private String grammerRegex = DEFAULT_GRAMER_REGEX;

    /**
     * 模板表达式模式，默认为POI_TL_MODE
     */
    private ELMode elMode = ELMode.POI_TL_STANDARD_MODE;

    /**
     * 渲染数据为null时，是保留还是清空模板标签
     */
    private boolean nullToBlank = true;

    private Configure() {
        plugin(GramerSymbol.TEXT, new TextRenderPolicy());
        plugin(GramerSymbol.TABLE, new TableRenderPolicy());
        plugin(GramerSymbol.Picture, new PictureRenderPolicy());
    }

    /**
     * 获取默认配置
     */
    public static Configure createDefault() {
        return newBuilder().build();
    }

    /**
     * 获取构建器
     */
    public static ConfigureBuilder newBuilder() {
        return new ConfigureBuilder();
    }

    /**
     * 新增语法插件
     *
     * @param c
     *            模板语法
     * @param policy
     *            策略
     */
    public Configure plugin(char c, RenderPolicy policy) {
        defaultPolicys.put(c, policy);
        return this;
    }

    Configure plugin(GramerSymbol symbol, RenderPolicy policy) {
        defaultPolicys.put(symbol.getSymbol(), policy);
        return this;
    }

    /**
     * 自定义模板
     *
     * @param tagName
     *            模板名称
     * @param policy
     *            策略
     */
    public void customPolicy(String tagName, RenderPolicy policy) {
        customPolicys.put(tagName, policy);
    }

    /**
     * 获取标签策略
     *
     * @param tagName
     *            模板名称
     * @param sign
     *            语法
     */
    public RenderPolicy getPolicy(String tagName, Character sign) {
        RenderPolicy policy = getCustomPolicy(tagName);
        return null == policy ? getDefaultPolicy(sign) : policy;
    }

    private RenderPolicy getCustomPolicy(String tagName) {
        return customPolicys.get(tagName);
    }

    private RenderPolicy getDefaultPolicy(Character sign) {
        return defaultPolicys.get(sign);
    }

    public Set<Character> getGramerChars() {
        return defaultPolicys.keySet();
    }

    public ELMode getElMode() {
        return elMode;
    }

    public void setElMode(ELMode elMode) {
        this.elMode = elMode;
    }

    public String getGrammerRegex() {
        return grammerRegex;
    }

    public void setGrammerRegex(String grammerRegex) {
        this.grammerRegex = grammerRegex;
    }

    public String getGramerPrefix() {
        return gramerPrefix;
    }

    public void setGramerPrefix(String gramerPrefix) {
        this.gramerPrefix = gramerPrefix;
    }

    public String getGramerSuffix() {
        return gramerSuffix;
    }

    public void setGramerSuffix(String gramerSuffix) {
        this.gramerSuffix = gramerSuffix;
    }

    public boolean isNullToBlank() {
        return nullToBlank;
    }

    public static class ConfigureBuilder {
        private Configure config = new Configure();

        public ConfigureBuilder() {}

        public ConfigureBuilder buildGramer(String prefix, String suffix) {
            config.gramerPrefix = prefix;
            config.gramerSuffix = suffix;
            return this;
        }

        public ConfigureBuilder supportNullToBlank(boolean nullToBlank) {
            config.nullToBlank = nullToBlank;
            return this;
        }

        public ConfigureBuilder customPolicy(String tagName, RenderPolicy policy) {
            config.customPolicy(tagName, policy);
            return this;
        }

        public ConfigureBuilder addPlugin(char c, RenderPolicy policy) {
            config.plugin(c, policy);
            return this;
        }

        public Configure build() {
            return config;
        }

    }



}

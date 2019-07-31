package com.jg.poiet.template;

public class ElementTemplate {

    protected Character sign;
    protected String tagName;
    protected String source;

    public ElementTemplate() {}

    public Character getSign() {
        return sign;
    }

    public void setSign(Character sign) {
        this.sign = sign;
    }

    public String getTagName() {
        return tagName;
    }

    public void setTagName(String tagName) {
        this.tagName = tagName;
    }

    public String getSource() {
        return source;
    }

    public void setSource(String source) {
        this.source = source;
    }

    @Override
    public String toString() {
        return source;
    }
}

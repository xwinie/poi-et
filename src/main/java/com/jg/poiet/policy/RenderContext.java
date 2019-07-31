package com.jg.poiet.policy;

import com.jg.poiet.XSSFTemplate;
import com.jg.poiet.template.ElementTemplate;

public class RenderContext {

    private ElementTemplate eleTemplate;
    private Object data;
    private XSSFTemplate template;

    public RenderContext(ElementTemplate eleTemplate, Object data, XSSFTemplate template) {
        this.eleTemplate = eleTemplate;
        this.data = data;
        this.template = template;
    }

    public ElementTemplate getEleTemplate() {
        return eleTemplate;
    }

    public void setEleTemplate(ElementTemplate eleTemplate) {
        this.eleTemplate = eleTemplate;
    }

    public Object getData() {
        return data;
    }

    public void setData(Object data) {
        this.data = data;
    }

    public XSSFTemplate getTemplate() {
        return template;
    }

    public void setTemplate(XSSFTemplate template) {
        this.template = template;
    }

}

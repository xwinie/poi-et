package com.jg.poiet.render;

import com.jg.poiet.el.ELObject;
import com.jg.poiet.exception.ExpressionEvalException;

public class ELObjectRenderDataCompute implements RenderDataCompute{

    private ELObject elObject;
    private boolean isStrict;

    public ELObjectRenderDataCompute(Object root, boolean isStrict) {
        elObject = ELObject.create(root);
        this.isStrict = isStrict;
    }

    @Override
    public Object compute(String el) {
        try {
            return elObject.eval(el);
        } catch (ExpressionEvalException e) {
            if (isStrict)
                throw e;
            // mark：无法计算或者读取表达式，默认返回null
            return null;
        }
    }

}

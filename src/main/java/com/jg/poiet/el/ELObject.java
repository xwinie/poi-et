package com.jg.poiet.el;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 对象点缀法访问
 */
public class ELObject {

    final Object model;

    final Map<String, Object> cache = new ConcurrentHashMap<>(32);

    public ELObject(Object model) {
        this.model = model;
    }

    public static ELObject create(Object model) {
        return new ELObject(model);
    }

    public Object eval(String el) {
        Dot dot = new Dot(el);
        return dot.eval(this);
    }

}

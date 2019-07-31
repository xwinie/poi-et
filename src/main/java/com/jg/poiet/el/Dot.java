package com.jg.poiet.el;

import com.jg.poiet.exception.ExpressionEvalException;
import com.jg.poiet.util.ObjectUtils;

import java.lang.reflect.Field;
import java.util.Collection;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * 点缀对象
 */
public class Dot {

    private String el;
    private Dot target;
    private String key;

    // EL通用正则
    final static Pattern EL_PATTERN = Pattern.compile("^[^\\.]+(\\.[^\\.]+)*$");

    public Dot(String el) {
        ObjectUtils.requireNonNull(el, "EL cannot be null.");
        if (!EL_PATTERN.matcher(el)
                .matches()) { throw new ExpressionEvalException("Error EL fomart: " + el); }

        this.el = el;
        int dotIndex = el.lastIndexOf(".");
        if (-1 == dotIndex) {
            this.key = el;
        } else {
            this.key = el.substring(dotIndex + 1);
            this.target = new Dot(el.substring(0, dotIndex));
        }
    }

    public Object eval(ELObject elObject) {
        if (elObject.cache.containsKey(el)) return elObject.cache.get(el);
        Object result = null != target ? evalKey(target.eval(elObject))
                : evalKey(elObject.model);
        if (null != result) elObject.cache.put(el, result);
        return result;
    }

    private Object evalKey(Object obj) {
        ObjectUtils.requireNonNull(obj,
                "Cannot read value from null Prefix-Model, Prefix-Model EL: " + target);
        final Class<?> objClass = obj.getClass();
        if (obj instanceof String || obj instanceof Number || obj instanceof java.util.Date
                || obj instanceof Collection || objClass.isArray()
                || objClass.isPrimitive()) { throw new ExpressionEvalException(
                "Prefix-Model must be JavaBean or Map, Prefix-Model EL: " + target
                        + ", Prefix-Model type: " + objClass); }

        if (obj instanceof Map) { return ((Map<?, ?>) obj).get(key); }

        Field field = FieldFinder.find(objClass, key);
        if (null == field) {
            throw new ExpressionEvalException(
                    "Cannot find the key:" + key + " from Prefix-Model EL:" + target);
        } else {
            try {
                return field.get(obj);
            } catch (Exception e) {
                throw new ExpressionEvalException(
                        "Error read the property:" + key + " from " + objClass);
            }
        }
    }

    public String getEl() {
        return el;
    }

    public Dot getTarget() {
        return target;
    }

    public String getKey() {
        return key;
    }

    @Override
    public String toString() {
        return null == el ? "[root el]" : el;
    }

}

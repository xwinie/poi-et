package com.jg.poiet.policy;

import com.jg.poiet.XSSFTemplate;
import com.jg.poiet.exception.RenderException;
import com.jg.poiet.template.ElementTemplate;
import com.jg.poiet.template.cell.CellTemplate;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.io.IOException;

/**
 * 抽象策略
 */
public abstract class AbstractRenderPolicy<T> implements RenderPolicy {

    /**
     * 校验data model
     *
     * @param data  data
     * @return  result
     */
    protected boolean validate(T data) {
        return true;
    }

    /**
     * 校验失败
     *
     */
    protected void doValidError(RenderContext context) {
        logger.debug("Validate the data of the element {} error, the data may be null or empty: {}",
                context.getEleTemplate().getSource(), context.getData());
        if (context.getTemplate().getConfig().isNullToBlank()) {
            logger.debug("[config.isNullToBlank == true] clear the element {} from the word file.",
                    context.getEleTemplate().getSource());
            clearCell(context);
        } else {
            logger.debug("The element {} Unable to be rendered, nothing to do.",
                    context.getEleTemplate().getSource());
        }
    }

    protected void beforeRender(RenderContext context) {}

    protected void afterRender(RenderContext context) {}

    /**
     * 执行模板渲染
     *
     * @param cellTemplate
     *            文档模板
     * @param data
     *            数据模型
     * @param template
     *            文档对象
     */
    public abstract void doRender(CellTemplate cellTemplate, T data, XSSFTemplate template) throws Exception;

    /*
     * 骨架 (non-Javadoc)
     *
     * @see
     * com.deepoove.poi.policy.RenderPolicy#render(com.deepoove.poi.template.
     * ElementTemplate, java.lang.Object, com.deepoove.poi.XWPFTemplate)
     */
    @SuppressWarnings("unchecked")
    @Override
    public void render(ElementTemplate eleTemplate, Object data, XSSFTemplate template) {
        CellTemplate cellTemplate = (CellTemplate) eleTemplate;

        // type safe
        T model;
        try {
            model = (T) data;
        } catch (ClassCastException e) {
            throw new RenderException(
                    "Error Render Data format for template: " + eleTemplate.getSource(), e);
        }

        // validate
        RenderContext context = new RenderContext(eleTemplate, data, template);
        if (null == model || !validate(model)) {
            doValidError(context);
            return;
        }

        // do render
        try {
            beforeRender(context);
            doRender(cellTemplate, model, template);
            afterRender(context);
        } catch (Exception e) {
            doRenderException(cellTemplate, model, e);
        }

    }

    /**
     * 发生异常
     *
     * @param cellTemplate  cellTemplate
     * @param data  data
     * @param e e
     */
    protected void doRenderException(CellTemplate cellTemplate, T data, Exception e) {
        throw new RenderException("Render template:" + cellTemplate + " error", e);
    }

    /**
     * 继承这个方法，实现自定义清空标签的方案
     *
     * @param context   context
     */
    protected void clearCell(RenderContext context) {
        XSSFCell cell = ((CellTemplate) context.getEleTemplate()).getCell();
        cell.setCellValue("");
    }
}

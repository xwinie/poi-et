package com.jg.poiet.render;

import com.jg.poiet.NiceXSSFWorkbook;
import com.jg.poiet.XSSFTemplate;
import com.jg.poiet.config.Configure;
import com.jg.poiet.exception.RenderException;
import com.jg.poiet.policy.RenderPolicy;
import com.jg.poiet.resolver.TemplateVisitor;
import com.jg.poiet.template.ElementTemplate;
import com.jg.poiet.util.ObjectUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 渲染器，支持表达式计算接口RenderDataCompute的扩展
 */
public class Render {

    private final Logger LOGGER = LoggerFactory.getLogger(Render.class);

    private RenderDataCompute renderDataCompute;

    /**
     * 默认计算器为ELObjectRenderDataCompute
     *
     * @param root  root
     */
    public Render(Object root) {
        ObjectUtils.requireNonNull(root, "Data root is null, should be setted first.");
        renderDataCompute = new ELObjectRenderDataCompute(root, false);
    }

    public Render(RenderDataCompute dataCompute) {
        this.renderDataCompute = dataCompute;
    }

    public void render(XSSFTemplate template) {
        ObjectUtils.requireNonNull(template, "Template is null, should be setted first.");
        LOGGER.info("Render the template file start...");

        NiceXSSFWorkbook workbook = template.getXSSFWorkbook();
        if (null == workbook) return ;
        // 策略
        RenderPolicy policy;

        try {
            TemplateVisitor visitor = new TemplateVisitor(template.getConfig());
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {    //遍历sheet
                XSSFSheet sheet = workbook.getSheetAt(i);
                if (null == sheet) continue;
                for (int j = 0; j <= sheet.getLastRowNum(); j++) {  //遍历行
                    XSSFRow row = sheet.getRow(j);
                    if (null == row) continue;
                    for (int k = 0; k <= row.getLastCellNum(); k++) {   //遍历表格
                        XSSFCell cell = row.getCell(k);
                        ElementTemplate cellTemplates = visitor.visitCell(cell);
                        if (null != cellTemplates) {
                            policy = findPolicy(template.getConfig(), cellTemplates);
                            doRender(cellTemplates, policy, template);
                        }
                    }
                }
            }
        } catch (Exception e) {
            LOGGER.info("Render the template file failed.");
            throw new RenderException("Render xlsx failed.", e);
        }
        LOGGER.info("Render the template file successed.");
    }

    private RenderPolicy findPolicy(Configure config, ElementTemplate cellTemplate) {
        RenderPolicy policy;
        policy = config.getPolicy(cellTemplate.getTagName(), cellTemplate.getSign());
        if (null == policy) { throw new RenderException(
                "Cannot find render policy: [" + cellTemplate.getTagName() + "]"); }
        return policy;
    }

    private void doRender(ElementTemplate ele, RenderPolicy policy, XSSFTemplate template) {
        LOGGER.debug("Start render TemplateName:{}, Sign:{}, policy:{}", ele.getTagName(),
                ele.getSign(), policy.getClass().getSimpleName());
        policy.render(ele, renderDataCompute.compute(ele.getTagName()), template);
    }

}

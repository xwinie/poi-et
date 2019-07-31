package com.jg.poiet.policy;

import com.jg.poiet.XSSFTemplate;
import com.jg.poiet.template.cell.CellTemplate;
import com.jg.poiet.data.TextRenderData;
import com.jg.poiet.util.StyleUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;

public class TextRenderPolicy extends AbstractRenderPolicy<Object> {

    @Override
    public void doRender(CellTemplate cellTemplate, Object renderData, XSSFTemplate template) {
        XSSFCell cell = cellTemplate.getCell();
        Helper.renderTextCell(cell, renderData, template);
    }

    public static class Helper {

        public static void renderTextCell(XSSFCell cell, Object renderData,XSSFTemplate template) {
            if (null == renderData) {
                renderData = new TextRenderData();
            }
            // text
            TextRenderData textRenderData = renderData instanceof TextRenderData
                    ? (TextRenderData) renderData
                    : new TextRenderData(renderData.toString());

            String data = null == textRenderData.getText() ? "" : textRenderData.getText();

            StyleUtils.styleCell(template, cell, textRenderData.getStyle());

            cell.setCellValue(data);
        }

    }

}

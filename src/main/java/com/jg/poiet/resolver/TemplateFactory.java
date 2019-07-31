package com.jg.poiet.resolver;

import com.jg.poiet.config.Configure;
import com.jg.poiet.template.cell.CellTemplate;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.util.Set;

public class TemplateFactory {

    public static final char EMPTY_CHAR = '\0';

    public static CellTemplate createCellTemplate(String tag, Configure config, XSSFCell cell) {
        CellTemplate template = new CellTemplate();
        Set<Character> gramerChars = config.getGramerChars();
        char fisrtChar = tag.charAt(0);
        Character symbol = EMPTY_CHAR;
        for (Character chara : gramerChars) {
            if (chara.equals(fisrtChar)) {
                symbol = fisrtChar;
                break;
            }
        }
        template.setSource(config.getGramerPrefix() + tag + config.getGramerSuffix());
        template.setTagName(symbol.equals(EMPTY_CHAR) ? tag : tag.substring(1));
        template.setSign(symbol);
        template.setCell(cell);
        return template;
    }

}

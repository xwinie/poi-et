package com.jg.poiet.resolver;

import com.jg.poiet.config.Configure;
import com.jg.poiet.template.ElementTemplate;
import com.jg.poiet.util.CellUtils;
import com.jg.poiet.util.RegexUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

/**
 * 模板解析器
 */
public class TemplateVisitor implements Visitor {

    private static Logger logger = LoggerFactory.getLogger(TemplateVisitor.class);

    private Configure config;
    private List<ElementTemplate> eleTemplates;

    private Pattern templatePattern;
    private Pattern gramerPattern;

    static final String FORMAT_TEMPLATE = "{0}{1}{2}{3}";
    static final String FORMAT_GRAMER = "({0})|({1})";

    public TemplateVisitor(Configure config) {
        this.config = config;
        initPattern();
    }

    @Override
    public List<ElementTemplate> visitExcel(XSSFWorkbook workbook) {
        if (null == workbook) return null;
        this.eleTemplates = new ArrayList<>();
        logger.info("Visit the document start...");
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            XSSFSheet sheet = workbook.getSheetAt(i);
            visitSheet(sheet);
        }
        logger.info("Visit the document end, resolve and create {} ElementTemplates.",
                this.eleTemplates.size());
        return eleTemplates;
    }

    @Override
    public ElementTemplate visitCell(XSSFCell cell) {
        if (null == cell) return null;
        String text = CellUtils.getCellValue(cell);
        if (StringUtils.isBlank(text)) return null;
        return parseTemplateFactory(text, cell);
    }

    void visitSheet(XSSFSheet sheet) {
        if (null == sheet) return;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            XSSFRow row = sheet.getRow(i);
            visitRow(row);
        }
    }

    void visitRow(XSSFRow row) {
        if (null == row) return;
        for (int i = 0; i <= row.getLastCellNum(); i++) {
            XSSFCell cell = row.getCell(i);
            if (null == cell) continue;
            String text = CellUtils.getCellValue(cell);
            if (StringUtils.isBlank(text)) continue;
            ElementTemplate elementTemplate = parseTemplateFactory(text, cell);
            if (null != elementTemplate) eleTemplates.add(elementTemplate);
        }
    }

    private <T> ElementTemplate parseTemplateFactory(String text, T obj) {
        logger.debug("Resolve text: {}, and create ElementTemplate", text);
        // temp ,future need to word analyze
        if (templatePattern.matcher(text).matches()) {
            String tag = gramerPattern.matcher(text).replaceAll("").trim();
            if (obj.getClass() == XSSFCell.class) {
                return TemplateFactory.createCellTemplate(tag, config, (XSSFCell) obj);
            }
        }
        return null;
    }

    private void initPattern() {
        String signRegex = getGramarRegex(config);
        String prefixRegex = RegexUtils.escapeExprSpecialWord(config.getGramerPrefix());
        String suffixRegex = RegexUtils.escapeExprSpecialWord(config.getGramerSuffix());

        templatePattern = Pattern.compile(MessageFormat.format(FORMAT_TEMPLATE, prefixRegex,
                signRegex, config.getGrammerRegex(), suffixRegex));
        gramerPattern = Pattern
                .compile(MessageFormat.format(FORMAT_GRAMER, prefixRegex, suffixRegex));
    }

    private String getGramarRegex(Configure config) {
        List<Character> gramerChar = new ArrayList<>(config.getGramerChars());
        StringBuilder reg = new StringBuilder("(");
        for (int i = 0;; i++) {
            Character chara = gramerChar.get(i);
            String escapeExprSpecialWord = RegexUtils.escapeExprSpecialWord(chara.toString());
            if (i == gramerChar.size() - 1) {
                reg.append(escapeExprSpecialWord).append(")?");
                break;
            } else reg.append(escapeExprSpecialWord).append("|");
        }
        return reg.toString();
    }


}

package com.jg.poiet.resolver;

import com.jg.poiet.template.ElementTemplate;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;

public interface Visitor {

    List<ElementTemplate> visitExcel(XSSFWorkbook workbook);

    ElementTemplate visitCell(XSSFCell cell);

}

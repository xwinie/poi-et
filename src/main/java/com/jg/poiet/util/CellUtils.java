package com.jg.poiet.util;

import org.apache.poi.xssf.usermodel.XSSFCell;

import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

public class CellUtils {
    public static String getCellValue(XSSFCell cell) {
        String cellValue = "";
        if (null != cell) {
            // 以下是判断数据的类型
            switch (cell.getCellType()) {
                case STRING: // 字符串
                    cellValue = cell.getStringCellValue();
                    break;
                case NUMERIC: // 数字
                    if(org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)){
                        Date d = cell.getDateCellValue();
                        DateFormat formater = new SimpleDateFormat("yyyy-MM-dd");
                        cellValue = formater.format(d);
                    } else {
                        DecimalFormat df = new DecimalFormat("0");

                        cellValue = df.format(cell.getNumericCellValue());
                    }
                    break;
                case BOOLEAN: // Boolean
                    cellValue = cell.getBooleanCellValue() + "";
                    break;
                case FORMULA: // 公式
                    cellValue = cell.getCellFormula() + "";
                    break;
                case BLANK: // 空值
                    cellValue = "";
                    break;
                case ERROR: // 故障
                    cellValue = "非法字符";
                    break;
                default:
                    cellValue = "未知类型";
                    break;
            }
        }
        return cellValue;
    }
}

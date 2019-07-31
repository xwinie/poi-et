package com.jg.poiet.util;

import java.text.SimpleDateFormat;
import java.util.Date;

public class DateUtils {

    /**
     * 获取当前日期
     * @param format    日期格式
     * @return  当前日期
     */
    public static String getNowDate(String format) {
        SimpleDateFormat sdf = new SimpleDateFormat(format);
        return sdf.format(new Date());
    }

    /**
     * 输出当前日期
     */
    public static void printNowDate(String note) {
        System.out.println(note+":"+getNowDate("yyyy-MM-dd HH:mm:ss SSS"));
    }

}

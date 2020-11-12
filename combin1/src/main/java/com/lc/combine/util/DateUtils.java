package com.lc.combine.util;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @Author chenglezheng
 * @Date 2020/11/12 10:14
 */
public class DateUtils {


    /**
     * 加上天数，获取日期串
     * @param day
     * @return
     * @throws ParseException
     */
    public static String handleDate(long day) throws ParseException {
        day=day-2;
        SimpleDateFormat dateFormat=new SimpleDateFormat("yyyy/MM/dd");
        Date date = dateFormat.parse("1900/01/01");
        long time = date.getTime();
        day = day*24*60*60*1000;
        time+=day;
        return dateFormat.format(new Date(time));
    }
}

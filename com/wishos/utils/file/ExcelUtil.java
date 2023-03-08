package com.wishos.utils.file;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.read.listener.PageReadListener;

import java.util.ArrayList;
import java.util.List;

/**
 * Excel工具类
 * @author jimu
 */
public class ExcelUtil {
    /**
     * 批量读取Excel，在保证内存的情况下使用，默认会去掉标题行
     *
     * @param filePath Excel 文件地址
     * @param tClass   实体类，需要保证类属性定义与Excel保持一致
     * @param <T>      T
     * @return 列表
     */
    public static <T> List<T> readExcel(String filePath, Class<T> tClass) {
        List<T> list = new ArrayList<>();
        EasyExcel.read(filePath, tClass, new PageReadListener<T>(list::addAll)).sheet().doRead();
        return list;
    }
}

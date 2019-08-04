package com.github.fanlychie.excelutils.annotation;

import lombok.Data;
import com.github.fanlychie.excelutils.spec.Align;

/**
 * 单元格注解字段
 * Created by fanlychie on 2017/3/5.
 */
@Data
public class CellField {

    /**
     * 单元格索引
     *
     * @return
     */
    private int index;

    /**
     * 单元格标题
     *
     * @return
     */
    private String name;

    /**
     * 数据格式
     *
     * @return
     */
    private String format;

    /**
     * 对齐方式
     *
     * @return
     */
    private Align align;

    /**
     * 字段名称
     */
    private String field;

    /**
     * 字段类型
     */
    private Class<?> type;

}
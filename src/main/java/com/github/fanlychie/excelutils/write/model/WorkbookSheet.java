package com.github.fanlychie.excelutils.write.model;

import lombok.Data;

import java.util.Map;

/**
 * 工作表(Sheet)
 *
 * @author fanlychie
 */
@Data
public class WorkbookSheet {

    /**
     * 工作表名称
     */
    private String name = "Sheet";

    /**
     * 单元格宽度
     */
    private Integer cellWidth;

    /**
     * 标题行样式
     */
    private RowStyle titleStyle;

    /**
     * 主体行样式
     */
    private RowStyle bodyStyle;

    /**
     * 脚部行样式
     */
    private RowStyle footerStyle;

    /**
     * 值映射, POJO字段的值写出到EXCEL时, 可以匹配这些值并替换掉它们
     */
    private Map<Object, Object> mapping;

}
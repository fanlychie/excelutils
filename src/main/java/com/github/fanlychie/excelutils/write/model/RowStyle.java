package com.github.fanlychie.excelutils.write.model;

import lombok.Data;

/**
 * 行样式
 *
 * @author fanlychie
 */
@Data
public class RowStyle {

    /**
     * 行的索引值
     */
    private Integer index;

    /**
     * 行的高度
     */
    private Integer height;

    /**
     * 水平方向对齐方式(参考{@link org.apache.poi.ss.usermodel.CellStyle})
     */
    private Short align;

    /**
     * 垂直方向对齐方式(参考{@link org.apache.poi.ss.usermodel.CellStyle})
     */
    private Short verticalAlign;

    /**
     * 是否自动换行
     */
    private Boolean autoWrap;

    /**
     * 字体名称
     */
    private String fontName;

    /**
     * 字体大小
     */
    private Integer fontSize;

    /**
     * 字体颜色(参考{@link org.apache.poi.ss.usermodel.IndexedColors})
     */
    private Short fontColor;

    /**
     * 边框(参考{@link org.apache.poi.ss.usermodel.CellStyle})
     */
    private Short border;

    /**
     * 边框颜色(参考{@link org.apache.poi.ss.usermodel.IndexedColors})
     */
    private Short borderColor;

    /**
     * 背景颜色(参考{@link org.apache.poi.ss.usermodel.IndexedColors})
     */
    private Short backgroundColor;

    /**
     * 行数据格式
     */
    private String format;

}
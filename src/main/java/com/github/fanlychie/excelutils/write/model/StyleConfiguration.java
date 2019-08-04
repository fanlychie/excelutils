package com.github.fanlychie.excelutils.write.model;

import lombok.Data;

import java.util.Map;

/**
 * 样式配置, 用于接收YAML配置文件的配置项
 *
 * @author fanlychie
 */
@Data
public class StyleConfiguration {

    /**
     * 全局配置
     */
    private GlobalStyleConfiguration global;

    /**
     * 标题行配置
     */
    private RowStyleConfiguration titleStyle;

    /**
     * 主体行配置
     */
    private RowStyleConfiguration bodyStyle;

    @Data
    public static class GlobalStyleConfiguration {
        /**
         * 单元格宽度
         */
        private Integer cellWidth;
    }

    @Data
    public static class RowStyleConfiguration {
        /**
         * 起始索引, 从0开始
         */
        private Integer index;
        /**
         * 行高
         */
        private Integer height;
        /**
         * 字体名称
         */
        private String fontName;
        /**
         * 字体大小
         */
        private Integer fontSize;
        /**
         * 字体颜色
         */
        private String fontColor;
        /**
         * 是否自动换行
         */
        private Boolean autoWrap;
        /**
         * 背景颜色
         */
        private String backgroundColor;
        /**
         * 水平方向对齐方式
         */
        private String align;
        /**
         * 垂直方向对齐方式
         */
        private String verticalAlign;
        /**
         * 单元格格式
         */
        private String format;
        /**
         * 边框
         */
        private Short border;
        /**
         * 边框颜色
         */
        private String borderColor;
        /**
         * 关键字映射
         */
        private Map<Object, Object> mapping;
    }

}
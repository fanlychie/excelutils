package com.github.fanlychie.excelutils.annotation;

import com.github.fanlychie.excelutils.spec.Align;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 注解, 用于标注单元格数据字段
 * Created by fanlychie on 2017/3/4.
 */
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Cell {

    /**
     * 单元格索引, 从左至右数, 数值从0开始
     *
     * @return
     */
    int index();

    /**
     * 单元格标题名称, 用于写出EXCEL文件时的标题行文字
     *
     * @return
     */
    String name();

    /**
     * 数据格式{@link com.github.fanlychie.excelutils.spec.Format#format}, 默认为字符串格式
     *
     * @return
     */
    String format() default "";

    /**
     * 对齐方式{@link Align}
     *
     * @return
     */
    Align align() default Align.LEFT;

}
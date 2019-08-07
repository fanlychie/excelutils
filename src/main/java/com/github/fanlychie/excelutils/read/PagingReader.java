package com.github.fanlychie.excelutils.read;

import java.util.List;

/**
 * 用于分页数据处理。当解析读取到的数据达到pageSize设定的阀值时, 触发PagingReader处理
 */
public interface PagingReader<T> {

    void read(List<T> items);

}
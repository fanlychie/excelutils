package com.github.fanlychie.excelutils.read;

import java.util.List;

/**
 * 用于处理分页的数据。当读取到的数据达到ExcelReaderBuilder.pageSize设定的阀值时, 触发PagingHandler.handle处理数据
 */
public interface PagingHandler<T> {

    /**
     * 解析EXCEL文件数据完成, 回调该函数来处理数据
     *
     * @param items 分页中的一页数据
     */
    void handle(List<T> items);

}
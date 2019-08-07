package com.github.fanlychie.excelutils.write;

import java.util.List;

/**
 * 分页查询接口
 *
 * @author fanlychie
 */
public interface PagingQuerier {

    /**
     * 查询一页的数据
     *
     * @param page   页码
     * @param offset 起始索引
     * @param size   每页的大小
     * @return 返回查询的结果集
     */
    List queryPage(int page, int offset, int size);

}
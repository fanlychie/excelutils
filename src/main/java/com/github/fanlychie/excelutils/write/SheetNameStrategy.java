package com.github.fanlychie.excelutils.write;

/**
 * Sheet页名称策略
 */
public interface SheetNameStrategy {

    /**
     * 获取Sheet页的名称
     *
     * @param sheetIndex 当前Sheet页的索引, 第一个Sheet页的索引为1, 以此类推
     * @return 返回Sheet页的名称, 用于构建新的Sheet页
     */
    String getSheetName(int sheetIndex);

}
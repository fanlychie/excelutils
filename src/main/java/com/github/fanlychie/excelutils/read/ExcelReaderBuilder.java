package com.github.fanlychie.excelutils.read;

import com.github.fanlychie.excelutils.exception.ExcelCastException;
import com.github.fanlychie.excelutils.read.ExcelReader.Paging;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;

/**
 * EXCEL读取器的构建工具, 用于构建一个{@link ExcelReader}实例
 *
 * @author fanlychie
 */
public final class ExcelReaderBuilder {

    /**
     * EXCEL文件流
     */
    private InputStream excelStream;

    /**
     * EXCEL文件每一个映射到的POJO类
     */
    private Class<?> pojoClass;

    /**
     * 正文从第几行开始解析
     */
    private int rownum;

    /**
     * 分页
     */
    private Paging paging;

    /**
     * 数据处理
     */
    private PagingReader reader;

    /**
     * 配置EXCEL文件流
     *
     * @param in EXCEL文件流
     * @return 返回 {@link ExcelReaderBuilder}
     */
    public ExcelReaderBuilder stream(InputStream in) {
        this.excelStream = in;
        return this;
    }

    /**
     * 配置要解析的EXCEL文件
     *
     * @param file EXCEL文件
     * @return 返回 {@link ExcelReaderBuilder}
     */
    public ExcelReaderBuilder stream(File file) {
        try {
            return stream(new FileInputStream(file));
        } catch (FileNotFoundException e) {
            throw new ExcelCastException(e);
        }
    }

    /**
     * 配置要解析的EXCEL文件路径
     *
     * @param pathname EXCEL文件路径
     * @return 返回 {@link ExcelReaderBuilder}
     */
    public ExcelReaderBuilder stream(String pathname) {
        return stream(new File(pathname));
    }

    /**
     * 配置EXCEL文件数据行映射的POJO类型
     *
     * @param clazz EXCEL文件数据行映射的POJO类型
     * @return 返回 {@link ExcelReaderBuilder}
     */
    public ExcelReaderBuilder payload(Class<?> clazz) {
        this.pojoClass = clazz;
        return this;
    }

    /**
     * 分页大小
     *
     * @param size 每一页的数据大小
     * @return 返回 {@link ExcelReaderBuilder}
     */
    public ExcelReaderBuilder pageSize(int size) {
        paging = new Paging();
        paging.size = size;
        return this;
    }

    /**
     * 当解析读取到的数据达到pageSize设定的阀值时, 触发PagingReader处理
     *
     * @param reader {@link PagingReader}
     * @return 返回 {@link ExcelReaderBuilder}
     */
    public ExcelReaderBuilder reader(PagingReader reader) {
        this.reader = reader;
        return this;
    }

    /**
     * 配置从文件的第几行开始解析
     *
     * @param row 从文件的第几行开始解析
     * @return 返回 {@link ExcelReaderBuilder}
     */
    public ExcelReaderBuilder start(int row) {
        this.rownum = row;
        return this;
    }

    /**
     * 构建 {@link ExcelReader} 实例
     *
     * @return 返回 {@link ExcelReader}
     */
    public ExcelReader build() {
        ExcelReader excelReader = new ExcelReader();
        excelReader.setStream(excelStream);
        excelReader.setStart(rownum);
        excelReader.setTargetClass(pojoClass);
        excelReader.setPaging(paging);
        excelReader.setReader(reader);
        excelReader.init();
        return excelReader;
    }

}
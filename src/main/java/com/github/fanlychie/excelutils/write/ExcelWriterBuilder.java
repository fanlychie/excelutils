package com.github.fanlychie.excelutils.write;

import com.github.fanlychie.excelutils.exception.WriteExcelException;
import com.github.fanlychie.excelutils.spec.Align;
import com.github.fanlychie.excelutils.spec.Format;
import com.github.fanlychie.excelutils.write.ExcelWriter.Paging;
import com.github.fanlychie.excelutils.write.model.StyleConfiguration;

/**
 * EXCEL写操作的构建工具, 用于构建一个{@link ExcelWriter}实例
 * *
 * * @author fanlychie
 */
public final class ExcelWriterBuilder {

    /**
     * POJO 类
     */
    private Class<?> pojoClass;

    /**
     * 工作表样式
     */
    private StyleConfiguration config;

    /**
     * 工作表配置
     */
    private Configurable configSheet;

    /**
     * 分页参数
     */
    private Paging paging;

    /**
     * 分页查询接口
     */
    private PagingQuerier querier;

    /**
     * Sheet页名称策略
     */
    private SheetNameStrategy strategy;

    /**
     * 使用YAML配置文件配置样式
     *
     * @param pathname 配置文件路径
     * @return 返回 {@link ExcelWriterBuilder}
     */
    public ExcelWriterBuilder configure(String pathname) {
        this.configSheet = new Configurable(pathname);
        return this;
    }

    /**
     * 使用内置的样式
     *
     * @return 返回 {@link ExcelWriterBuilder}
     */
    public ExcelWriterBuilder builtin() {
        return configure("jexcel-default.yml");
    }

    /**
     * 使用编码自定义配置样式
     *
     * @return 返回 {@link ExcelWriterBuilder}
     */
    public ExcelWriterBuilder define() {
        builtin();
        this.config = configSheet.config;
        return this;
    }

    /**
     * 数据载体, POJO 类
     *
     * @param pojoClass POJO 类
     * @return 返回 {@link ExcelWriterBuilder}
     */
    public ExcelWriterBuilder payload(Class<?> pojoClass) {
        this.pojoClass = pojoClass;
        return this;
    }

    /**
     * 配置标题行样式
     *
     * @return 返回 {@link ExcelWriterBuilder}
     */
    public TitleRowStyleBuilder title() {
        return new TitleRowStyleBuilder(this);
    }

    /**
     * 配置主体行样式
     *
     * @return 返回 {@link ExcelWriterBuilder}
     */
    public BodyRowStyleBuilder body() {
        return new BodyRowStyleBuilder(this);
    }

    /**
     * 分页查询
     *
     * @param querier 分页查询实现
     * @return 返回 {@link PagingBuilder}
     */
    public PagingBuilder pagingQuery(PagingQuerier querier) {
        this.querier = querier;
        return new PagingBuilder(this);
    }

    /**
     * 构建{@link ExcelWriter}实例, 用于输出EXCEL文件
     *
     * @return 返回 {@link ExcelWriter}
     */
    public ExcelWriter build() {
        if (configSheet == null) {
            throw new WriteExcelException("Configurable can not be null");
        }
        if (pojoClass == null) {
            throw new WriteExcelException("payload can not be null");
        }
        if (config != null) {
            return new ExcelWriter().prepare(configSheet.buildWorkbookSheet(config), pojoClass, paging, querier, strategy);
        }
        return new ExcelWriter().prepare(configSheet.buildWorkbookSheet(), pojoClass, paging, querier, strategy);
    }

    public static class BodyRowStyleBuilder extends BasicRowStyleBuilder<BodyRowStyleBuilder> {

        protected BodyRowStyleBuilder(ExcelWriterBuilder builder) {
            super(builder);
            init(this, builder.config.getBodyStyle());
        }

        /**
         * 值映射, POJO字段的值写出到EXCEL时, 可以匹配这些值并替换掉它们
         *
         * @param value POJO字段值
         * @param fake  写出到EXCEL的值
         * @return 返回当前引用
         */
        public BodyRowStyleBuilder map(String value, String fake) {
            builder.config.getBodyStyle().getMapping().put(value, fake);
            return this;
        }

        /**
         * 配置完成, 返回上层
         *
         * @return 返回 {@link ExcelWriterBuilder}
         */
        public ExcelWriterBuilder and() {
            return builder;
        }

    }

    public static class TitleRowStyleBuilder extends BasicRowStyleBuilder<TitleRowStyleBuilder> {

        protected TitleRowStyleBuilder(ExcelWriterBuilder builder) {
            super(builder);
            init(this, builder.config.getTitleStyle());
        }

        /**
         * 配置完成, 返回上层
         *
         * @return 返回 {@link ExcelWriterBuilder}
         */
        public ExcelWriterBuilder and() {
            return builder;
        }

    }

    protected static class BasicRowStyleBuilder<T> {

        private T reference;

        private StyleConfiguration.RowStyleConfiguration style;

        protected ExcelWriterBuilder builder;

        protected BasicRowStyleBuilder(ExcelWriterBuilder builder) {
            this.builder = builder;
        }

        protected void init(T reference, StyleConfiguration.RowStyleConfiguration style) {
            this.style = style;
            this.reference = reference;
        }

        /**
         * 设置行的起始索引, 第一行的索引值为0, 以此类推
         *
         * @param index 索引值
         * @return 返回当前引用
         */
        public T index(Integer index) {
            style.setIndex(index);
            return reference;
        }

        /**
         * 设置行高
         *
         * @param height 行高
         * @return 返回当前引用
         */
        public T height(Integer height) {
            style.setHeight(height);
            return reference;
        }

        /**
         * 设置水平对齐方式(参考{@link Align})
         *
         * @param alignment 水平对齐方式
         * @return 返回当前引用
         */
        public T align(String alignment) {
            style.setAlign(alignment);
            return reference;
        }

        /**
         * 设置垂直对齐方式(参考{@link Align})
         *
         * @param alignment 垂直对齐方式
         * @return 返回当前引用
         */
        public T vertical(String alignment) {
            style.setVerticalAlign(alignment);
            return reference;
        }

        /**
         * 设置是否自动换行(当文字内容超出单元格宽度时)
         *
         * @param auto 是否自动换行
         * @return 返回当前引用
         */
        public T autoWrap(Boolean auto) {
            style.setAutoWrap(auto);
            return reference;
        }

        /**
         * 设置字体名称
         *
         * @param name 字体名称
         * @return 返回当前引用
         */
        public T fontName(String name) {
            style.setFontName(name);
            return reference;
        }

        /**
         * 设置字体大小
         *
         * @param size 字体大小
         * @return 返回当前引用
         */
        public T fontSize(Integer size) {
            style.setFontSize(size);
            return reference;
        }

        /**
         * 设置字体颜色(参考{@link org.apache.poi.ss.usermodel.IndexedColors})
         *
         * @param color 字体颜色
         * @return 返回当前引用
         */
        public T fontColor(String color) {
            style.setFontColor(color);
            return reference;
        }

        /**
         * 设置边框(参考{@link org.apache.poi.ss.usermodel.CellStyle})
         *
         * @param border 边框
         * @return 返回当前引用
         */
        public T border(Short border) {
            style.setBorder(border);
            return reference;
        }

        /**
         * 设置边框颜色(参考{@link org.apache.poi.ss.usermodel.IndexedColors})
         *
         * @param color 边框颜色
         * @return 返回当前引用
         */
        public T borderColor(String color) {
            style.setBorderColor(color);
            return reference;
        }

        /**
         * 设置背景颜色(参考{@link org.apache.poi.ss.usermodel.IndexedColors})
         *
         * @param color 背景颜色
         * @return 返回当前引用
         */
        public T background(String color) {
            style.setBackgroundColor(color);
            return reference;
        }

        /**
         * 设置单元格格式(参考{@link Format})
         *
         * @param format 单元格格式
         * @return 返回当前引用
         */
        public T format(String format) {
            style.setFormat(format);
            return reference;
        }

    }

    public static class PagingBuilder {

        protected ExcelWriterBuilder builder;

        protected PagingBuilder(ExcelWriterBuilder builder) {
            this.builder = builder;
            builder.paging = new Paging();
        }

        /**
         * 设置分页起始的页码, 从1开始
         *
         * @param page 页码
         * @return 返回当前引用
         */
        public PagingBuilder pageNumber(int page) {
            builder.paging.page = page;
            return this;
        }

        /**
         * 设置每页的大小
         *
         * @param size 每页的大小
         * @return 返回当前引用
         */
        public PagingBuilder pageSize(int size) {
            builder.paging.size = size;
            return this;
        }

        /**
         * 设置每个Sheet页最大的数据行数, 超出这个阀值将自动另起一个新的Sheet页
         *
         * @param max 最大的行数
         * @return 返回当前引用
         */
        public PagingBuilder maxRowsPerSheet(int max) {
            builder.paging.max = max;
            return this;
        }

        /**
         * 设置Sheet页名称策略
         *
         * @param sheetNameStrategy Sheet页名称策略
         * @return 返回当前引用
         */
        public PagingBuilder sheetNameStrategy(SheetNameStrategy sheetNameStrategy) {
            builder.strategy = sheetNameStrategy;
            return this;
        }

        /**
         * 配置完成, 返回上层
         *
         * @return 返回 {@link ExcelWriterBuilder}
         */
        public ExcelWriterBuilder and() {
            return builder;
        }

    }

}
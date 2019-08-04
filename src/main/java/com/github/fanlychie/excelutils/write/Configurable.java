package com.github.fanlychie.excelutils.write;

import com.github.fanlychie.excelutils.write.model.RowStyle;
import com.github.fanlychie.excelutils.write.model.StyleConfiguration;
import com.github.fanlychie.excelutils.write.model.WorkbookSheet;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import com.github.fanlychie.excelutils.spec.Align;
import com.github.fanlychie.excelutils.spec.Format;
import org.yaml.snakeyaml.Yaml;

/**
 * 可配置样式的EXCEL工作表, 根据YAML配置文件的配置项构建表格样式
 *
 * @author fanlychie
 */
public class Configurable {

    /**
     * 样式配置
     */
    StyleConfiguration config;

    /**
     * 构建{@link Configurable}实例
     *
     * @param pathname YAML配置文件路径
     */
    public Configurable(String pathname) {
        this.config = new Yaml().loadAs(getClass().getResourceAsStream("/" + pathname),
                StyleConfiguration.class);
    }

    /**
     * 构建{@link WorkbookSheet}实例
     *
     * @return 返回 {@link WorkbookSheet} 实例
     */
    public WorkbookSheet buildWorkbookSheet() {
        return buildWorkbookSheet(config);
    }

    WorkbookSheet buildWorkbookSheet(StyleConfiguration config) {
        // 工作表
        WorkbookSheet sheet = new WorkbookSheet();
        // 单元格宽度
        sheet.setCellWidth(config.getGlobal().getCellWidth());
        // 标题行样式
        sheet.setTitleStyle(buildRowStyle(config.getTitleStyle()));
        // 主体行样式
        sheet.setBodyStyle(buildRowStyle(config.getBodyStyle()));
        // 关键字映射表
        sheet.setMapping(config.getBodyStyle().getMapping());
        return sheet;
    }

    /**
     * 构建单元格的行样式
     *
     * @param config 行样式配置
     * @return 返回 {@link RowStyle} 实例
     */
    private RowStyle buildRowStyle(StyleConfiguration.RowStyleConfiguration config) {
        RowStyle style = new RowStyle();
        if (config != null) {
            // 起始索引
            style.setIndex(config.getIndex());
            // 背景颜色
            style.setBackgroundColor(config.getBackgroundColor() == null ? null :
                    IndexedColors.valueOf(config.getBackgroundColor().toUpperCase()).index);
            // 字体样式
            style.setFontName(config.getFontName());
            style.setFontSize(config.getFontSize());
            style.setFontColor(config.getFontColor() == null ? null :
                    IndexedColors.valueOf(config.getFontColor().toUpperCase()).index);
            // 高度
            style.setHeight(config.getHeight());
            // 水平方向对齐方式
            style.setAlign(config.getAlign() == null ? null :
                    Align.valueOf(config.getAlign().toUpperCase()).getValue());
            // 垂直方向对齐方式
            style.setVerticalAlign(config.getVerticalAlign() == null ? null :
                    Align.valueOf("VERTICAL_" + config.getVerticalAlign().toUpperCase()).getValue());
            // 是否自动换行
            style.setAutoWrap(config.getAutoWrap());
            // 数据格式
            style.setFormat(config.getFormat() == null ? null :
                    Format.valueOf(config.getFormat().toUpperCase()).getFormat());
            // 边框样式
            style.setBorder(config.getBorder() == null ? CellStyle.BORDER_THIN : config.getBorder());
            style.setBorderColor(config.getBorderColor() == null ? IndexedColors.GREY_25_PERCENT.index :
                    IndexedColors.valueOf(config.getBorderColor().toUpperCase()).index);
        }
        return style;
    }

}
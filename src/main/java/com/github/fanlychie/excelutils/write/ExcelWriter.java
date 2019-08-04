package com.github.fanlychie.excelutils.write;

import com.github.fanlychie.beanutils.BeanUtils;
import com.github.fanlychie.beanutils.operator.FieldOperator;
import com.github.fanlychie.excelutils.annotation.AnnotationHandler;
import com.github.fanlychie.excelutils.annotation.CellField;
import com.github.fanlychie.excelutils.exception.ExcelCastException;
import com.github.fanlychie.excelutils.write.model.RowStyle;
import com.github.fanlychie.excelutils.write.model.WorkbookSheet;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

/**
 * EXCEL写操作, 用于将POJO数据写出到EXCEL文件
 *
 * @author fanlychie
 */
public class ExcelWriter {

    /**
     * SXSSF 工作表
     */
    private SXSSFSheet sheet;

    /**
     * SXSSF 工作薄
     */
    private SXSSFWorkbook workbook;

    /**
     * 工作表
     */
    private WorkbookSheet workbookSheet;

    /**
     * 单元格注解字段列表
     */
    private List<CellField> cellFields;

    /**
     * 工作表计数
     */
    private int sheetCount = 1;

    /**
     * 行的边界索引值
     */
    private int rowBoundIndex;

    ExcelWriter() {}

    /**
     * 写出到一个工作表(Sheet)
     *
     * @param data 数据列表
     * @return 返回当前对象
     */
    public ExcelWriter write(List<?> data) {
        return write(workbookSheet.getName() + (sheetCount++), data);
    }

    /**
     * 写出到一个工作表(Sheet)
     *
     * @param sheetName 工作表名称
     * @param data      数据列表
     * @return 返回当前对象
     */
    public ExcelWriter write(String sheetName, List<?> data) {
        return buildSheet(sheetName, data, true);
    }

    /**
     * 追加到当前的工作表(Sheet), 只会追加数据不会构建标题行
     *
     * @param data 数据列表
     * @return 返回当前对象
     */
    public ExcelWriter append(List<?> data) {
        return buildSheet(null, data, false);
    }

    /**
     * 输出到文件
     *
     * @param pathname 文件路径名称
     */
    public void toFile(String pathname) {
        try {
            workbook.write(new FileOutputStream(new File(pathname)));
        } catch (Throwable e) {
            throw new ExcelCastException(e);
        }
    }

    /**
     * 输出到文件
     *
     * @param file 文件对象
     */
    public void toFile(File file) {
        OutputStream os = null;
        try {
            workbook.write(os);
        } catch (Throwable e) {
            throw new ExcelCastException(e);
        } finally {
            if (os != null) {
                try {
                    os.close();
                } catch (IOException e) {
                }
            }
        }
    }

    /**
     * 输出到输出流
     *
     * @param os 输出流
     */
    public void toStream(OutputStream os) {
        try {
            workbook.write(os);
        } catch (Throwable e) {
            throw new ExcelCastException(e);
        }
    }

    /**
     * 写出到客户端响应, 用于供客户端下载文件
     *
     * @param response HttpServletResponse
     * @param filename 下载时存储的文件名称
     */
    public void toHttp(HttpServletResponse response, String filename) {
        try {
            filename = new String(filename.getBytes("UTF-8"), "ISO-8859-1");
        } catch (UnsupportedEncodingException e) {
            throw new ExcelCastException(e);
        }
        response.reset();
        response.setHeader("Content-Disposition", "attachment; filename=" + filename);
        response.setContentType("application/octet-stream; charset=ISO-8859-1");
        try {
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            throw new ExcelCastException(e);
        }
    }

    ExcelWriter prepare(WorkbookSheet workbookSheet, Class<?> pojoClass) {
        this.workbookSheet = workbookSheet;
        this.workbook = new SXSSFWorkbook();
        this.cellFields = new LinkedList<>(AnnotationHandler.getCellFieldMapping(pojoClass).values());
        return this;
    }

    /**
     * 构建工作表(Sheet)
     *
     * @param sheetName 工作表的名称
     * @param data      填充的数据
     * @param create    是否创建新的工作表
     * @return 返回 {@link ExcelWriter}
     */
    private ExcelWriter buildSheet(String sheetName, List<?> data, boolean create) {
        try {
            if (create) {
                // 创建新的工作表
                createNewSheet(sheetName);
            }
            // 处理工作表数据
            if (data != null && data.size() > 0) {
                // 主体索引
                int bodyIndex = rowBoundIndex == 0 ? workbookSheet.getBodyStyle().getIndex() : rowBoundIndex;
                // 边界索引
                rowBoundIndex += data.size();
                for (Object item : data) {
                    // 构建行数据
                    fillDataRow(bodyIndex++, item);
                }
            }
            return this;
        } catch (Throwable e) {
            throw new ExcelCastException(e);
        }
    }

    /**
     * 创建一个新的工作表(Sheet)
     *
     * @param sheetName 工作表名称
     */
    private void createNewSheet(String sheetName) {
        // 重置索引
        rowBoundIndex = 0;
        // 创建新的工作表
        sheet = workbook.createSheet(sheetName);
        // 添加标题行
        addTitleRow(workbookSheet.getTitleStyle());
        // 对其余的行使用格式刷
        formatColumnStyle(workbookSheet.getBodyStyle());
    }

    /**
     * 添加标题行
     *
     * @param style 行样式
     */
    private void addTitleRow(RowStyle style) {
        int titleIndex = rowBoundIndex == 0 ? style.getIndex() : rowBoundIndex;
        SXSSFRow row = sheet.createRow(titleIndex);
        if (style.getHeight() != null) {
            row.setHeightInPoints(style.getHeight());
        }
        for (CellField cellField : cellFields) {
            int index = cellField.getIndex();
            if (workbookSheet.getCellWidth() != null) {
                sheet.setColumnWidth(index, workbookSheet.getCellWidth() * 256 + 184);
            }
            SXSSFCell cell = row.createCell(index);
            cell.setCellStyle(buildCellStyle(style));
            cell.setCellValue(cellField.getName());
        }
        ++rowBoundIndex;
    }

    /**
     * 格式化单元格的样式
     *
     * @param style 行样式
     */
    private void formatColumnStyle(RowStyle style) {
        for (int i = 0; i < cellFields.size(); i++) {
            CellField cellField = cellFields.get(i);
            CellStyle cellStyle = buildCellStyle(style);
            cellStyle.setAlignment(cellField.getAlign().getValue());
            cellStyle.setDataFormat(workbook.createDataFormat().getFormat(cellField.getFormat()));
            sheet.setDefaultColumnStyle(i, cellStyle);
            if (style.getHeight() != null) {
                sheet.setDefaultRowHeightInPoints(style.getHeight());
            }
        }
    }

    /**
     * 填充行数据
     *
     * @param index 行索引
     * @param item  填充单元格的数据
     */
    private void fillDataRow(int index, Object item) {
        SXSSFRow row = sheet.createRow(index);
        RowStyle style = workbookSheet.getBodyStyle();
        if (style.getHeight() != null) {
            row.setHeightInPoints(style.getHeight());
        }
        FieldOperator fieldOperator = BeanUtils.fieldOperate(item.getClass());
        for (CellField cellField : cellFields) {
            SXSSFCell cell = row.createCell(cellField.getIndex());
            Object value = fieldOperator.getValueByFieldName(item, cellField.getField());
            setCellValue(cell, value, cellField.getType());
            cell.setCellStyle(sheet.getColumnStyle(cellField.getIndex()));
        }
    }

    /**
     * 设置单元格的值
     *
     * @param cell  单元格对象
     * @param value 值
     * @param type  值的类型
     */
    private void setCellValue(SXSSFCell cell, Object value, Class<?> type) {
        Map<Object, Object> mapping = workbookSheet.getMapping();
        if (value == null) {
            cell.setCellValue("");
        } else if ((type == Boolean.TYPE || type == Boolean.class) &&
                (mapping == null || !mapping.containsKey(value))) {
            cell.setCellValue((boolean) value);
        } else if ((Number.class.isAssignableFrom(value.getClass()))) {
            cell.setCellValue(Double.parseDouble(value.toString()));
        } else if (type == Date.class) {
            cell.setCellValue((Date) value);
        } else {
            String cellValue = value.toString();
            if (mapping != null && mapping.containsKey(value)) {
                cellValue = mapping.get(value).toString();
            }
            cell.setCellValue(cellValue);
        }
    }

    /**
     * 构建单元格样式实例
     *
     * @param style 行样式
     * @return 返回 {@link CellStyle} 实例
     */
    private CellStyle buildCellStyle(RowStyle style) {
        CellStyle cellStyle = workbook.createCellStyle();
        if (style.getAlign() != null) {
            cellStyle.setAlignment(style.getAlign());
        }
        if (style.getVerticalAlign() != null) {
            cellStyle.setVerticalAlignment(style.getVerticalAlign());
        }
        if (style.getFontName() != null || style.getFontSize() != null || style.getFontColor() != null) {
            cellStyle.setFont(createFont(style.getFontName(), style.getFontSize(), style.getFontColor()));
        }
        if (style.getAutoWrap() != null) {
            cellStyle.setWrapText(style.getAutoWrap());
        }
        if (style.getBorder() != null && style.getBorderColor() != null) {
            cellStyle.setBorderTop(style.getBorder());
            cellStyle.setTopBorderColor(style.getBorderColor());
            cellStyle.setBorderLeft(style.getBorder());
            cellStyle.setLeftBorderColor(style.getBorderColor());
            cellStyle.setBorderRight(style.getBorder());
            cellStyle.setRightBorderColor(style.getBorderColor());
            cellStyle.setBorderBottom(style.getBorder());
            cellStyle.setBottomBorderColor(style.getBorderColor());
        }
        if (style.getBackgroundColor() != null) {
            cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
            cellStyle.setFillForegroundColor(style.getBackgroundColor());
        }
        if (style.getFormat() != null) {
            cellStyle.setDataFormat(workbook.createDataFormat().getFormat(style.getFormat()));
        }
        return cellStyle;
    }

    /**
     * 构建字体实例
     *
     * @param name  字体名称
     * @param size  字体大小
     * @param color 字体颜色
     * @return 返回 {@link Font} 实例
     */
    private Font createFont(String name, Integer size, Short color) {
        Font font = workbook.createFont();
        if (name != null) {
            font.setFontName(name);
        }
        if (color != null) {
            font.setColor(color);
        }
        if (size != null) {
            font.setFontHeightInPoints(size.shortValue());
        }
        return font;
    }

}
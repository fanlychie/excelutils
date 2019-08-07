package com.github.fanlychie.excelutils.read;

import com.github.fanlychie.beanutils.BeanUtils;
import com.github.fanlychie.beanutils.operator.ConstructorOperator;
import com.github.fanlychie.beanutils.operator.FieldOperator;
import com.github.fanlychie.excelutils.annotation.AnnotationHandler;
import com.github.fanlychie.excelutils.annotation.CellField;
import com.github.fanlychie.excelutils.exception.ExcelCastException;
import com.github.fanlychie.excelutils.exception.ReadExcelException;
import lombok.Setter;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.SAXParserFactory;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * EXCEL读取器, 用于读取EXCEL表格的数据到POJO的列表中
 *
 * @author fanlychie
 */
public class ExcelReader {

    @Setter
    private int start;

    @Setter
    private Class<?> targetClass;

    @Setter
    private InputStream stream;

    @Setter
    private Paging paging;

    @Setter
    private PagingReader reader;

    private StylesTable stylesTable;

    private SheetIterator sheetIterator;

    private ReadOnlySharedStringsTable sharedStringsTable;

    private Map<Integer, CellField> cellFieldMapping;

    ExcelReader() {}

    /**
     * 解析指定索引的工作表(Sheet)
     *
     * @param index 工作表索引, 索引值从1开始
     * @param <T>   期望的结果类型
     * @return 返回期望的结果类型的集合
     */
    public <T> List<T> read(int index) {
        int sheetCount = 1;
        while (sheetIterator.hasNext()) {
            if (index == sheetCount++) {
                return processSheet(false);
            }
            sheetIterator.next();
        }
        throw new ReadExcelException("can not found sheet index : " + index);
    }

    /**
     * 解析所有的工作表(Sheet)
     *
     * @param <T> 期望的结果类型
     * @return 返回期望的结果类型的集合
     */
    public <T> List<T> read() {
        List<T> list = new ArrayList<>();
        while (sheetIterator.hasNext()) {
            list.addAll(this.<T>processSheet(false));
        }
        return list;
    }

    /**
     * 分页解析指定索引的工作表(Sheet)
     *
     * @param index 工作表索引, 索引值从1开始
     */
    public void paging(int index) {
        if (paging == null) {
            throw new NullPointerException("Paging can not be null");
        }
        if (reader == null) {
            throw new NullPointerException("PagingReader can not be null");
        }
        int sheetCount = 1;
        while (sheetIterator.hasNext()) {
            if (index == sheetCount++) {
                processSheet(true);
            }
            sheetIterator.next();
        }
        throw new ReadExcelException("can not found sheet index : " + index);
    }

    /**
     * 分页解析所有的工作表(Sheet)
     */
    public void paging() {
        if (paging == null) {
            throw new NullPointerException("Paging can not be null");
        }
        if (reader == null) {
            throw new NullPointerException("PagingReader can not be null");
        }
        while (sheetIterator.hasNext()) {
            processSheet(true);
        }
    }

    // 初始化
    void init() {
        try {
            OPCPackage opcPackage = OPCPackage.open(stream);
            XSSFReader reader = new XSSFReader(opcPackage);
            this.sharedStringsTable = new ReadOnlySharedStringsTable(opcPackage);
            this.stylesTable = reader.getStylesTable();
            this.sheetIterator = (SheetIterator) reader.getSheetsData();
            this.cellFieldMapping = AnnotationHandler.getCellFieldMapping(targetClass);
        } catch (Throwable e) {
            throw new ExcelCastException(e);
        }
    }

    // 解析工作表
    private void parseSheet(InputStream sheetInputStream, final List list, final boolean pagination) throws Throwable {
        final FieldOperator fieldOperator = BeanUtils.fieldOperate(targetClass);
        final ConstructorOperator constructorOperator = BeanUtils.constructorOperate(targetClass);
        XMLReader sheetParser = SAXParserFactory.newInstance().newSAXParser().getXMLReader();
        sheetParser.setContentHandler(new XSSFSheetHandler(stylesTable, sharedStringsTable) {
            Object item = constructorOperator.invokeConstructor();
            @Override
            public void postCellHandle(int index, String name, String value, int row, boolean newRow) {
                if (row >= start) {
                    if (newRow) {
                        if (start != row) {
                            list.add(item);
                            doPaging(list, pagination, false);
                        }
                        item = constructorOperator.invokeConstructor();
                    }
                    CellField cellField = cellFieldMapping.get(index);
                    try {
                        Object cellValue = ValueConverter.convertObjectValue(value, cellField.getType());
                        fieldOperator.setValueByFieldName(item, cellField.getField(), cellValue);
                    } catch (Exception e) {
                        throw new ReadExcelException("Parse " + name + " error : " + e);
                    }
                }
            }
            @Override
            public void endDocument() throws SAXException {
                list.add(item);
                doPaging(list, pagination, true);
            }
        });
        sheetParser.parse(new InputSource(sheetInputStream));
    }

    private void doPaging(final List list, final boolean pagination, final boolean flush) {
        if (pagination) {
            // 解析完一行, 计数+1
            paging.current++;
            // 解析的数据集达到设定的大小
            if (paging.current >= paging.size || flush) {
                // 调用读取器处理数据
                reader.read(list);
                // 重置计数
                paging.current = 0;
                // 清空集合
                list.clear();
            }
        }
    }

    // 处理工作表
    private <T> List<T> processSheet(boolean pagination) {
        InputStream stream = null;
        try {
            stream = sheetIterator.next();
            List<T> list = new ArrayList<>(pagination ? paging.size : 16);
            parseSheet(stream, list, pagination);
            return list;
        } catch (Throwable e) {
            throw new ExcelCastException(e);
        } finally {
            if (stream != null) {
                try {
                    stream.close();
                } catch (IOException e) {
                }
            }
        }
    }

    static class Paging {

        @Setter
        int size = 100;

        int current = 0;

    }

}
package com.github.fanlychie.excelutils.annotation;

import com.github.fanlychie.beanutils.BeanUtils;
import com.github.fanlychie.excelutils.spec.Format;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 注解解析器
 * Created by fanlychie on 2017/3/5.
 */
public final class AnnotationHandler {

    /**
     * 用于存储Class中@Cell注解字段相关的信息
     */
    private static Map<Class<?>, Map<Integer, CellField>> cellFieldCacheMap = new ConcurrentHashMap<>();

    /**
     * 解析{@link CellField}列表
     *
     * @param pojoClass 目标类
     * @return 返回解析出的数据列表
     */
    public static Map<Integer, CellField> getCellFieldMapping(Class<?> pojoClass) {
        if (!cellFieldCacheMap.containsKey(pojoClass)) {
            synchronized (pojoClass) {
                if (!cellFieldCacheMap.containsKey(pojoClass)) {
                    Map<Integer, CellField> mapping = new HashMap<>();
                    for (CellField cellField : parseCellFields(pojoClass)) {
                        mapping.put(cellField.getIndex(), cellField);
                    }
                    cellFieldCacheMap.put(pojoClass, mapping);
                }
            }
        }
        return cellFieldCacheMap.get(pojoClass);
    }

    /**
     * 解析{@link CellField}表
     *
     * @param pojoClass 目标类
     * @return 返回 List<CellField>
     */
    private static List<CellField> parseCellFields(Class<?> pojoClass) {
        Map<Field, Cell> annotationMap = BeanUtils.fieldOperate(pojoClass)
                .getAnnotationFieldMap(Cell.class);
        if (annotationMap.isEmpty()) {
            throw new UnsupportedOperationException(
                    "you must mark the data field with the @Cell annotation in " + pojoClass);
        }
        List<CellField> cellFields = buildCellFields(annotationMap);
        return sortByCellIndex(cellFields);
    }

    /**
     * 构建{@link CellField}列表
     *
     * @param annotationMap 注解信息表
     * @return 返回 List<CellField>
     */
    private static List<CellField> buildCellFields(Map<Field, Cell> annotationMap) {
        List<CellField> cellFields = new ArrayList<>();
        for (Field field : annotationMap.keySet()) {
            Cell cell = annotationMap.get(field);
            CellField cellField = new CellField();
            cellField.setType(field.getType());
            cellField.setField(field.getName());
            cellField.setName(cell.name());
            cellField.setIndex(cell.index());
            cellField.setAlign(cell.align());
            String format = cell.format();
            cellField.setFormat(format != null && !format.isEmpty() ? format : Format.getDefault(field.getType()));
            cellFields.add(cellField);
        }
        return cellFields;
    }

    /**
     * 根据{@link Cell}注解的index排序
     *
     * @param cellFields 单元格字段列表
     * @return 返回排序后的列表
     */
    private static List<CellField> sortByCellIndex(List<CellField> cellFields) {
        Collections.sort(cellFields, new Comparator<CellField>() {
            @Override
            public int compare(CellField cf1, CellField cf2) {
                int i1 = cf1.getIndex();
                int i2 = cf2.getIndex();
                return (i1 < i2) ? -1 : ((i1 == i2) ? 0 : 1);
            }
        });
        return cellFields;
    }

}
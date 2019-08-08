package com.github.fanlychie.excelutils.test;

import com.github.fanlychie.excelutils.annotation.Cell;
import com.github.fanlychie.excelutils.read.ExcelReaderBuilder;
import com.github.fanlychie.excelutils.read.PagingReader;
import com.github.fanlychie.excelutils.spec.Align;
import com.github.fanlychie.excelutils.write.ExcelWriterBuilder;
import com.github.fanlychie.excelutils.write.PagingQuerier;
import com.github.fanlychie.excelutils.write.SheetNameStrategy;
import lombok.Data;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.FixMethodOrder;
import org.junit.Test;
import org.junit.runners.MethodSorters;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

import static org.junit.Assert.assertEquals;

@FixMethodOrder(MethodSorters.NAME_ASCENDING)
public class ExcelBuilderTest {

    private static String filename = "customers.xlsx";

    private static String pathname = System.getProperty("user.dir") + "/";

    private static List<Customer> customers = queryByPage(0, Integer.MAX_VALUE);

    /**
     * 自定义样式
     */
    @Test
    public void testDefine() {
        new ExcelWriterBuilder()
                .payload(Customer.class)
                .define()
                    .title()
                        .fontName("Microsoft YaHei")
                        .fontSize(12)
                        .height(20)
                        .background("LEMON_CHIFFON")
                        .and()
                    .body()
                        .height(18)
                        .fontSize(11)
                        .background("LIGHT_TURQUOISE")
                        .and()
                .build()
                .write(customers)
                .toFile(pathname + filename);
    }

    /**
     * YAML配置文件样式
     */
    @Test
    public void testConfigure() {
        new ExcelWriterBuilder()
                .payload(Customer.class)
                .configure("jexcel-full-config.yml")
                .build()
                .write(customers)
                .toFile(pathname + filename);
    }

    /**
     * 内置样式
     */
    @Test
    public void testBuiltin() {
        new ExcelWriterBuilder()
                .payload(Customer.class)
                .builtin()
                .build()
                .write(customers)
                .toFile(pathname + filename);
    }

    @Test
    public void testRead() {
        List<Customer> list = new ExcelReaderBuilder()
                .stream(pathname + filename)
                .start(2)
                .payload(Customer.class)
                .build()
                .read();
        for (Customer customer : list) {
            System.out.println(customer);
        }
    }

    @Test
    public void testPagingRead() {
        new ExcelReaderBuilder().stream(pathname + filename)
                .payload(Customer.class)
                .start(2) // 从第二行开始解析(第一行是标题行, 跳过)
                .pageSize(2) // 每次处理2条
                .reader(new PagingReader<Customer>(){ // 每当解析EXCEL行数达到pageSize设定的阀值时, 触发PagingReader.read, 可以在此处处理解析的结果
                    @Override
                    public void read(List<Customer> items) {
                        for (Customer item : items) {
                            System.out.println(item);
                        }
                    }
                })
                .build()
                .paging(); // 调用分页处理, 无返回值, 数据分批交于PagingReader.read处理
    }

    @Test
    public void testPagingWrite() {
        new ExcelWriterBuilder()
                // POJO 类
                .payload(Customer.class)
                // 使用内置的样式
                .builtin()
                // 分页查询数据
                .pagingQuery(new PagingQuerier() {
                    @Override
                    public List queryPage(int page, int offset, int size) {
                        return queryByPage(offset, size);
                    }
                })
                // 每次查询200条数据
                .pageSize(200)
                // 每个Sheet页最多写500行
                .maxRowsPerSheet(500)
                // Sheet页的名称策略
                .sheetNameStrategy(new SheetNameStrategy() {
                    @Override
                    public String getSheetName(int sheetIndex) {
                        return "Sheet" + sheetIndex;
                    }
                })
                .and()
                // 构建实例
                .build()
                // 分页处理
                .paging()
                // 将数据写出到文件
                .toFile(pathname + filename);
    }

    @Data
    public static class Customer {

        @Cell(index = 0, name = "姓名", align = Align.CENTER)
        private String name;

        @Cell(index = 1, name = "手机", align = Align.CENTER)
        private String mobile;

        @Cell(index = 2, name = "年龄", align = Align.CENTER)
        private int age;

    }

    /**
     * 测试分页查询
     */
    @Test
    public void testQueryByPage() {
        List<Customer> customers = queryByPage(0, 5);
        for (Customer customer : customers) {
            System.out.println(customer);
        }
        assertEquals(5, customers.size());
    }

    /**
     * 模拟数据库的分页查询
     *
     * @param offset 起始索引
     * @param size   每页的大小
     */
    private static List<Customer> queryByPage(int offset, int size) {
        List<Customer> list = new ArrayList<>();
        int current = 0;
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(
                ExcelBuilderTest.class.getResourceAsStream("/customers.txt"), "UTF-8"))) {
            String line;
            while ((line = reader.readLine()) != null) {
                if (current++ < offset) {
                    continue;
                }
                list.add(convertCustomer(line));
                if (current >= offset + size) {
                    break;
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return list;
    }

    private static Customer convertCustomer(String line) {
        Customer customer = new Customer();
        String[] items = line.split(",");
        customer.setName(items[0]);
        customer.setMobile(items[1]);
        customer.setAge(Integer.parseInt(items[2]));
        return customer;
    }

    @BeforeClass
    public static void before() {
        System.out.println(">>>>>>>>>>>>> 单元测试开始");
    }

    @AfterClass
    public static void after() {
        System.out.println("<<<<<<<<<<<<< 单元测试结束");
    }

}
package com.github.fanlychie.excelutils.test;

import com.github.fanlychie.excelutils.annotation.Cell;
import com.github.fanlychie.excelutils.read.ExcelReaderBuilder;
import com.github.fanlychie.excelutils.read.PagingHandler;
import com.github.fanlychie.excelutils.spec.Align;
import com.github.fanlychie.excelutils.write.ExcelWriterBuilder;
import com.github.fanlychie.excelutils.write.PagingQuery;
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
import java.util.List;

import static org.junit.Assert.assertEquals;

@FixMethodOrder(MethodSorters.NAME_ASCENDING)
public class ExcelBuilderTest {

    private static String filename = "customers.xlsx";

    private static String pathname = System.getProperty("user.dir") + "/";

    private static List<Customer> customers = selectByPage(0, Integer.MAX_VALUE);

    /**
     * 自定义EXCEL样式, 将数据导出到文档
     */
    @Test
    public void testDefine() {
        new ExcelWriterBuilder()
                // 数据载体, POJO 类
                .payload(Customer.class)
                // 启用自定义样式配置
                .define()
                    // 标题行配置
                    .title()
                        // 字体名称
                        .fontName("Microsoft YaHei")
                        // 字体大小
                        .fontSize(12)
                        // 单元格行高
                        .height(20)
                        // 单元格背景颜色
                        .background("LEMON_CHIFFON")
                        // 完成配置, 返回上层
                        .complete()
                    // 主体行配置
                    .body()
                        // 单元格行高
                        .height(18)
                        // 字体大小
                        .fontSize(11)
                        // 背景颜色
                        .background("LIGHT_TURQUOISE")
                        // 完成配置, 返回上层
                        .complete()
                // 构建EXCEL写实例
                .build()
                    // 写出数据到文档
                    .write(customers)
                    // 输出文档到文件
                    .toFile(pathname + filename);
    }

    /**
     * YAML配置文件样式, 将数据导出到文档
     */
    @Test
    public void testConfigure() {
        new ExcelWriterBuilder()
                // 数据载体, POJO 类
                .payload(Customer.class)
                // 启用YAML样式配置文件
                // 参考[ https://github.com/fanlychie/excelutils/blob/master/src/main/resources/jexcel-full-config.yml ]
                .configure("jexcel-full-config.yml")
                // 构建EXCEL写实例
                .build()
                    // 写出数据到文档
                    .write(customers)
                    // 输出文档到文件
                    .toFile(pathname + filename);
    }

    /**
     * 内置样式, 将数据导出到文档
     */
    @Test
    public void testBuiltin() {
        new ExcelWriterBuilder()
                // 数据载体, POJO 类
                .payload(Customer.class)
                // 启用内置的样式
                .builtin()
                // 构建EXCEL写实例
                .build()
                    // 写出数据到文档
                    .write(customers)
                    // 输出文档到文件
                    .toFile(pathname + filename);
    }

    /**
     * 内置样式, 将数据导出到文档
     * 使用分页查询, 每次查询一页数据, 然后写入EXCEL, 再查询一页, 然后追加到EXCEL, 以此循环, 直至分页数据全部处理完成
     */
    @Test
    public void testPagingWrite() {
        new ExcelWriterBuilder()
                // 数据载体, POJO 类
                .payload(Customer.class)
                // 启用内置的样式
                .builtin()
                // 分页查询数据
                .pagingQuery(new PagingQuery() {
                    @Override
                    public List queryByPage(int page, int offset, int size) {
                        // 每处理完一页, offset会自动更新成下一页的起始偏移索引
                        // 如每页查询10条, 第一页offset为0, 第二页offset会自动更新成10
                        // 如使用MySQL, SQL语句类似于: LIMIT OFFSET, SIZE
                        return selectByPage(offset, size);
                    }
                })
                    // 每页的数据大小设为200条
                    .pageSize(200)
                    // 每个Sheet页设为最大的行数为500条数据
                    // 如果分页的查询的总数据超出这个阀值, 则自动创建另一个新的Sheet页
                    .maxRowsPerSheet(500)
                    // Sheet页的名称策略, 因分页的查询的总数据可能会超出设置的阀值, 因此需为自动新建的Sheet页设计一个命名的策略
                    .sheetNameStrategy(new SheetNameStrategy() {
                        @Override
                        public String getSheetName(int sheetIndex) {
                            return "优质客户表-" + sheetIndex;
                        }
                    })
                    // 完成配置, 返回上层
                    .complete()
                // 构建EXCEL写实例
                .build()
                    // 启用分页查询写出
                    .paging()
                    // 输出文档到文件
                    .toFile(pathname + filename);
    }

    /**
     * 读取EXCEL文件
     */
    @Test
    public void testRead() {
        List<Customer> list = new ExcelReaderBuilder()
                                // 数据载体, POJO 类
                                .payload(Customer.class)
                                // EXCEL文件流
                                .stream(pathname + filename)
                                // 从第二行开始解析(第一行是标题行, 跳过)
                                .start(2)
                                // 构建EXCEL读实例
                                .build()
                                    // 开始读取
                                    .read();
        int current = 1;
        for (Customer customer : list) {
            System.out.println(current++ + ". " + customer);
        }
    }

    /**
     * 读取EXCEL文件, 通过指定解析的Sheet索引值
     */
    @Test
    public void testReadIndex() {
        List<Customer> list = new ExcelReaderBuilder()
                                // 数据载体, POJO 类
                                .payload(Customer.class)
                                // EXCEL文件流
                                .stream(pathname + filename)
                                // 从第二行开始解析(第一行是标题行, 跳过)
                                .start(2)
                                // 构建EXCEL读实例
                                .build()
                                    // 只读取第一个Sheet页的数据
                                    .read(1);
        int current = 1;
        for (Customer customer : list) {
            System.out.println(current++ + ". " + customer);
        }
    }

    /**
     * 分页读取EXCEL文件
     */
    @Test
    public void testPagingRead() {
        new ExcelReaderBuilder()
                // 数据载体, POJO 类
                .payload(Customer.class)
                // EXCEL文件流
                .stream(pathname + filename)
                // 从第二行开始解析(第一行是标题行, 跳过)
                .start(2)
                // 每次处理EXCEL文件中的100行
                .pageSize(100)
                // 每页读取数据处理
                // 当解析EXCEL文件内容的行数达到pageSize设定的阀值时, 触发PagingHandler.handle函数来处理数据
                .pagingHandler(new PagingHandler<Customer>(){
                    @Override
                    public void handle(List<Customer> items) {
                        for (Customer item : items) {
                            System.out.println(item);
                        }
                        System.out.println("=========================================================================");
                    }
                })
                // 构建EXCEL读实例
                .build()
                    // 启用分页处理, 数据分批交于PagingHandler.handle处理
                    .paging();
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
        List<Customer> customers = selectByPage(0, 5);
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
    private static List<Customer> selectByPage(int offset, int size) {
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
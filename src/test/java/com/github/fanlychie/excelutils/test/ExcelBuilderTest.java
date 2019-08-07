package com.github.fanlychie.excelutils.test;

import com.github.fanlychie.excelutils.annotation.Cell;
import com.github.fanlychie.excelutils.read.ExcelReaderBuilder;
import com.github.fanlychie.excelutils.read.PagingReader;
import com.github.fanlychie.excelutils.spec.Align;
import com.github.fanlychie.excelutils.write.ExcelWriterBuilder;
import lombok.Data;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.FixMethodOrder;
import org.junit.Test;
import org.junit.runners.MethodSorters;

import java.util.LinkedList;
import java.util.List;

@FixMethodOrder(MethodSorters.NAME_ASCENDING)
public class ExcelBuilderTest {

    private static List<Customer> customers;

    private static String filename = "customers.xlsx";

    private static String pathname = System.getProperty("user.dir") + "/";

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

    @Data
    public static class Customer {

        @Cell(index = 0, name = "姓名", align = Align.CENTER)
        private String name;

        @Cell(index = 1, name = "手机", align = Align.CENTER)
        private String mobile;

        @Cell(index = 2, name = "年龄", align = Align.CENTER)
        private int age;

    }

    @BeforeClass
    public static void before() {
        System.out.println(">>>>>>>>>>>>> 单元测试开始");
        customers = new LinkedList<>();
        Customer customer1 = new Customer();
        customer1.name = "张三";
        customer1.mobile = "13800138000";
        customer1.age = 20;
        Customer customer2 = new Customer();
        customer2.name = "李四";
        customer2.mobile = "13800138001";
        customer2.age = 21;
        Customer customer3 = new Customer();
        customer3.name = "王五";
        customer3.mobile = "13800138002";
        customer3.age = 22;
        Customer customer4 = new Customer();
        customer4.name = "赵六";
        customer4.mobile = "13800138003";
        customer4.age = 25;
        Customer customer5 = new Customer();
        customer5.name = "孙七";
        customer5.mobile = "13800138004";
        customer5.age = 26;
        Customer customer6 = new Customer();
        customer6.name = "周八";
        customer6.mobile = "13800138005";
        customer6.age = 23;
        Customer customer7 = new Customer();
        customer7.name = "吴九";
        customer7.mobile = "13800138006";
        customer7.age = 24;
        customers.add(customer1);
        customers.add(customer2);
        customers.add(customer3);
        customers.add(customer4);
        customers.add(customer5);
        customers.add(customer6);
        customers.add(customer7);
    }

    @AfterClass
    public static void after() {
        System.out.println("<<<<<<<<<<<<< 单元测试结束");
    }

}
package com.github.fanlychie.excelutils.test.read;

import lombok.Data;
import com.github.fanlychie.excelutils.annotation.Cell;
import com.github.fanlychie.excelutils.read.ExcelReaderBuilder;
import com.github.fanlychie.excelutils.spec.Align;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;

import java.util.List;

public class ExcelReaderBuilderTest {

    private static String pathname = System.getProperty("user.dir");

    @Test
    public void testParse() {
        List<Customer> list = new ExcelReaderBuilder()
                .stream(pathname + "Customer_Define.xlsx")
                .start(2)
                .payload(Customer.class)
                .build()
                .parse();
        for (Customer customer : list) {
            System.out.println(customer);
        }
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
        System.out.println("[文件读取] >>>>>>>>>>>>> 单元测试开始");
    }

    @AfterClass
    public static void after() {
        System.out.println("[文件读取] <<<<<<<<<<<<< 单元测试结束");
    }

}
<p align="center">
    <a href="#">
        <img src="https://raw.githubusercontent.com/fanlychie/mdimg/master/excelutils_logo.png">
    </a>
</p>
<p align="center">
    基于 POI 的 Excel 文件读写
</p>
<p align="center">
    <a href="https://circleci.com/gh/fanlychie/excelutils" target="_blank" title="Circle CI">
        <img src="https://circleci.com/gh/fanlychie/excelutils.svg?style=svg&circle-token=1173052afd21856384d886a4aac200286199cc15">
    </a>
    <a href="https://codecov.io/gh/fanlychie/excelutils" target="_blank" title="Codecov">
        <img src="https://codecov.io/gh/fanlychie/excelutils/branch/master/graph/badge.svg" />
    </a>
    <a href="https://www.codacy.com/app/fanlychie/excelutils?utm_source=github.com&amp;utm_medium=referral&amp;utm_content=fanlychie/excelutils&amp;utm_campaign=Badge_Grade" target="_blank" title="Codacy">
        <img src="https://api.codacy.com/project/badge/Grade/84ba5a46a2844916836f77c038bc51a0"/>
    </a>
    <a href="http://www.apache.org/licenses/LICENSE-2.0" target="_blank" title="License">
        <img src="https://img.shields.io/github/license/fanlychie/excelutils.svg">
    </a>
    <a href="https://jitpack.io/#fanlychie/excelutils" target="_blank" title="Jitpack">
        <img src="https://jitpack.io/v/fanlychie/excelutils.svg">
    </a>
</p>

---

## 依赖声明

```xml
<repositories>
    <repository>
        <id>jitpack.io</id>
        <url>https://jitpack.io</url>
    </repository>
</repositories>

<dependencies>
    ... ...
    <dependency>
        <groupId>com.github.fanlychie</groupId>
        <artifactId>excelutils</artifactId>
        <version>1.0.0</version>
    </dependency>
</dependencies>
```

## 使用样例

```java
@Data
public class Customer {

    @Cell(index = 0, name = "姓名", align = Align.CENTER)
    private String name;

    @Cell(index = 1, name = "手机", align = Align.CENTER)
    private String mobile;

    @Cell(index = 2, name = "年龄", align = Align.CENTER)
    private int age;

}
 ````
 
### 写出数据到EXCEL文件(内置样式)

 ```java
@Test
public void testBuiltin() {
    new ExcelWriterBuilder()
            .payload(Customer.class)
            .builtin()
            .build()
            .write(customers)
            .toFile(pathname + "Customer_Builtin.xlsx");
}
```
 
### 写出数据到EXCEL文件(YAML配置文件样式)

 ```java
@Test
public void testConfigure() {
    new ExcelWriterBuilder()
            .payload(Customer.class)
            .configure("jexcel-full-config.yml")
            .build()
            .write(customers)
            .toFile(pathname + "Customer_Configure.xlsx");

}
```
 
### 写出数据到EXCEL文件(自定义样式)

 ```java
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
            .toFile(pathname + "Customer_Define.xlsx");
}
```

### 读取EXCEL文件到POJO

```java
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
```
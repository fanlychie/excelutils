package com.github.fanlychie.excelutils.test;

import com.github.fanlychie.excelutils.test.read.ExcelReaderBuilderTest;
import com.github.fanlychie.excelutils.test.write.ExcelWriterBuilderTest;
import org.junit.runner.RunWith;
import org.junit.runners.Suite;

@RunWith(Suite.class)
@Suite.SuiteClasses({ExcelWriterBuilderTest.class, ExcelReaderBuilderTest.class})
public class ExcelBuilderTest {}
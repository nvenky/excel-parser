package com.thoughtworks.excelparser.helper;

import com.thoughtworks.excelparser.exception.ExcelParsingException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.Date;
import java.util.function.Consumer;

import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.CoreMatchers.not;
import static org.hamcrest.core.IsNull.nullValue;
import static org.junit.Assert.assertThat;

public class HSSFHelperTest {
    Sheet sheet;
    String sheetName = "Sheet1";
    InputStream inputStream;

    @Before
    public void setUp() throws IOException {
        inputStream = getClass().getClassLoader().getResourceAsStream("Student Profile.xls");
        sheet = new HSSFWorkbook(inputStream).getSheet(sheetName);
    }

    @After
    public void tearDown() throws IOException {
        inputStream.close();
    }

    @Test(expected = ExcelParsingException.class)
    public void testShouldThrowExceptionOnInvalidDateCell() {
        int rowNumber = 2;
        int columnNumber = 2;
        Cell cell = HSSFHelper.getCell(sheet, rowNumber, columnNumber);
        HSSFHelper.getDateCell(cell, new Locator(sheetName, rowNumber, columnNumber), e -> {
            throw e;
        });
    }

    @Test
    public void testShouldReturnValidDate() {
        int rowNumber = 6;
        int columnNumber = 4;
        Cell cell = HSSFHelper.getCell(sheet, rowNumber, columnNumber);
        Date actual = HSSFHelper.getDateCell(cell, new Locator(sheetName, rowNumber, columnNumber), e -> {
            throw e;
        });

        assertThat(actual, is(not(nullValue(Date.class))));
    }

    @Test
    public void testShouldReturnValidStringValue() {
        Consumer<ExcelParsingException> errorHandler = e -> {
            throw e;
        };
        assertThat(HSSFHelper.getStringCell(HSSFHelper.getCell(sheet, 6, 1), errorHandler), is("1"));
        assertThat(HSSFHelper.getStringCell(HSSFHelper.getCell(sheet, 6, 5), errorHandler), is("A"));
        assertThat(HSSFHelper.getStringCell(HSSFHelper.getCell(sheet, 8, 3), errorHandler), is("James"));
    }

    @Test
    public void testShouldReturnValidNumericValue() {
        assertThat(HSSFHelper.getIntegerCell(HSSFHelper.getCell(sheet, 6, 1), false, new Locator(sheetName, 6, 1), e -> { throw e; }), is(1));
        assertThat(HSSFHelper.getIntegerCell(HSSFHelper.getCell(sheet, 6, 8), true, new Locator(sheetName, 6, 8),e -> { throw e; }), is(0));
        assertThat(HSSFHelper.getIntegerCell(HSSFHelper.getCell(sheet, 6, 8), false, new Locator(sheetName, 6, 8), e -> { throw e; }), is(nullValue()));

        assertThat(HSSFHelper.getLongCell(HSSFHelper.getCell(sheet, 6, 2), false, new Locator(sheetName, 6, 2), e -> { throw e; }), is(2001L));
        assertThat(HSSFHelper.getLongCell(HSSFHelper.getCell(sheet, 10, 2), true, new Locator(sheetName, 10, 2), e -> { throw e; }), is(0L));
        assertThat(HSSFHelper.getLongCell(HSSFHelper.getCell(sheet, 10, 2), false, new Locator(sheetName, 10, 2), e -> { throw e; }), is(nullValue(Long.class)));

        assertThat(HSSFHelper.getDoubleCell(HSSFHelper.getCell(sheet, 7, 8), false, new Locator(sheetName, 7, 8), e -> { throw e; }), is(450.3d));
        assertThat(HSSFHelper.getDoubleCell(HSSFHelper.getCell(sheet, 8, 8), false, new Locator(sheetName, 8, 8), e -> { throw e; }), is(300d));
    }

}

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
    HSSFHelper hssfHelper;
    Sheet sheet;
    String sheetName = "Sheet1";
    InputStream inputStream;

    @Before
    public void setUp() throws IOException {
        hssfHelper = new HSSFHelper();
        inputStream = getClass().getClassLoader().getResourceAsStream("Student Profile.xls");
        sheet = new HSSFWorkbook(inputStream).getSheet(sheetName);
    }

    @After
    public void tearDown() throws IOException {
        inputStream.close();
    }

    @Test(expected = ExcelParsingException.class)
    public void testShouldThrowExceptionOnInvalidDateCell() throws ExcelParsingException {
        int rowNumber = 2;
        int columnNumber = 2;
        Cell cell = hssfHelper.getCell(sheet, rowNumber, columnNumber);
        hssfHelper.getDateCell(cell, new Locator(sheetName, rowNumber, columnNumber), e -> {
            throw e;
        });
    }

    @Test
    public void testShouldReturnValidDate() throws ExcelParsingException {
        int rowNumber = 6;
        int columnNumber = 4;
        Cell cell = hssfHelper.getCell(sheet, rowNumber, columnNumber);
        Date actual = hssfHelper.getDateCell(cell, new Locator(sheetName, rowNumber, columnNumber), e -> {
            throw e;
        });

        assertThat(actual, is(not(nullValue(Date.class))));
    }

    @Test
    public void testShouldReturnValidStringValue() throws ExcelParsingException {
        Consumer<ExcelParsingException> errorHandler = e -> {
            throw e;
        };
        assertThat(hssfHelper.getStringCell(hssfHelper.getCell(sheet, 6, 1), errorHandler), is("1"));
        assertThat(hssfHelper.getStringCell(hssfHelper.getCell(sheet, 6, 5), errorHandler), is("A"));
        assertThat(hssfHelper.getStringCell(hssfHelper.getCell(sheet, 8, 3), errorHandler), is("James"));
    }

    @Test
    public void testShouldReturnValidNumericValue() throws ExcelParsingException {
        assertThat(hssfHelper.getIntegerCell(hssfHelper.getCell(sheet, 6, 1), false, new Locator(sheetName, 6, 1), e -> { throw e; }), is(1));
        assertThat(hssfHelper.getIntegerCell(hssfHelper.getCell(sheet, 6, 8), true, new Locator(sheetName, 6, 8),e -> { throw e; }), is(0));
        assertThat(hssfHelper.getIntegerCell(hssfHelper.getCell(sheet, 6, 8), false, new Locator(sheetName, 6, 8), e -> { throw e; }), is(nullValue()));

        assertThat(hssfHelper.getLongCell(hssfHelper.getCell(sheet, 6, 2), false, new Locator(sheetName, 6, 2), e -> { throw e; }), is(2001L));
        assertThat(hssfHelper.getLongCell(hssfHelper.getCell(sheet, 10, 2), true, new Locator(sheetName, 10, 2), e -> { throw e; }), is(0L));
        assertThat(hssfHelper.getLongCell(hssfHelper.getCell(sheet, 10, 2), false, new Locator(sheetName, 10, 2), e -> { throw e; }), is(nullValue(Long.class)));

        assertThat(hssfHelper.getDoubleCell(hssfHelper.getCell(sheet, 7, 8), false, new Locator(sheetName, 7, 8), e -> { throw e; }), is(450.3d));
        assertThat(hssfHelper.getDoubleCell(hssfHelper.getCell(sheet, 8, 8), false, new Locator(sheetName, 8, 8), e -> { throw e; }), is(300d));
    }

}

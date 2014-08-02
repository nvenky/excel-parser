package com.thoughtworks.excelparser;

import com.thoughtworks.example.domain.Section;
import com.thoughtworks.excelparser.exception.ExcelParsingException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.List;

import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.CoreMatchers.nullValue;
import static org.junit.Assert.assertThat;

public class SheetParserTest {

    InputStream inputStream;

    @After
    public void tearDown() throws IOException {
        inputStream.close();
    }

    @Test
    public void shouldCreateEntityBasedOnAnnotationFromExcel97File() throws ExcelParsingException, IOException {
        performTestUsing(openSheet("Student Profile.xls"));
    }

    @Test
    public void shouldCreateEntityBasedOnAnnotationFromExcel2007File() throws ExcelParsingException, IOException {
        performTestUsing(openSheet("Student Profile.xlsx"));
    }

    private void performTestUsing(Sheet sheet) throws ExcelParsingException {
        SheetParser parser = new SheetParser();
        List<Section> entityList = parser.createEntity(sheet, "Sheet1", Section.class);
        assertThat(entityList.size(), is(1));
        Section section = entityList.get(0);
        assertThat(section.getYear(), is("IV"));
        assertThat("B", section.getSection(), is("B"));
        assertThat(section.getStudents().size(), is(3));

        assertThat(section.getStudents().get(0).getRoleNumber(), is(2001L));
        assertThat(section.getStudents().get(0).getName(), is("Adam"));

        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("dd/MM/yyyy");

        assertThat(simpleDateFormat.format(section.getStudents().get(0).getDateOfBirth()), is("01/01/2002"));
        assertThat(section.getStudents().get(0).getFatherName(), is("A"));
        assertThat("D", section.getStudents().get(0).getMotherName(), is("D"));
        assertThat("XYZ", section.getStudents().get(0).getAddress(), is("XYZ"));
        assertThat(section.getStudents().get(0).getTotalScore(), is(nullValue(Double.class)));
    }

    private Sheet openSheet(String fileName) throws IOException {
        inputStream = getClass().getClassLoader().getResourceAsStream(fileName);
        Workbook workbook;
        if(fileName.endsWith("xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            workbook = new XSSFWorkbook(inputStream);
        }
        return workbook.getSheet("Sheet1");
    }
}

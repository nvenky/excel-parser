package org.javafunk.excelparser;

import org.javafunk.example.domain.Section;
import org.javafunk.excelparser.exception.ExcelParsingException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Test;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import static java.math.BigDecimal.ROUND_FLOOR;
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
    public void shouldCreateEntityBasedOnAnnotationFromExcel97File() throws Exception {
        performTestUsing(openSheet("Student Profile.xls"));
    }

    @Test
    public void shouldCreateEntityBasedOnAnnotationFromExcel2007File() throws Exception {
        performTestUsing(openSheet("Student Profile.xlsx"));
    }

    @Test
    public void shouldCallErrorHandlerWhenRowCannotBeParsed() throws Exception {
        List<ExcelParsingException> errors = new ArrayList<>();
        SheetParser parser = new SheetParser();

        List<Section> entityList = parser.createEntity(openSheet("Errors.xlsx"), Section.class, errors::add);

        assertThat(entityList.size(), is(1));
        Section section = entityList.get(0);
        assertThat(section.getStudents().get(0).getDateOfBirth(), is(nullValue(Date.class)));

        assertThat(errors.size(), is(3));
        assertThat(errors.get(0).getMessage(), is("Invalid date found in sheet Sheet1 at row 6, column 4"));
        assertThat(errors.get(1).getMessage(), is("Invalid date found in sheet Sheet1 at row 7, column 4"));
        assertThat(errors.get(2).getMessage(), is("Invalid date found in sheet Sheet1 at row 8, column 4"));
    }

    private void performTestUsing(Sheet sheet) {
        SheetParser parser = new SheetParser();
        List<Section> entityList = parser.createEntity(sheet, Section.class, error -> { throw error; });
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
        assertThat(section.getStudents().get(0).getTotalScore(), is(nullValue(BigDecimal.class)));
        assertThat(section.getStudents().get(0).getAdmissionDate(), is(LocalDate.of(2002, 10, 10)));
        assertThat(section.getStudents().get(0).getAdmissionDateTime(), is(LocalDateTime.of(2002, 10, 10, 9, 0, 0)));

        assertThat(section.getStudents().get(1).getRoleNumber(), is(2002L));
        assertThat(section.getStudents().get(1).getName(), is("Even"));
        assertThat(simpleDateFormat.format(section.getStudents().get(1).getDateOfBirth()), is("05/01/2002"));
        assertThat(section.getStudents().get(1).getFatherName(), is("B"));
        assertThat("D", section.getStudents().get(1).getMotherName(), is("E"));
        assertThat("XYZ", section.getStudents().get(1).getAddress(), is("ABX"));
        assertThat(section.getStudents().get(1).getTotalScore().setScale(2, ROUND_FLOOR), is(new BigDecimal("450.30")));
        assertThat(section.getStudents().get(1).getAdmissionDate(), is(LocalDate.of(2002, 10, 11)));
        assertThat(section.getStudents().get(1).getAdmissionDateTime(), is(LocalDateTime.of(2002, 10, 11, 10, 0, 0)));
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

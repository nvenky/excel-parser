package org.javafunk.example.domain;

import org.javafunk.excelparser.annotations.ExcelField;
import org.javafunk.excelparser.annotations.ExcelObject;
import org.javafunk.excelparser.annotations.MappedExcelObject;
import org.javafunk.excelparser.annotations.ParseType;
import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Value;
import lombok.experimental.FieldDefaults;

import java.util.List;

@Value
@AllArgsConstructor
@FieldDefaults(level = AccessLevel.PRIVATE)
@ExcelObject(parseType = ParseType.COLUMN, start = 2, end = 2)
public class Section {

    @ExcelField(position = 2)
    String year;

    @ExcelField(position = 3)
    String section;

    @MappedExcelObject
    List<Student> students;

    @SuppressWarnings("UnusedDeclaration")
    private Section() {
        this(null, null, null);
    }
}

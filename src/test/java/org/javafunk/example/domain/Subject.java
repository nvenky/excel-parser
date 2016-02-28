package org.javafunk.example.domain;

import org.javafunk.excelparser.annotations.ExcelField;
import org.javafunk.excelparser.annotations.ExcelObject;
import org.javafunk.excelparser.annotations.ParseType;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Value;
import lombok.experimental.FieldDefaults;

@Value
@AllArgsConstructor
@FieldDefaults(level = AccessLevel.PRIVATE)
@ExcelObject(parseType = ParseType.ROW, start = 2)
public class Subject {

    @ExcelField(position = 1)
    String code;

    @ExcelField(position = 2)
    String name;

    @ExcelField(position = 3)
    Integer volume;

    @SuppressWarnings("UnusedDeclaration")
    private Subject() {
	this(null, null, null);
    }
}

package com.thoughtworks.example.domain;

import com.thoughtworks.excelparser.annotations.ExcelField;
import com.thoughtworks.excelparser.annotations.ExcelObject;
import com.thoughtworks.excelparser.annotations.ParseType;
import lombok.AccessLevel;
import lombok.Data;
import lombok.experimental.FieldDefaults;

import java.util.Date;

@Data
@FieldDefaults(level = AccessLevel.PRIVATE)
@ExcelObject(parseType = ParseType.ROW, start = 6, end = 8)
public class Student {

    @ExcelField(position = 2)
    Long roleNumber;

    @ExcelField(position = 3)
    String name;

    @ExcelField(position = 4)
    Date dateOfBirth;

    @ExcelField(position = 5)
    String fatherName;

    @ExcelField(position = 6)
    String motherName;

    @ExcelField(position = 7)
    String address;

    @ExcelField(position = 8)
    Double totalScore;

}

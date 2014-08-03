package com.thoughtworks.excelparser.helper;

import lombok.Value;

@Value
public class Locator {
    String sheetName;
    int row;
    int col;
}

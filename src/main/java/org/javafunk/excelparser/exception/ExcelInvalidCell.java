package org.javafunk.excelparser.exception;

/**
 * Created by Victor.Ikoro on 4/14/2016.
 */
public class ExcelInvalidCell {

    private int row;
    private int column;
    private String message;
    private String value;

    public ExcelInvalidCell(int row, int column)
    {
        this(row, column, null, null);
    }
    public ExcelInvalidCell(int row, int column, String message)
    {
        this(row, column, message, null);
    }

    public ExcelInvalidCell(int row, int column, String value, String message)
    {
        this.row = row;
        this.column = column;
        this.value =  value;
        this.message =  message;
    }


    public int getColumn() {
        return column;
    }

    public int getRow() {
        return row;
    }

    public String getMessage() {
        return message;
    }

    public String getValue() {
        return value;
    }

    public void setColumn(int column) {
        this.column = column;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    public void setRow(int row) {
        this.row = row;
    }

    public void setValue(String value) {
        this.value = value;
    }
}

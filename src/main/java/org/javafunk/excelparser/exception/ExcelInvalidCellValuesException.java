package org.javafunk.excelparser.exception;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Created by Victor.Ikoro on 4/14/2016.
 */
public class ExcelInvalidCellValuesException extends  ExcelParsingException {
    List<ExcelInvalidCell> invalidCells;
    public ExcelInvalidCellValuesException(String message) {
        super(message);
        invalidCells = new ArrayList<>();
    }

    public ExcelInvalidCellValuesException(String message, Exception exception) {
        super(message, exception);
        invalidCells = new ArrayList<>();
    }

    public List<ExcelInvalidCell> getInvalidCells() {
        return invalidCells;
    }

    public void setInvalidCells(List<ExcelInvalidCell> invalidCells) {
        this.invalidCells = invalidCells;
    }

    public void addInvalidCell(ExcelInvalidCell excelInvalidCell)
    {
        invalidCells.add(excelInvalidCell);
    }
}

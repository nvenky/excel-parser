package com.thoughtworks.excelparser.helper;

import com.thoughtworks.excelparser.exception.ExcelParsingException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.text.DecimalFormat;
import java.text.MessageFormat;
import java.util.Date;
import java.util.function.Consumer;

public class HSSFHelper {

    @SuppressWarnings("unchecked")
    public static <T> T getCellValue(Sheet sheet, String sheetName, Class<T> type, Integer row, Integer col, boolean zeroIfNull, Consumer<ExcelParsingException> errorHandler) {
        Cell cell = getCell(sheet, row, col);
        if (type.equals(String.class)) {
            return cell == null ? null : (T) getStringCell(cell, errorHandler);
        } else if (type.equals(Date.class)) {
            return cell == null ? null : (T) getDateCell(cell, new Locator(sheetName, row, col), errorHandler);
        }

        if (type.equals(Integer.class)) {
            return (T) getIntegerCell(cell, zeroIfNull, sheetName, row, col, errorHandler);
        } else if (type.equals(Double.class)) {
            return (T) getDoubleCell(cell, zeroIfNull, sheetName, row, col, errorHandler);
        } else if (type.equals(Long.class)) {
            return (T) getLongCell(cell, zeroIfNull, sheetName, row, col, errorHandler);
        }
        throw new ExcelParsingException(getErrorMessage("{0} Data type not supported for parsing", type.getName()));
    }

    static Cell getCell(Sheet sheet, int rowNumber, int columnNumber) {
        Row row = sheet.getRow(rowNumber - 1);
        return row == null ? null : row.getCell(columnNumber - 1);
    }

    static String getStringCell(Cell cell, Consumer<ExcelParsingException> errorHandler) {
        if (cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA) {
            int type = cell.getCachedFormulaResultType();
            switch (type) {
                case HSSFCell.CELL_TYPE_NUMERIC:
                    DecimalFormat df = new DecimalFormat("###.#");
                    return df.format(cell.getNumericCellValue());
                case HSSFCell.CELL_TYPE_ERROR:
                    return "";
                case HSSFCell.CELL_TYPE_STRING:
                    return cell.getRichStringCellValue().getString().trim();
                case HSSFCell.CELL_TYPE_BOOLEAN:
                    return "" + cell.getBooleanCellValue();

            }
        } else if (cell.getCellType() != HSSFCell.CELL_TYPE_NUMERIC) {
            return cell.getRichStringCellValue().getString().trim();
        }
        DecimalFormat df = new DecimalFormat("###.#");
        return df.format(cell.getNumericCellValue());
    }

    static Date getDateCell(Cell cell, Locator position, Consumer<ExcelParsingException> errorHandler) {
        try {
            if (!HSSFDateUtil.isCellDateFormatted(cell)) {
                errorHandler.accept(new ExcelParsingException(getErrorMessage("Invalid date found in sheet {0} at row {1}, column {2}", position.getSheetName(), position.getRow(), position.getCol())));
            }
            return HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
        } catch (IllegalStateException illegalStateException) {
            errorHandler.accept(new ExcelParsingException(getErrorMessage("Invalid date found in sheet {0} at row {1}, column {2}", position.getSheetName(), position.getRow(), position.getCol())));
        }
        return null;
    }

    private static String getErrorMessage(String errorMessage, Object... errorMessageArgs) {
        return MessageFormat.format(errorMessage, errorMessageArgs);
    }

    static Double getDoubleCell(Cell cell, boolean zeroIfNull, Object... errorMessageArgs) {
        if (cell == null) {
            return zeroIfNull ? 0d : null;
        }

        switch (cell.getCellType()) {
            case HSSFCell.CELL_TYPE_NUMERIC:
            case HSSFCell.CELL_TYPE_FORMULA:
                return cell.getNumericCellValue();
            case HSSFCell.CELL_TYPE_BLANK:
                return zeroIfNull ? 0d : null;
            default:
                throw new ExcelParsingException(getErrorMessage("Invalid number found in sheet {0} at row {1}, column {2}", errorMessageArgs));
        }
    }

    static Long getLongCell(Cell cell, boolean zeroIfNull, Object... errorMessageArgs) {
        Double doubleValue = getNumberWithoutDecimals(cell, zeroIfNull, errorMessageArgs);
        return doubleValue == null ? null : doubleValue.longValue();
    }

    static Integer getIntegerCell(Cell cell, boolean zeroIfNull, Object... errorMessageArgs) {
        Double doubleValue = getNumberWithoutDecimals(cell, zeroIfNull, errorMessageArgs);
        return doubleValue == null ? null : doubleValue.intValue();
    }

    private static Double getNumberWithoutDecimals(Cell cell, boolean zeroIfNull, Object... errorMessageArgs)
            throws ExcelParsingException {
        Double doubleValue = getDoubleCell(cell, zeroIfNull, errorMessageArgs);
        if (doubleValue != null && doubleValue % 1 != 0) {
            throw new ExcelParsingException(getErrorMessage("Invalid number found in sheet {0} at row {1}, column {2}", errorMessageArgs));
        }
        return doubleValue;
    }

}

package com.thoughtworks.excelparser.helper;

import com.thoughtworks.excelparser.exception.ExcelParsingException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Date;
import java.util.function.Consumer;

import static java.text.MessageFormat.format;

public class HSSFHelper {

    @SuppressWarnings("unchecked")
    public static <T> T getCellValue(Sheet sheet, Class<T> type, Integer row, Integer col, boolean zeroIfNull, Consumer<ExcelParsingException> errorHandler) {
        Cell cell = getCell(sheet, row, col);

        if (type.equals(String.class)) {
            return (T) getStringCell(cell, errorHandler);
        }

        if (type.equals(Date.class)) {
            return cell == null ? null : (T) getDateCell(cell, new Locator(sheet.getSheetName(), row, col), errorHandler);
        }

        if (type.equals(LocalDate.class)) {
            return cell == null ? null : (T) getLocalDateCell(cell, new Locator(sheet.getSheetName(), row, col), errorHandler);
        }

        if (type.equals(LocalDateTime.class)) {
            return cell == null ? null : (T) getLocalDateTimeCell(cell, new Locator(sheet.getSheetName(), row, col), errorHandler);
        }

        if (type.equals(Integer.class)) {
            return (T) getIntegerCell(cell, zeroIfNull, new Locator(sheet.getSheetName(), row, col), errorHandler);
        }

        if (type.equals(Double.class)) {
            return (T) getDoubleCell(cell, zeroIfNull, new Locator(sheet.getSheetName(), row, col), errorHandler);
        }

        if (type.equals(Long.class)) {
            return (T) getLongCell(cell, zeroIfNull, new Locator(sheet.getSheetName(), row, col), errorHandler);
        }

        if (type.equals(BigDecimal.class)) {
            return (T) getBigDecimalCell(cell, zeroIfNull, new Locator(sheet.getSheetName(), row, col), errorHandler);
        }

        errorHandler.accept(new ExcelParsingException(format("{0} data type not supported for parsing", type.getName())));
        return null;
    }

    private static LocalDate getLocalDateCell(Cell cell, Locator locator, Consumer<ExcelParsingException> errorHandler) {
        try {
            if (!HSSFDateUtil.isCellDateFormatted(cell)) {
                errorHandler.accept(new ExcelParsingException(format("Invalid date found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol())));
            }

            Instant instant = Instant.ofEpochMilli((long) cell.getNumericCellValue());
            LocalDateTime res = LocalDateTime.ofInstant(instant, ZoneId.systemDefault());
            return res.toLocalDate();

        } catch (IllegalStateException illegalStateException) {
            errorHandler.accept(new ExcelParsingException(format("Invalid date found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol())));
        }
        return null;
    }

    private static LocalDateTime getLocalDateTimeCell(Cell cell, Locator locator, Consumer<ExcelParsingException> errorHandler) {
        try {
            if (!HSSFDateUtil.isCellDateFormatted(cell)) {
                errorHandler.accept(new ExcelParsingException(format("Invalid date found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol())));
            }

            Instant instant = Instant.ofEpochMilli((long) cell.getNumericCellValue());
            return LocalDateTime.ofInstant(instant, ZoneId.systemDefault());

        } catch (IllegalStateException illegalStateException) {
            errorHandler.accept(new ExcelParsingException(format("Invalid date found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol())));
        }
        return null;
    }

    private static BigDecimal getBigDecimalCell(Cell cell, boolean zeroIfNull, Locator locator, Consumer<ExcelParsingException> errorHandler) {
        Double val = getDoubleCell(cell, zeroIfNull, locator, errorHandler);
        if(val == null) {
            if(zeroIfNull) {
                return BigDecimal.ZERO;
            }
            return null;
        }
        return new BigDecimal(val);
    }

    static Cell getCell(Sheet sheet, int rowNumber, int columnNumber) {
        Row row = sheet.getRow(rowNumber - 1);
        return row == null ? null : row.getCell(columnNumber - 1);
    }

    static String getStringCell(Cell cell, Consumer<ExcelParsingException> errorHandler) {
        if (cell == null) {
            return null;
        }

        if (cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA) {
            int type = cell.getCachedFormulaResultType();

            if (type == HSSFCell.CELL_TYPE_NUMERIC) {
                DecimalFormat df = new DecimalFormat("###.#");
                return df.format(cell.getNumericCellValue());
            }

            if (type == HSSFCell.CELL_TYPE_ERROR) {
                return "";
            }

            if (type == HSSFCell.CELL_TYPE_STRING) {
                return cell.getRichStringCellValue().getString().trim();
            }

            if (type == HSSFCell.CELL_TYPE_BOOLEAN) {
                return "" + cell.getBooleanCellValue();
            }

        } else if (cell.getCellType() != HSSFCell.CELL_TYPE_NUMERIC) {
            return cell.getRichStringCellValue().getString().trim();
        }

        DecimalFormat df = new DecimalFormat("###.#");
        return df.format(cell.getNumericCellValue());
    }

    static Date getDateCell(Cell cell, Locator locator, Consumer<ExcelParsingException> errorHandler) {
        try {
            if (!HSSFDateUtil.isCellDateFormatted(cell)) {
                errorHandler.accept(new ExcelParsingException(format("Invalid date found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol())));
            }
            return HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
        } catch (IllegalStateException illegalStateException) {
            errorHandler.accept(new ExcelParsingException(format("Invalid date found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol())));
        }
        return null;
    }

    static Double getDoubleCell(Cell cell, boolean zeroIfNull, Locator locator, Consumer<ExcelParsingException> errorHandler) {
        if (cell == null) {
            return zeroIfNull ? 0d : null;
        }

        if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC || cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA) {
            return cell.getNumericCellValue();
        }

        if (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
            return zeroIfNull ? 0d : null;
        }

        errorHandler.accept(new ExcelParsingException(format("Invalid number found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol())));
        return null;
    }

    static Long getLongCell(Cell cell, boolean zeroIfNull, Locator locator, Consumer<ExcelParsingException> errorHandler) {
        Double doubleValue = getNumberWithoutDecimals(cell, zeroIfNull, locator, errorHandler);
        return doubleValue == null ? null : doubleValue.longValue();
    }

    static Integer getIntegerCell(Cell cell, boolean zeroIfNull, Locator locator, Consumer<ExcelParsingException> errorHandler) {
        Double doubleValue = getNumberWithoutDecimals(cell, zeroIfNull, locator, errorHandler);
        return doubleValue == null ? null : doubleValue.intValue();
    }

    private static Double getNumberWithoutDecimals(Cell cell, boolean zeroIfNull, Locator locator, Consumer<ExcelParsingException> errorHandler) {
        Double doubleValue = getDoubleCell(cell, zeroIfNull, locator, errorHandler);
        if (doubleValue != null && doubleValue % 1 != 0) {
            errorHandler.accept(new ExcelParsingException(format("Invalid number found in sheet {0} at row {1}, column {2}", locator.getSheetName(), locator.getRow(), locator.getCol())));
        }
        return doubleValue;
    }

}

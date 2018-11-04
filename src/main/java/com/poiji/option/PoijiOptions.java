package com.poiji.option;

import com.poiji.exception.PoijiException;

import java.time.format.DateTimeFormatter;
import java.util.Arrays;

import static com.poiji.util.PoijiConstants.DEFAULT_DATE_PATTERN;
import static com.poiji.util.PoijiConstants.DEFAULT_DATE_TIME_FORMATTER;

/**
 * Created by hakan on 17/01/2017.
 */
public final class PoijiOptions {

    private int skip;
    private int sheetIndex;
    private String sheetName;
    private String password;
    private String dateRegex;
    private String datePattern;
    private boolean dateLenient;
    private boolean trimCellValue;
    private boolean ignoreHiddenSheets;
    private boolean preferNullOverDefault;
    private DateTimeFormatter dateTimeFormatter;

    private PoijiOptions() {
        super();
    }

    private PoijiOptions setSkip(int skip) {
        this.skip = skip;
        return this;
    }

    private PoijiOptions setDatePattern(String datePattern) {
        this.datePattern = datePattern;
        return this;
    }

    private PoijiOptions setDateTimeFormatter(DateTimeFormatter dateTimeFormatter) {
        this.dateTimeFormatter = dateTimeFormatter;
        return this;
    }

    private PoijiOptions setPreferNullOverDefault(boolean preferNullOverDefault) {
        this.preferNullOverDefault = preferNullOverDefault;
        return this;
    }

    public String getPassword() {
        return password;
    }

    private PoijiOptions setPassword(String password) {
        this.password = password;
        return this;
    }

    public String datePattern() {
        return datePattern;
    }

    public DateTimeFormatter dateTimeFormatter() {
        return dateTimeFormatter;
    }

    public boolean preferNullOverDefault() {
        return preferNullOverDefault;
    }

    /**
     * the number of skipped rows
     *
     * @return n rows skipped
     */
    public int skip() {
        return skip;
    }

    public boolean ignoreHiddenSheets() {
        return ignoreHiddenSheets;
    }

    private PoijiOptions setIgnoreHiddenSheets(boolean ignoreHiddenSheets) {
        this.ignoreHiddenSheets = ignoreHiddenSheets;
        return this;
    }

    public boolean trimCellValue() {
        return trimCellValue;
    }

    public PoijiOptions setTrimCellValue(boolean trimCellValue) {
        this.trimCellValue = trimCellValue;
        return this;
    }

    private PoijiOptions setSheetIndex(int sheetIndex) {
        this.sheetName = null;
        this.sheetIndex = sheetIndex;
        return this;
    }

    public int sheetIndex() {
        return sheetIndex;
    }

    private PoijiOptions setSheetName(String sheetName) {
        this.sheetIndex = 0;
        this.sheetName = sheetName;
        return this;
    }

    public String sheetName(){
        return sheetName;
    }

    public String getDateRegex() {
        return dateRegex;
    }

    private PoijiOptions setDateRegex(String dateRegex) {
        this.dateRegex = dateRegex;
        return this;
    }

    public boolean getDateLenient() {
        return dateLenient;
    }

    private PoijiOptions setDateLenient(boolean dateLenient) {
        this.dateLenient = dateLenient;
        return this;
    }

    public static class PoijiOptionsBuilder {

        private int skip = 1;
        private int sheetIndex;
        private String sheetName;
        private String password;
        private String dateRegex;
        private boolean dateLenient;
        private boolean trimCellValue;
        private boolean ignoreHiddenSheets;
        private boolean preferNullOverDefault;
        private String datePattern = DEFAULT_DATE_PATTERN;
        private DateTimeFormatter dateTimeFormatter = DEFAULT_DATE_TIME_FORMATTER;

        private PoijiOptionsBuilder() {
        }

        private PoijiOptionsBuilder(int skip) {
            this.skip = skip;
        }

        public PoijiOptions build() {
            final PoijiOptions poijiOptions = new PoijiOptions()
                    .setSkip(skip)
                    .setPassword(password)
                    .setPreferNullOverDefault(preferNullOverDefault)
                    .setDatePattern(datePattern)
                    .setDateTimeFormatter(dateTimeFormatter)
                    .setIgnoreHiddenSheets(ignoreHiddenSheets)
                    .setTrimCellValue(trimCellValue)
                    .setDateRegex(dateRegex)
                    .setDateLenient(dateLenient);
            if (sheetName != null) {
                poijiOptions.setSheetName(sheetName);
            } else {
                poijiOptions.setSheetIndex(sheetIndex);
            }
            return poijiOptions;
        }

        public static PoijiOptionsBuilder settings() {
            return new PoijiOptionsBuilder();
        }

        /**
         * set a date time formatter, default date time formatter is "dd/M/yyyy"
         * for java.time.LocalDate
         *
         * @param dateTimeFormatter date time formatter
         * @return this
         */
        public PoijiOptionsBuilder dateTimeFormatter(DateTimeFormatter dateTimeFormatter) {
            this.dateTimeFormatter = dateTimeFormatter;
            return this;
        }

        /**
         * set date pattern, default date format is "dd/M/yyyy" for
         * java.util.Date
         *
         * @param datePattern date time formatter
         * @return this
         */
        public PoijiOptionsBuilder datePattern(String datePattern) {
            this.datePattern = datePattern;
            return this;
        }

        /**
         * set whether or not to use null instead of default values for Integer,
         * Double, Float, Long, String and java.util.Date types.
         *
         * @param preferNullOverDefault boolean
         * @return this
         */
        public PoijiOptionsBuilder preferNullOverDefault(boolean preferNullOverDefault) {
            this.preferNullOverDefault = preferNullOverDefault;
            return this;
        }

        /**
         * skip number of row
         *
         * @param skip number
         * @return this
         */
        public PoijiOptionsBuilder skip(int skip) {
            this.skip = skip;
            return this;
        }

        /**
         * set sheet index, default is 0
         *
         * @param sheetIndex number
         * @return this
         */
        public PoijiOptionsBuilder sheetIndex(int sheetIndex) {
            if (sheetIndex < 0) {
                throw new PoijiException("Sheet index must be greater than or equal to 0");
            }
            this.sheetIndex = sheetIndex;
            this.sheetName = null;
            return this;
        }

        public PoijiOptionsBuilder sheetName(String sheetName) {
            if (sheetName == null) {
                throw new PoijiException("Sheet name cannot be null");
            }
            if (sheetName.isEmpty()) {
                throw new PoijiException("Sheet name cannot be empty");
            }
            if (sheetName.length() > 31) {
                throw new PoijiException("Sheet name length cannot exceed 31 characters");
            }
            final String[] invalidChars = { ":", "\\", "/", "?", "*", "[", "]" };
            if (Arrays.stream(invalidChars).anyMatch(sheetName::contains)) {
                throw new PoijiException("Sheet name cannot contains any of characters: " + String.join(" ", invalidChars));
            }
            this.sheetName = sheetName;
            this.sheetIndex = 0;
            return this;
        }

        /**
         * Skip the n rows of the excel data. Default is 1
         *
         * @param skip ignored row number
         * @return builder itself
         */
        public static PoijiOptionsBuilder settings(int skip) {
            return new PoijiOptionsBuilder(skip);
        }

        /**
         * set password for encrypted excel file, Default is null
         *
         * @param password
         * @return this
         */
        public PoijiOptionsBuilder password(String password) {
            this.password = password;
            return this;
        }

        /**
         * Ignore hidden sheets
         *
         * @param ignoreHiddenSheets whether or not to ignore any hidden sheets
         * in the work book.
         * @return this
         */
        public PoijiOptionsBuilder ignoreHiddenSheets(boolean ignoreHiddenSheets) {
            this.ignoreHiddenSheets = ignoreHiddenSheets;
            return this;
        }

        /**
         * Trim cell value
         *
         * @param trimCellValue trim the cell value before processing work book.
         * @return this
         */
        public PoijiOptionsBuilder trimCellValue(boolean trimCellValue) {
            this.trimCellValue = trimCellValue;
            return this;
        }

        /**
         * Date regex, if would like to specify a regex patter the date must be
         * in, e.g.\\d{2}/\\d{1}/\\d{4}.
         *
         * @param dateRegex date regex pattern
         * @return this
         */
        public PoijiOptionsBuilder dateRegex(String dateRegex) {
            this.dateRegex = dateRegex;
            return this;
        }

        /**
         * If the simple date format is lenient, use to
         * set how strict the date formatting must be, defaults to lenient false.
         * It works only for java.util.Date.
         *
         * @param dateLenient true or false
         * @return this
         */
        public PoijiOptionsBuilder dateLenient(boolean dateLenient) {
            this.dateLenient = dateLenient;
            return this;
        }

    }
}

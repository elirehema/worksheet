package org.library.worksheet.cellstyles;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.HashMap;
import java.util.Map;

public class ECellStyle {
    private Cell cell;
    private Sheet sheet;

    /**
     * Empty constructor
     **/
    public ECellStyle() {
    }

    public ECellStyle(Sheet sheet) {
        this.sheet = sheet;
    }

    public ECellStyle(Cell cell) {
        this.cell = cell;
    }

    public ECellStyle(Sheet sheet, Cell cell) {
        this.cell = cell;
        this.sheet = sheet;
    }

    public ECellStyle withCell(Cell cell) {
        this.cell = cell;
        return this;
    }

    public ECellStyle withCellAndSheet(Cell cell, Sheet sheet) {
        this.cell = cell;
        this.sheet = sheet;
        return this;
    }

    /**
     * Set Bottom Styles
     **/
    public CellStyle getCellStyle() {

        CellStyle cellStyle = this.cell.getCellStyle();
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBorderLeft(XSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_DOTTED);

        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
        cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cellStyle.setLeftBorderColor(IndexedColors.WHITE.getIndex());
        cellStyle.setBottomBorderColor(IndexedColors.WHITE.getIndex());
        this.cell.setCellStyle(cellStyle);
        return cellStyle;

    }

    public CellStyle getCellStyles(Cell cell) {
        this.cell = cell;
        CellStyle cellStyle = cell.getCellStyle();
        cellStyle.setFont(setTextFonts(cellStyle));
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBottomBorderColor(IndexedColors.LIGHT_GREEN.getIndex());
        cellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

        return cellStyle;
    }

    public ECellStyle setDefaultCellStyle() {
        CellStyle cellStyle = cell.getCellStyle();
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBorderLeft(XSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_DOTTED);
        cellStyle.setLeftBorderColor(IndexedColors.WHITE.getIndex());
        cellStyle.setBottomBorderColor(IndexedColors.WHITE.getIndex());
        cell.setCellStyle(cellStyle);
        return this;
    }


    /**
     * Set Excel sheet default BackgroundStyle.
     **/
    public ECellStyle setDefaultBackgroundStyle() {
        CellStyle backgroundStyle = getSheetCellStyle();
        backgroundStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
        backgroundStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        this.cell.setCellStyle(backgroundStyle);
        return this;
    }


    public ECellStyle setDefaultHeaderBackground() {
        CellStyle cellStyle = getSheetCellStyle();
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBottomBorderColor(IndexedColors.LIGHT_GREEN.getIndex());
        cellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        this.cell.setCellStyle(cellStyle);
        setHeaderFonts(cellStyle);
        return this;
    }


    /**
     * Apply Header fonts
     **/
    private Font setHeaderFonts(CellStyle cellStyle) {
        Font font = getSheetTextFont();
        font.setBoldweight((short) 12);
        font.setFontName("Operator Mono Medium");
        font.setFontHeightInPoints((short) 15);
        font.setColor(IndexedColors.WHITE.getIndex());
        cellStyle.setFont(font);
        return font;
    }

    /**
     * Apply fonts
     **/
    private Font setTextFonts(CellStyle cellStyle) {
        Font font = getSheetTextFont();
        font.setFontName("Operator Mono Medium");
        font.setColor(IndexedColors.LIGHT_GREEN.getIndex());
        cellStyle.setFont(font);
        return font;
    }

    /**
     * Set sheet text font styles
     **/
    private Font getSheetTextFont() {
        Font font = this.sheet.getWorkbook().createFont();
        if (font.equals(null)) {
            font = this.cell.getSheet().getWorkbook().createFont();
        }
        return font;
    }

    /**
     * Get CellStyle from Sheet or Cell
     **/
    private CellStyle getSheetCellStyle() {
        CellStyle cellStyle = this.sheet.getWorkbook().createCellStyle();
        if (cellStyle == null) {
            cellStyle = this.cell.getCellStyle();
        }
        return cellStyle;
    }
    public static Map<CellEnum, CellStyle> createStyles(Workbook wb){
        Map<CellEnum, CellStyle> styles = new HashMap<>();

        CellStyle style;

        Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short)18);
        titleFont.setColor(IndexedColors.WHITE.getIndex());
        titleFont.setBoldweight((short) 18);

        Font monthFont = wb.createFont();
        monthFont.setFontName("Operator Mono Book");
        monthFont.setFontHeightInPoints((short)11);
        monthFont.setColor(IndexedColors.WHITE.getIndex());

        Font headerFont = wb.createFont();
        headerFont.setFontName("Operator Mono Medium");
        headerFont.setFontHeightInPoints((short) 13);


        /**Default Header, Title and Cell **/
        /**@param HEADER **/

        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setFont(titleFont);
        style.setVerticalAlignment(CellStyle.ALIGN_CENTER);
        styles.put(CellEnum.DEFAULT_TITLE, style);


        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.ALIGN_CENTER);
        style.setWrapText(true);
        style.setFont(headerFont);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        styles.put(CellEnum.DEFAULT_HEADER, style);

        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setWrapText(true);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        styles.put(CellEnum.DEFAULT_CELL, style);


        /** Dashed Cell styles **/



        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.ALIGN_CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put(CellEnum.FORMULA_1, style);

        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.ALIGN_CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put(CellEnum.FORMULA_2, style);

        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.ALIGN_CENTER);
        style.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put(CellEnum.FORMULA_3, style);


        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        titleFont.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(titleFont);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.WHITE.getIndex());
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.WHITE.getIndex());
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setTopBorderColor(IndexedColors.WHITE.getIndex());
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.WHITE.getIndex());
        style.setVerticalAlignment(CellStyle.ALIGN_CENTER);
        style.setFillForegroundColor(IndexedColors.TEAL.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        styles.put(CellEnum.TEAL_TITLE, style);


        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.ALIGN_CENTER);
        style.setWrapText(true);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(headerFont);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.WHITE.getIndex());
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.WHITE.getIndex());
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setTopBorderColor(IndexedColors.WHITE.getIndex());
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.WHITE.getIndex());
        style.setFillForegroundColor(IndexedColors.TEAL.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        styles.put(CellEnum.TEAL_HEADER, style);

        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setWrapText(true);
        monthFont.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(monthFont);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.WHITE.getIndex());
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.WHITE.getIndex());
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setTopBorderColor(IndexedColors.WHITE.getIndex());
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.WHITE.getIndex());

        style.setFillForegroundColor(IndexedColors.TEAL.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        styles.put(CellEnum.TEAL_CELL, style);


        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.ALIGN_CENTER);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put(CellEnum.LINE_1, style);

        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.ALIGN_CENTER);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put(CellEnum.LINE_2, style);




        return styles;
    }

}

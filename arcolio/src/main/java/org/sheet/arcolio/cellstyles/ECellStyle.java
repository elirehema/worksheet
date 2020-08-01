package org.sheet.arcolio.cellstyles;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

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

}

package org.library.worksheet;

import android.content.Context;

import androidx.annotation.Keep;

import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.library.worksheet.cellstyles.ECellStyle;
import org.library.worksheet.cellstyles.CellEnum;
import org.library.worksheet.notifier.Toaster;
import org.library.worksheet.storage.ExternalStorage;
import org.library.worksheet.cellstyles.EWorkBook;

import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

@Keep
public class ExcelBookImpl {
    private Logger logger = Logger.getLogger(ExcelBookImpl.class.getName());
    private ECellStyle eCellStyle = null;
    private EWorkBook eWorkBook = new EWorkBook();
    private Workbook workbook;
    Map<CellEnum, CellStyle> styles = null;
    private Context context;
    private Cell cell = null;
    Toaster toaster;
    private ExternalStorage externalStorage;

    public ExcelBookImpl(Context context) {
        this.context = context;
        externalStorage = new ExternalStorage(context);
        externalStorage.createExternalDirectory();
        toaster = new Toaster();
    }

    /**
     * Write an Excel File with single Sheet
     **/
    @Keep
    public void ExcelSheet(List<?> objectList, String excelFilePath, CellEnum HeaderCellStyle, CellEnum TitleCellStyles, CellEnum DataCellStyle) throws IOException {
        Workbook workbook = eWorkBook.getDefaultExcelWorkbook(appendFileExtensionFormatIfNotProvided(excelFilePath));
        styles = eCellStyle.createStyles(workbook);
        Sheet sheet = workbook.createSheet(excelFilePath.toLowerCase());
        populateSingleSheetWithData(sheet, objectList, HeaderCellStyle, TitleCellStyles, DataCellStyle);

        try {
            externalStorage.writeFileToExternalStorage(appendFileExtensionFormatIfNotProvided(excelFilePath), workbook);
        } catch (Exception e) {
            logger.info("File write failed: " + e.toString());
        }

    }

    @SuppressWarnings("unchecked")
    public void ExcelSheets(
            List<Map<String, List<?>>> objs, String excelFilePath, CellEnum HeaderCellStyle,
            CellEnum TitleCellStyles, CellEnum DataCellStyle) throws IOException {
        workbook = eWorkBook.getDefaultExcelWorkbook(appendFileExtensionFormatIfNotProvided(excelFilePath));
        styles = eCellStyle.createStyles(workbook);

        for (Map<String, List<?>> objectMap : objs) {
            Iterator iterator = objectMap.entrySet().iterator();
            while (iterator.hasNext()) {
                Map.Entry entry = (Map.Entry) iterator.next();
                Sheet sheet = workbook.createSheet(entry.getKey().toString());
                List<Object> objects = (List<Object>) entry.getValue();
                populateSingleSheetWithData(sheet, objects, HeaderCellStyle, TitleCellStyles, DataCellStyle);
            }

        }

        try {
            externalStorage.writeFileToExternalStorage(appendFileExtensionFormatIfNotProvided(excelFilePath), workbook);
        } catch (Exception e) {
            logger.info("File write failed: " + e.toString());
        }

    }

    private void createHeaderRow(Object o, Sheet sheet, CellEnum HeaderCellStyle,
                                 CellEnum TitleCellStyles) {
        Cell dataCell, indexcells = null;
        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setFitHeight((short) 1);
        printSetup.setFitWidth((short) 1);
        printSetup.setLandscape(true);

        /**Set Page Number on Footer **/
        Footer ft = sheet.getFooter();
        ft.setRight("Page " + HeaderFooter.page() + " of " + HeaderFooter.numPages());

        sheet.setDefaultColumnWidth(16);
        sheet.setFitToPage(true);
        sheet.setAutobreaks(true);
        sheet.setHorizontallyCenter(true);
        sheet.removeMergedRegion(0);

        Row titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(45);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue(sheet.getSheetName().substring(0, 1).toUpperCase() + sheet.getSheetName().substring(1));
        titleCell.setCellStyle(styles.get(TitleCellStyles == null ? CellEnum.DEFAULT_TITLE : TitleCellStyles));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, o.getClass().getDeclaredFields().length));

        Row headerRow = sheet.createRow(1);
        headerRow.setHeightInPoints(35);
        indexcells = headerRow.createCell(0);
        indexcells.setCellStyle(styles.get(HeaderCellStyle == null ? CellEnum.DEFAULT_CELL : HeaderCellStyle));
        indexcells.setCellValue("#");
        int index = 0;
        Field[] fields = o.getClass().getDeclaredFields();
        for (int i = 0; i < getFieldNames(fields).size(); i++) {
            dataCell = headerRow.createCell(++index);
            sheet.setColumnWidth(i, (fields[i].getName().length() + 12) * 256);
            dataCell.setCellStyle(styles.get(HeaderCellStyle == null ? CellEnum.DEFAULT_CELL : HeaderCellStyle));
            dataCell.setCellValue(fields[i].getName().toUpperCase());
        }
        index = 0;

    }

    /**
     * Write data to created excel sheet
     * Create initial column named #
     * Give {@param rowNumber} to initial column
     **/

    private void writeExcelSheetBook(Object obj, Integer rowNumber, Row row, CellEnum DataCellStyle) {
        Cell cell = row.createCell(0);
        cell.setCellValue(rowNumber.toString());
        cell.setCellStyle(styles.get(DataCellStyle == null ? CellEnum.DEFAULT_CELL : DataCellStyle));
        Field[] fields = obj.getClass().getDeclaredFields();
        int rowCount = 0;
        for (Field f : fields) {
            f.setAccessible(true);
            cell = row.createCell(++rowCount);
            try {
                Field field = obj.getClass().getDeclaredField(f.getName().toString());
                field.setAccessible(true);
                Object value = field.get(obj);
                cell.setCellStyle(styles.get(DataCellStyle == null ? CellEnum.DEFAULT_CELL : DataCellStyle));
                cell.setCellValue(value.toString());
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

    }

    private static List<String> getFieldNames(Field[] fields) {
        List<String> fieldNames = new ArrayList<>();
        for (Field field : fields) {
            fieldNames.add(field.getName());
        }

        return fieldNames;
    }


    /**
     * Get default excel format
     **/
    private String appendFileExtensionFormatIfNotProvided(String extension) {
        String filename = null;
        if (extension.endsWith(".xls")) {
            filename = extension;
        } else if (extension.endsWith(".xlsx")) {
            filename = extension;
        } else {
            filename = extension.concat(".xls");
        }

        return filename;
    }

    private void populateSingleSheetWithData(Sheet sheet, List<?> objects, CellEnum HeaderCellStyle, CellEnum TitleCellStyles, CellEnum DataCellStyles) {
        int rowCount = 0;
        for (Object o : objects) {
            createHeaderRow(o, sheet, HeaderCellStyle, TitleCellStyles);
            Row row = sheet.createRow(++rowCount);
            try {
                writeExcelSheetBook(o, rowCount, row, DataCellStyles);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        rowCount = 0;
    }
}
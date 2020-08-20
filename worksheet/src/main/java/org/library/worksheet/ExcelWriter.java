package org.library.worksheet;

import android.content.Context;

import androidx.annotation.Keep;

import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.library.worksheet.cellstyles.ECellStyle;
import org.library.worksheet.notifier.Toaster;
import org.library.worksheet.storage.EMedia;
import org.library.worksheet.storage.ExternalStorage;
import org.library.worksheet.workbook.EWorkBook;

import java.io.IOException;
import java.lang.reflect.Field;
import java.util.List;
import java.util.logging.Logger;

@Keep
public class ExcelWriter {
    private Logger logger = Logger.getLogger(ExcelWriter.class.getName());
    private ECellStyle eCellStyle = null;
    private EWorkBook eWorkBook = new EWorkBook();
    private Context context;
    private Cell cell = null;
    Toaster toaster;
    private ExternalStorage externalStorage;

    public ExcelWriter(Context context) {
        this.context = context;
        externalStorage = new ExternalStorage(context);
        externalStorage.createExternalDirectory();
        toaster = new Toaster();
    }

    /**
     * Write an Excel File with single Sheet
     **/
    public void writeExcel(List<?> objectList, String excelFilePath) throws IOException {
        if (objectList == null) {
            return;
        }
        Workbook workbook = eWorkBook.createWorkBook(excelFilePath);
        Sheet sheet = workbook.createSheet(excelFilePath.toLowerCase());
        toaster.showToastMessage(this.context, excelFilePath + EMedia.MESSAGE_CREATE_EXCEL_SHEET);
        int rowCount = 0;
        for (Object object : objectList) {
            Row row = sheet.createRow(++rowCount);
            try {
                createHeaderRow(object, sheet);
                writeExcelSheetBook(object, rowCount, row);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

        try {
            externalStorage.writeFileToExternalStorage(excelFilePath, workbook);
        } catch (Exception e) {
            logger.info("File write failed: " + e.toString());
        }
    }

    public void WriteMultipleSheetExcel(List<?> objects, String excelFilePath) throws IOException {
        Workbook workbook = eWorkBook.createWorkBook(excelFilePath);
        int rowCount = 0;
        for (Object obj : objects) {
            logger.info("Logger Info: " + obj.toString());
            Sheet sheet = workbook.createSheet(excelFilePath.toLowerCase());
            try {
                createHeaderRow(obj, sheet);
                Row row = sheet.createRow(++rowCount);
            } catch (Exception e) {
                e.printStackTrace();
            }
            Field[] fields = obj.getClass().getDeclaredFields();

            /**
            for (Field f : fields) {
                try {
                    Field field =  obj.getClass().getDeclaredField(f.getName());
                    field.setAccessible(true);
                    Object object = field.get(obj);



                    if (object instanceof Collection) {
                        //logger.info("Collection Object: " + object.toString());
                        /**for (Object obj: ((Collection<Object>) object)) {

                         try {
                         writeExcelSheetBook(obj, rowCount, row);
                         } catch (Exception e) {
                         e.printStackTrace();
                         }
                         }


                        return;
                    }else {
                        logger.info("String Object: " + object.toString());
                    }

                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
            **/
        }
        try {
            externalStorage.writeFileToExternalStorage(excelFilePath, workbook);
        } catch (Exception e) {
            logger.info("File write failed: " + e.toString());
        }
    }


    /**
     * Create Page Header and Footer
     * Style Header and Footer contents
     **/
    private void createHeaderRow(Object o, Sheet sheet) {
        if (o == null) {
            return;
        }
        PrintSetup printSetup = sheet.getPrintSetup();
        sheet.setAutobreaks(true);
        printSetup.setFitHeight((short) 1);
        printSetup.setFitWidth((short) 1);


        Footer ft = sheet.getFooter();
        ft.setRight("Page " + HeaderFooter.page() + " of " + HeaderFooter.numPages());

        Row row = sheet.createRow(0);
        sheet.setDefaultColumnWidth(15);
        sheet.setFitToPage(true);

        cell = row.createCell(0);
        eCellStyle = new ECellStyle(sheet, cell);
        eCellStyle.setDefaultHeaderBackground();
        cell.setCellValue("#");

        Field[] fields = o.getClass().getDeclaredFields();
        int index = 0;
        for (Field f : fields) {
            cell = row.createCell(++index);
            eCellStyle = new ECellStyle(sheet, cell);
            eCellStyle.setDefaultHeaderBackground();
            cell.setCellValue(f.getName().toUpperCase());
        }

    }

    private void writeExcelSheetBook(Object obj, Integer rowNumber, Row row) {
        if (obj == null) {
            return;
        }
        cell = row.createCell(0);
        eCellStyle = new ECellStyle(cell);
        cell.setCellValue(rowNumber.toString());
        cell.setCellStyle(eCellStyle.getCellStyle());
        Field[] fields = obj.getClass().getDeclaredFields();
        int rowCount = 0;
        for (Field f : fields) {
            f.setAccessible(true);
            cell = row.createCell(++rowCount);
            eCellStyle = new ECellStyle(cell);
            try {
                Field field = obj.getClass().getDeclaredField(f.getName());
                field.setAccessible(true);
                Object value = field.get(obj);
                cell.setCellStyle(eCellStyle.getCellStyle());
                assert value != null;
                cell.setCellValue("" + value.toString());
            } catch (Exception e) {
                logger.info(e.getMessage());
                e.printStackTrace();
            }
        }

    }


}
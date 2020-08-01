package org.sheet.arcolio;

import android.content.Context;
import android.os.Build;
import android.os.Environment;
import android.util.Log;

import androidx.annotation.Keep;
import androidx.annotation.RequiresApi;

import java.io.File;
import java.lang.Class;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.sheet.arcolio.cellstyles.ECellStyle;
import org.sheet.arcolio.storage.EMedia;
import org.sheet.arcolio.storage.ExternalStorage;
import org.sheet.arcolio.workbook.EWorkBook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.Collection;
import java.util.List;
import java.util.logging.Logger;

@Keep
public class ExcelWriter {
    private static final Logger LOGGER = Logger.getLogger(ExcelWriter.class.getName());
    private ECellStyle eCellStyle = null;
    private EWorkBook eWorkBook = new EWorkBook();
    private ExternalStorage externalStorage = null;
    private Context context;
    public ExcelWriter(Context context){
        externalStorage = new ExternalStorage(context);
    }

    /**
     * Write an Excel File with single Sheet
     **/
    public void writeExcel(List<?> objectList, String excelFilePath) throws IOException {
        Workbook workbook = eWorkBook.createWorkBook(excelFilePath);
        Sheet sheet = workbook.createSheet(excelFilePath.toLowerCase());
        externalStorage.createExternalDirectory();
        int rowCount = 0;
        for (Object object : objectList) {
            createHeaderRow(object, sheet);
            Row row = sheet.createRow(++rowCount);
            try {
                writeExcelSheetBook(object, rowCount, row);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        try {
            File folder = Environment.getExternalStoragePublicDirectory(EMedia.DEFAULT_EXTERNAL_FILE_DIRECTORY);
            File  file = new File(folder, excelFilePath);
            FileOutputStream outputStream = new FileOutputStream(file);
            workbook.write(outputStream);
            outputStream.close();
        }catch (Exception e){
            Log.e("EXCEPTION", "File write failed: " +e.toString());
        }
    }

    public void writeMultipleSheetExcel(List<?> languages, String excelFilePath) throws IOException {
        Workbook workbook = eWorkBook.createWorkBook(excelFilePath);
        for (Object parentObject : languages) {
            System.out.println(parentObject.getClass().getDeclaredFields());
            int rowCount = 0;
            Field[] fields = parentObject.getClass().getDeclaredFields();
            for (Field f : fields) {

                try {
                    Field field = parentObject.getClass().getDeclaredField(f.getName().toString());

                    field.setAccessible(true);
                    Object object = field.get(parentObject);

                    if (object instanceof Collection) {
                        // System.out.println("Collection Instance" + object.toString());
                    }
                    if (object instanceof String) {
                        //  System.out.println("String Instance:  " + object.toString());
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
           /* for (Object obj : parentObject.getBooks()) {
                createHeaderRow(obj, sheet);
                Row row = sheet.createRow(++rowCount);
                try {
                    writeBook(obj, rowCount, row);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }*/
            rowCount = 0;
        }
        try {
            FileOutputStream outputStream = new FileOutputStream(excelFilePath);
            workbook.write(outputStream);
        }catch (Exception e){
            Log.e("EXCEPTION", "File write failed: " +e.toString());
        }
    }

    private void createHeaderRow(Object o, Sheet sheet) {
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        Cell cell =  null;
        PrintSetup printSetup = sheet.getPrintSetup();

        /**Set Page Number on Footer **/
        Footer ft = sheet.getFooter();
        ft.setRight("Page " + HeaderFooter.page() + " of " + HeaderFooter.numPages());

        /**Set Print Format**/
        sheet.setAutobreaks(true);
        printSetup.setFitHeight((short) 1);
        printSetup.setFitWidth((short) 1);


        Row row = sheet.createRow(0);
        sheet.setDefaultColumnWidth(15);
        sheet.setFitToPage(true);

        cell = row.createCell(0);
        eCellStyle = new ECellStyle(sheet, cell);
        eCellStyle.setDefaultHeaderBackground();
        cell.setCellValue("#");
        int index = 0;
        for (Field f : o.getClass().getDeclaredFields()) {
            cell = row.createCell(++index);
            eCellStyle = new ECellStyle(sheet, cell);
            eCellStyle
                    .setDefaultHeaderBackground();
            cell.setCellValue(f.getName().toUpperCase());
        }

    }

    private void writeExcelSheetBook(Object obj, Integer rowNumber, Row row) {
        Cell cell = row.createCell(0);
        eCellStyle = new ECellStyle(cell);
        cell.setCellValue(rowNumber.toString());
        Field[] fields = obj.getClass().getDeclaredFields();
        int rowCount = 0;
        for (Field f : fields) {
            f.setAccessible(true);
            cell = row.createCell(++rowCount);
            eCellStyle = new ECellStyle(cell);
            try {
                Field field = obj.getClass().getDeclaredField(f.getName().toString());
                field.setAccessible(true);
                Object value = field.get(obj);
                cell.setCellStyle(eCellStyle.getCellStyle());
                cell.setCellValue(value.toString());
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

    }


}
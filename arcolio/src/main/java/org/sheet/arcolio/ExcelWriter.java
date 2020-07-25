package org.sheet.arcolio;

import android.os.Build;

import androidx.annotation.Keep;
import androidx.annotation.RequiresApi;
import java.lang.Class;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.sheet.arcolio.cellstyles.CellStylesClass;
import org.sheet.arcolio.workbook.WorkBookHelper;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
@Keep
public class ExcelWriter {

    /*Initiate WorkBookHelper class*/
    private WorkBookHelper workBookHelper = new WorkBookHelper();

    /*Initiate CellStylesClass */
    private CellStylesClass cellStylesClass = new CellStylesClass();


    @RequiresApi(api = Build.VERSION_CODES.KITKAT)
    @Keep
    public void writeExcelSheet(List<Class> objectList, String excelFilePath) throws IOException {
       Workbook  workbook = workBookHelper.createWorkBook(excelFilePath);
        Sheet sheet = workbook.createSheet();
        int rowCount = 0;
       for (Class c: objectList) {
           Row row = sheet.createRow(++rowCount);
           writeBook(c,rowCount,row);
       }

       try(FileOutputStream outputStream = new FileOutputStream(excelFilePath)){
           workbook.write(outputStream);
        }
    }

    @Keep
    public void writeMultipleExcelSheet(){

    }

    @Keep
    private void createHeaderAndRow(){

    }

    @Keep
    private void writeBook(Class clss, Integer rowNumber, Row row){
        Cell cell = row.createCell(0);
        cell.setCellValue(rowNumber);
        cell.setCellStyle(cellStylesClass.getBackgroundStyle(cell));


        Field[] field = clss.getDeclaredFields();
        
        for (int i = 1; i<field.length; i++){
            cell=  row.createCell(i);
            cell.setCellValue("Name");
            cell.setCellStyle(cellStylesClass.getCellStyle(cell));

        }

    }


    private CellStyle createCellStyle(Cell cell){

        CellStyle cellStyle = cell.getCellStyle();


        Font font = cell.getSheet().getWorkbook().createFont();
        font.setFontName("Operator Mono Book");
        font.setColor(IndexedColors.GREEN.index);
        cellStyle.setFont(font);

        /** Set Bottom Styles **/
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBorderLeft(XSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_DOTTED);
        cellStyle.setLeftBorderColor(IndexedColors.WHITE.getIndex());
        cellStyle.setBottomBorderColor(IndexedColors.WHITE.getIndex());

        return cellStyle;
    }

}

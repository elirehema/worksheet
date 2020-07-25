package org.sheet.arcolio.workbook;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.sheet.arcolio.util.Constants;
import org.sheet.arcolio.util.ExceptionMessageClass;

import java.io.IOException;

public class WorkBookHelper {
    public Workbook getWorkBook(String excelFilePath) throws IOException{
        Workbook workbook = null;
        if (excelFilePath.endsWith(Constants.EXCEL_2003)){
            workbook = new XSSFWorkbook();
        }else if(excelFilePath.endsWith(Constants.EXCEL_2007)){
            workbook = new HSSFWorkbook();
        }else {
            throw new IllegalArgumentException(ExceptionMessageClass.ILLEGAL_ARGUMENT_EXCEPTION);
        }
        return workbook;
    }

    public Workbook createWorkBook(String excelFilePath) throws IOException {
        Workbook workbook = null;
        if (excelFilePath.endsWith(Constants.EXCEL_2003)) {
            workbook = new XSSFWorkbook();
        } else if (excelFilePath.endsWith(Constants.EXCEL_2007)) {
            workbook = new HSSFWorkbook();
        } else {
            throw new IllegalArgumentException(ExceptionMessageClass.ILLEGAL_ARGUMENT_EXCEPTION);
        }
        return workbook;
    }
}

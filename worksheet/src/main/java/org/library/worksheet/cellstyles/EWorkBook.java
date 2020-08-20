package org.library.worksheet.cellstyles;

import androidx.annotation.Keep;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

@Keep
public class EWorkBook {
    @Keep
    public Workbook createWorkBook(String excelFilePath) throws IOException {
        Workbook workbook = null;
        if (excelFilePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook();
        } else if (excelFilePath.endsWith("xls")) {
            workbook = new HSSFWorkbook();
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }

        return workbook;
    }

    /**
     * Safe
     * Accept any kind of @params excelFilePath name {@params excelFilePath }
     * if files extension is not provided
     * Ths excel file will be give .xls by default
     * **/
    @Keep
    public Workbook getDefaultExcelWorkbook(String excelFilePath) throws IOException {
        Workbook workbook = null;
        if (excelFilePath.endsWith(".xlsx")) {
            workbook = new XSSFWorkbook();
        } else if (excelFilePath.endsWith(".xls")) {
            workbook = new HSSFWorkbook();
        } else {
            excelFilePath.concat(".xls");
            workbook = new XSSFWorkbook();
        }

        return workbook;
    }

}

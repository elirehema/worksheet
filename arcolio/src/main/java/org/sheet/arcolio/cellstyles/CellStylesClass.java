package org.sheet.arcolio.cellstyles;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

/**
 * -- This file created by eli on 7/25/20 for project excel under Package Name: org.sheet.arcolio
 * -- Time: 7/25/2012:46 PM
 * -- Licensed to the Apache Software Foundation (ASF) under one
 * -- or more contributor license agreements. See the NOTICE file
 * -- distributed with this work for additional information
 * -- regarding copyright ownership. The ASF licenses this file
 * -- to you under the Apache License, Version 2.0 (the
 * -- "License"); you may not use this file except in compliance
 * -- with the License. You may obtain a copy of the License at
 * --
 * -- http://www.apache.org/licenses/LICENSE-2.0
 * --
 * -- Unless required by applicable law or agreed to in writing,
 * -- software distributed under the License is distributed on an
 * -- "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * -- KIND, either express or implied. See the License for the
 * -- specific language governing permissions and limitations
 * -- under the License.
 * --
 **/
public class CellStylesClass {

    public CellStyle getCellStyle(Cell cell){
        CellStyle cellStyle = cell.getCellStyle();
        Font font = cell.getSheet().getWorkbook().createFont();
        font.setFontName("Operator Mono Book");
        font.setColor(IndexedColors.GREEN.index);
        cellStyle.setFont(font);
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

    public CellStyle getBackgroundStyle(Cell cell){
        CellStyle backgroundStyle = getCellStyle(cell);
        // backgroundStyle.setFillBackgroundColor(IndexedColors.GREY_2_PERCENT.getIndex());
        backgroundStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
        backgroundStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        return backgroundStyle;
    }
}

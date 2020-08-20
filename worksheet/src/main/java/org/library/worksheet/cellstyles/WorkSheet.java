package org.library.worksheet.cellstyles;


import android.content.Context;

import org.library.worksheet.ExcelBookImpl;

import java.io.IOException;
import java.util.List;
import java.util.Map;

public class WorkSheet {
    private List<?> objects;
    private List<Map<String, List<?>>> map;
    private String path;
    private CellEnum title;
    private CellEnum cell;
    private CellEnum header;
    private CellEnum background;


    public WorkSheet(List<?> objects, List<Map<String, List<?>>> map, String path, CellEnum header, CellEnum title, CellEnum cell, CellEnum background) {
        this.objects = objects;
        this.path = path;
        this.title = title;
        this.cell = cell;
        this.header = header;
        this.background = background;
        this.map = map;
    }
    public WorkSheet(List<Map<String, List<?>>> map, String path, CellEnum header, CellEnum title, CellEnum cell, CellEnum background) {
        this.path = path;
        this.title = title;
        this.cell = cell;
        this.header = header;
        this.background = background;
        this.map = map;
    }



    public CellEnum getHeader() {
        return header;
    }

    public List<?> getObjects() {
        return objects;
    }

    public String getpath() {
        return path;
    }

    public CellEnum gettitle() {
        return title;
    }

    public CellEnum getcell() {
        return cell;
    }

    public CellEnum getBackground() {
        return background;
    }

    public List<Map<String, List<?>>> getMap() {
        return map;
    }

    public static class Builder {
        private List<?> objects;
        List<Map<String, List<?>>> map;
        private String path;
        private CellEnum title;
        private CellEnum cell;
        private CellEnum header;
        private CellEnum background;
        private ExcelBookImpl excelBook;

        public Builder(Context context, String path){
            this.path = path;
            excelBook  = new ExcelBookImpl(context);
        }

        public Builder(List<?> objects, List<Map<String, List<?>>> map,
                       String path, CellEnum header, CellEnum title, CellEnum cell, CellEnum background) {
            this.objects = objects;
            this.path = path;
            this.title = title;
            this.cell = cell;
            this.header = header;
            this.background = background;
            this.map = map;
        }

        public Builder(String path) {
            this.path = path;
        }

        public Builder(List<?> objects, String path) {
            this.objects = objects;
            this.path = path;
        }

        public Builder(String path, List<Map<String, List<?>>> map) {
            this.map = map;
            this.path = path;
        }

        public Builder setMap(List<Map<String, List<?>>> map) {
            this.map = map;
            return this;
        }

        public Builder setData(List<?> objects) {
            this.objects = objects;
            return this;
        }

        public Builder title(CellEnum title) {
            this.title = title;
            return this;
        }

        public Builder cell(CellEnum cell) {
            this.cell = cell;
            return this;
        }

        public Builder header(CellEnum header) {
            this.header = header;
            return this;
        }

        public Builder background(CellEnum background) {
            this.background = background;
            return this;
        }

        public WorkSheet build() {
            return new WorkSheet(objects, null, path, header, title, cell, background);
        }

        public WorkSheet write() throws IOException {
            excelBook.ExcelSheet(this.objects, path, header, title, cell);
            return new WorkSheet(objects, null, path, header, title, cell, background);
        }

        public WorkSheet writes() throws IOException {
            excelBook.ExcelSheets(this.map , path, header, title, cell);
            return new WorkSheet( map, path, header, title, cell, background);
        }


    }
}
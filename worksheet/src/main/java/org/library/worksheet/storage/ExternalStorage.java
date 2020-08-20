package org.library.worksheet.storage;

import android.content.Context;
import android.os.Environment;

import androidx.annotation.Keep;
import androidx.annotation.NonNull;

import org.apache.poi.ss.usermodel.Workbook;
import org.library.worksheet.notifier.Toaster;

import java.io.File;
import java.io.FileOutputStream;

/**
 * -- This file created by eli on 8/1/20 for project excel under Package Name: org.sheet.arcolio.storage
 * -- Time: 8/1/203:07 PM
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
public class ExternalStorage {
    private Context context;
    private Toaster toaster = null;

    public ExternalStorage(Context context) {
        this.context = context;
    }

    /**
     * Create External storage
     * Drop operation if exist
     **/
    public void createExternalDirectory() {
        File f =  this.getExternalFile();
        if (!f.exists()) {
            f.mkdirs();
        }
    }

    /**
     * Check External Media Availability
     * The Media might be mounted to a computer,
     * missing, read-only or some other state
     **/
    private Boolean isExternalMediaStorageAvailable() {
        boolean isMediaAvailable = false;
        boolean isMediaReadable = false;
        String state = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED.equals(state)) {
            //Both read and write operations available
            isMediaAvailable = true;
        } else if (Environment.MEDIA_MOUNTED_READ_ONLY.equals(state)) {
            isMediaAvailable = true;
            isMediaReadable = true;
        } else {
            //SD card not mounted
            isMediaAvailable = false;
        }

        return isMediaAvailable && isMediaReadable;
    }

    /**
     * Write to Defined External Storage
     **/
    @Keep
    public void writeFileToExternalStorage(@NonNull String excelFilePath, @NonNull File folder, @NonNull Workbook workbook) {
        File file = new File(folder, excelFilePath);
        WriteToOutputStream(file, workbook);
    }


    /**
     * Write to Default External Storage Folder
     **/
    @Keep
    public void writeFileToExternalStorage(@NonNull String excelFilePath, @NonNull Workbook workbook) {
        File folder = this.getExternalFile();
        File file = new File(folder, excelFilePath);
        WriteToOutputStream(file, workbook);
    }

    @Keep
    private void WriteToOutputStream(File file, Workbook workbook) {
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(file);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    /** Get default public Directory **/
    @Keep
    private File getExternalFile(){
        return Environment.getExternalStoragePublicDirectory(EMedia.DEFAULT_EXTERNAL_FILE_DIRECTORY);
    }

}

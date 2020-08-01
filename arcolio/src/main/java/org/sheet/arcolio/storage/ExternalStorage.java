package org.sheet.arcolio.storage;

import android.app.Application;
import android.content.Context;
import android.os.Environment;
import android.widget.Toast;

import org.sheet.arcolio.R;
import org.sheet.arcolio.notifier.Toaster;

import java.io.File;
import java.io.FileNotFoundException;
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
    public ExternalStorage(Context context){
        this.context = context;
    }

    /** Create External storage **/
    public void createExternalDirectory(){
        if (isExternalMediaStorageAvailable()) {
            File f = new File(Environment.getExternalStorageDirectory(), EMedia.DEFAULT_EXTERNAL_FILE_DIRECTORY);
            if (!f.exists()) {
                f.mkdirs();
            }
        }else {
            toaster = new Toaster(context, EMedia.ERROR_DEVICE_MEMORY_IS_NOT_AVAILABLE);
        }
    }

    /** Check External Media Availability
     * The Media might be mounted to a computer,
     * missing, read-only or some other state **/
    private Boolean isExternalMediaStorageAvailable(){
        boolean isMediaAvailable = false;
        boolean isMediaReadable = false;
        String state = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED.equals(true)){
            //Both read and write operations available
            isMediaAvailable = true;
        }else if(Environment.MEDIA_MOUNTED_READ_ONLY.equals(true)){
            isMediaAvailable = true;
            isMediaReadable = true;
        }else{
            //SD card not mounted
            isMediaAvailable = false;
        }

        return isMediaAvailable && isMediaReadable;
    }

    /** Write to External Storage **/
    public void writeFileToExternalStorage(){
        File folder = Environment.getExternalStoragePublicDirectory(EMedia.DEFAULT_EXTERNAL_FILE_DIRECTORY);
        File file = new File(folder, "File.xls");
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(file);
            fileOutputStream.write("Elirehema".getBytes());
            fileOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

}

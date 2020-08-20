package org.library.worksheet.notifier;

import android.app.Activity;
import android.content.Context;
import android.view.Gravity;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.TextView;
import android.widget.Toast;

import androidx.annotation.Keep;

import org.library.worksheet.R;

/**
 * -- This file created by eli on 8/1/20 for project excel under Package Name: org.sheet.arcolio.notify
 * -- Time: 8/1/204:38 PM
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
@Keep
public class Toaster {
    Context context;
    String message;

    public Toaster(){}
    public Toaster(Context context, String message) {
        this.context = context;
        this.message = message;
        showToastMessage(context, message);
    }

    public void showToastMessage(Context context, String message) {
        LayoutInflater inflater = ((Activity) context).getLayoutInflater();
        View layout = inflater.inflate(R.layout.custom_toast, (ViewGroup) ((Activity) context).findViewById(R.id.custom_toast));

        TextView text = (TextView) layout.findViewById(R.id.info_message);
        text.setText(message);

        // Toast..
        Toast toast = new Toast(context);
        toast.setMargin(0,0);
        toast.setGravity(Gravity.BOTTOM, 0, 0);
        toast.setDuration(Toast.LENGTH_LONG);
        toast.setView(layout);
        toast.show();
    }
}

package org.arcolio.excel;

import androidx.annotation.RequiresApi;
import androidx.appcompat.app.AppCompatActivity;

import android.os.Build;
import android.os.Bundle;
import android.os.Environment;

import org.arcolio.excel.models.Book;
import org.arcolio.excel.models.Language;
import org.sheet.arcolio.ExcelWriter;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class MainActivity extends AppCompatActivity {

    @RequiresApi(api = Build.VERSION_CODES.KITKAT)
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        String multipleFilePath = "BookList.xls";
        ExcelWriter excelWriter = new ExcelWriter(this.getApplicationContext());
        try {
            excelWriter.writeExcel(getListBook(),multipleFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }

    }




    public List<Book> getListBook(){
        Book book = null;
        Book book2 = new Book("Effective Java", "Joshua Bloch", 36,"Effective Java","overflow.com",
                "Learn Java ","Baron Quinn","M8S4DS3211");
        Book book3 = new Book("Clean Code", "Robert Martin", 42,"Effective Java","overflow.com",
                "Learn Java ","Baron Quinn","M8S4DS3211");
        Book book4 = new Book("Thinking in Java", "Bruce Eckel", 35,"Effective Java","overflow",
                "Learn Java ","Baron Quinn","M8S4DS3211");

        List<Book> listBook = Arrays.asList(book2,book3,book4);
        List<Book>  bookList = new ArrayList<>();
        for (int i=0; i<= 100; i++){
            book = new Book("Head First Java", "Kathy Bruce", 79+i,"Effective Java","overflow",
                    "Learn Java ","Baron Quinn","M8S4DS3211");
            bookList.add(book);
        }
        bookList.add(book4);
        bookList.add(book3);
        bookList.add(book2);


        return bookList;
    }
}

package org.arcolio.excel;

import androidx.annotation.RequiresApi;
import androidx.appcompat.app.AppCompatActivity;

import android.os.Build;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;

import org.arcolio.excel.models.Book;
import org.arcolio.excel.models.Language;
import org.library.worksheet.ExcelBookImpl;
import org.library.worksheet.cellstyles.CellEnum;
import org.library.worksheet.cellstyles.WorkSheet;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class MainActivity extends AppCompatActivity {
    String multipleFilePath = "BookList.xls";

    @RequiresApi(api = Build.VERSION_CODES.KITKAT)
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        Button button = findViewById(R.id.create_excel_sheet);
        final String path  = "ExternalFilePath";
        button.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                WorkSheet workSheet;
                try {
                    workSheet = new WorkSheet.Builder(getApplicationContext(), path)
                            .setSheets(getListOfObject())
                            .writeSheets();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        });


    }




    public static List<Book> getListBook(String bookName) {
        Book book = null;

        Book book2 = new Book(bookName, "Joshua Bloch", 36, "Effective Java", "overflow.com",
                "Learn Java ", "Baron Quinn", "M8S4DS3211");
        Book book3 = new Book(bookName, "Robert Martin", 42, "Effective Java", "overflow.com",
                "Learn Java ", "Baron Quinn", "M8S4DS3211");
        Book book4 = new Book(bookName, "Bruce Eckel", 35, "Effective Java", "overflow",
                "Learn Java ", "Baron Quinn", "M8S4DS3211");

        List<Book> listBook = Arrays.asList(book2, book3, book4);
        List<Book> bookList = new ArrayList<>();
        for (int i = 0; i <= 100; i++) {
            book = new Book(bookName, "Kathy Bruce", 79 + i, "Effective Java", "overflow",
                    "Learn Java ", "Baron Quinn", "M8S4DS3211");
            bookList.add(book);
        }
        bookList.add(book4);
        bookList.add(book3);
        bookList.add(book2);


        return bookList;
    }
    public static List<Map<String, List<?>>> getListOfObject(){
        List<Map<String, List<?>>> map =  new ArrayList<Map<String, List<?>>>();
        Map<String, List<?>> map1 = new HashMap<>();
        map1.put("Kotlin", getListBook("Kotlin Essential"));
        map1.put("Java ", getListBook("Effective Java"));
        map1.put("Python", getListBook("Python for beginners"));
        map1.put("Javascript", getListBook("Javascript for Geeks"));

        map1.put("VueJs", getListBook("VueJs for Geeks"));

        map1.put("Ruby", getListBook("Ruby build the world"));
        map.add(map1);

        return map;
    }

}

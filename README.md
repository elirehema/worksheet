## **Worksheet** 
  ![Android CI](https://github.com/elirehema/arcolios/workflows/Android%20CI/badge.svg) [![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://opensource.org/licenses/Apache-2.0) ![Maven Central](https://img.shields.io/badge/maven--central-v0.0.1-brightgreen)

---

A simplified way of creating a simple Excel Sheets in Android project. It is simple as explained below.

---
The library is written in Java and using [Apache poi](http://poi.apache.org/index.html)
<br><br>
### **How to use in your android project**
- #### **1. Add gradle dependency** 
---

-In your application `build.gradle` (app module) add this 
```java
dependencies{
        .....................................
        ..........................................

       implementation 'com.github.elirehema:worksheet:0.0.1'
```
If you are using gradle below version 0.8.0 make sure you have added maven repository in your `project build.gradle` file
```java
buildscript { 
    repositories { 
        maven { 
            url 'repo1.maven.org/maven2'; 
        } 
    } 
    ...............
    ........................
} 

repositories {
    mavenCentral()
}

```
>As of version 0.8.9, Android Studio supports the Maven Central Repository by default. So to add an external maven dependency all you need to do is edit the module's build.gradle file and insert a line into the dependencies section [Source](https://stackoverflow.com/a/26630422/7098524)

-  #### **2. Data type's**
---
Worksheet  version 0.0.1 support two data types
      
1. ##### **List of Objects**
    E.g `List<Book>` for creating an Excel Book with single Sheet. 
    ![Single excel Sheet](https://raw.githubusercontent.com/elirehema/arcolis/master/img/single_sheet.png)
    
2. ##### **List of Map of String to List of Objects **
    E.g `List<Map<String, List<Books>>>` fo creating an Excel Book with multiple Excel Sheet.
    ![Multiple excel sheets](https://raw.githubusercontent.com/elirehema/arcolis/master/img/Screenshot%20from%202020-08-24%2020-00-18.png)


    
- #### **3. Populate Excel Sheet with data**
---
In your Activity class create an instance of WorkSheet class
```java 
public class MainActivity extends AppCompatActivity {
    private WorkSheet workSheet;
    private Button button;
    ........
    ...................
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        button = findViewById(R.id.create_excel_sheet);
        final String path  = "ExternalFilePath";
        
        button.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                try {
                    workSheet = new WorkSheet.Builder(getApplicationContext(), path)
                            .setSheet(List<Object>)
                            .writeSheet();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        });


    }
}

```
The instance accept to variables application context and extenalFilePath as the name of the file.

With this few steps the excel file with name provided as externalFilePath String will be save in the external directory `worksheets` created by worksheet dependency in your device. 

<br>

- #### **IMPORTANT NOTE**
---
The `WorkSheet.Builder(getApplicationContext(), path)` have two chirld methods for accepting either ``List<Object>`` or `List<Map<String, List<Object>>>`.

- To create Excel book with multiple Single Sheet use:
```java
new WorkSheet.Builder(getApplicationContext(), path)
                            .setSheet(List<Objects>)
                            .writeSheet();
```
- To create an Excel book with mutliple Sheets use:

```java 
new WorkSheet.Builder(getApplicationContext(), path)
                            .setSheets(List<Map<String, List<Object>>>)
                            .writeSheets();
```

- #### **4. Cell Styling**
---
By default cell style create an Excel sheet with black texts and borders. But you can change either title, Header and Data cell styles by using the mentioned CellEnum types.

| Enum Type      |        Source       | Properties    |
| ---------------|:-------------------:|--------------:|
| DEFAULT_HEADER | CellEnum            | No Effect     |
| DEFAULT_TITLE  | CellEnum            | No Effect     |
| DEFAULT_CELL   |   CellEnum          |No Effect     |
| TEAL_HEADER    | CellEnum            |Teal text, bold,  white text |
| TEAL_TITLE     | CellEnum            |Teal text, medium,  white text |
| TEAL_CELL      | CellEnum            | Teal text, normal,  white text |
| FORMULA_1      | CellEnum            | Give it a try  |
| FORMULA_2      |  CellEnum           | Give it a try  |
| FORMULA_3      |  CellEnum           | Give it a try  |
| LINE_2         | CellEnum            |Give it a try  |
| LINE_1         |  CellEnum           |Give it a try  |

<br>

- - ##### **Using Styles**
In your Builder method add `.title(CellEnum_ENUM_TYPE)` to change title cell style or `.header(CellEnum.ENUM_TYPE)` for header cells or `.cell(CellEnum.ENUM_TYPE)` for data cells 


Example 1.

```java 
 new ESheet.Builder(multipleFilePath)
                .header(EType.TEAL_HEADER)
                .cell(EType.LINE_2)
                .title(EType.TEAL_TITLE)
                .setSheet(getListBook("New Sample Book"))
                .writeSheet();
```
![Teal header](https://github.com/elirehema/arcolis/blob/master/img/style_1.png?raw=true)

Example 2.

```java
 eSheet = new ESheet.Builder(multipleFilePath)

                .title(EType.TEAL_HEADER)
                .header(EType.FORMULA_1)
                .cell(EType.DEFAULT_CELL)
                .setData(getListBook("New Sample Book"))
                .write();
```
![Teal header](https://github.com/elirehema/arcolis/blob/master/img/exmple_1.png?raw=true)

Example 3.
```java
   List<Map<String, List<?>> > languageList = getListOfObject();
        String excelFilePath = "NiceJavaBooks";
        String multipleFilePath = "BookList";
        eSheet = new ESheet.Builder(multipleFilePath)

                .title(EType.TEAL_HEADER)
                .header(EType.FORMULA_1)
                .cell(EType.DEFAULT_CELL)
                .setSheets(languageList)
                .writeSheets();
```
<img alt="Gif image" src="https://github.com/elirehema/arcolis/blob/master/img/ezgif.com-video-to-gif.gif?raw=true" width="100%"></img>
- #### **Remember**
---
setSheet(List<Objects>)  goes with  createSheet()
setSheets(List<Map<String, List<Object>>>) goes   with createSheets()

- Repository is open for PR's, Issues, Fork, etc

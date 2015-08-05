[![Build Status](https://travis-ci.org/nvenky/excel-parser.svg)](https://travis-ci.org/nvenky/excel-parser)


# Excel Parser Examples


HSSF - Horrible Spreadsheet Format – not anymore. With few annotations, excel parsing can be done in one line.

We had a requirement in our current project to parse multiple excel sheets and store the information to database. I hope most of the projects involving excel sheet parsing would be doing the same. We built a extensible framework to parse multiple sheets and populate JAVA objects with annotations.

## Usage

This JAR is currently available in [Sonatype maven repository](https://oss.sonatype.org/#nexus-search;quick~excel-parser).

Maven:

````xml
<dependency>
  <groupId>org.javafunk</groupId>
  <artifactId>excel-parser</artifactId>
  <version>1.0</version>
</dependency>
````

Gradle:

````    
compile 'org.javafunk:excel-parser:1.0'
````    

Thanks to [tobyclemson](http://github.com/tobyclemson) for publishing this to Maven repository. 


## Student Information Example

Consider we have an excel sheet with student information.

![Student Information](http://3.bp.blogspot.com/_OAeb_UFifRg/S2WiIfweGiI/AAAAAAAACCA/Sv36FYxed1E/s640/Screenshot.png)

While parsing this excel sheet, we need to populate one “Section” object and multiple “Student” objects related to a Section. You can see that Student information is available in multiple rows whereas the Section details (Year, Section) is available in column B.

### Step 1: Annotate Domain Classes

First we will see the steps to annotate Section object:

````java
@ExcelObject(parseType = ParseType.COLUMN, start = 2, end = 2)
public class Section {
	@ExcelField(position = 2)
	private String year;
    
 	@ExcelField(position = 3)
 	private String section;
 
 	@MappedExcelObject
 	private List <Student> students;
}

````
You can find three different annotation in this class.

* `ExcelObject`: This annotation tells the parser about the parse type (Row or Column), number of objects to create (start, end). Based on the above annotation, Section value should be parsed Columnwise and information can be found in Column 2 (“B”) of the Excelsheet.
* `ExcelField`: This annotation tells the parser to fetch “year” information from Row 2 and “section” information from Row 3.
* `MappedExcelObject`: Apart from Simple datatypes like “Double”,”String”, we might also try to populate complex java objects while parsing. In this case, each section has a list of student information to be parsed from excel sheet. This annotation will help the parser in identifying such fields.

Then, annotate the Student class:

````java
@ExcelObject(parseType = ParseType.ROW, start = 6, end = 8)
public class Student {
    @ExcelField(position = 2)
    private Long roleNumber;

    @ExcelField(position = 3)
    private String name;

    @ExcelField(position = 4)
    private Date dateOfBirth;

    @ExcelField(position = 5)
    private String fatherName;

    @ExcelField(position = 6)
    private String motherName;

    @ExcelField(position = 7)
    private String address;

    @ExcelField(position = 8)
    private Double totalScore;
}
````

* `ExcelObject`: As shown above, this annotation tells parser to parse Rows 6 to 8 (create 3 student objects). NOTE: Optional field “zeroIfNull” , if set to true, will populate Zero to all number fields (Double,Long,Integer) by default if the data is not available in DB.
* `ExcelField`: Student class has 7 values to be parsed and stored in the database. This is denoted in the domain class as annotation.
* `MappedExcelObject`: Student class does not have any complex object, hence this annoation is not used in this domain class.


### Step 2: Invoke Sheet Parser

Once the annotation is done, you have just invoke the parser with the Sheet and the Root class you want to populate.

````java
//Get the sheet using POI API.
String sheetName = "Sheet1";
SheetParser parser = new SheetParser();
InputStream inputStream = getClass().getClassLoader().getResourceAsStream("Student Profile.xls");
Sheet sheet = new HSSFWorkbook(inputStream).getSheet(sheetName);

//Invoke the Sheet parser.
List entityList = parser.createEntity(sheet, sheetName, Section.class);
````

Thats all it requires. Parser would populate all the fields based on the annotation for you.

### Development
* JDK 8
* Run "gradle idea" to setup the project
* Install Lombok plugin
* Enable "Enable annotation processing" as this project uses Lombok library. [Compiler > Annotation Processors > Enable annotation processing: checked ]


### Contributors
* @nvenky
* @cv
* @tobyclemson

# EntitiesToExcel

This project was designed to allow me to quickly dump a list of Hibernate entities to an Excel document. I have only had one use case so far. That use case was dumping some prototype data to an excel sheet. This has allowed for easy debugging a ETL process by dumping to excel instead of into a table. Excel has been easier to organize the data, and test for patterns. I didn't like importing the CSV every time. I used reflections to obtain column names, and data types. 

This will export to multiple sheets if you use a List of Lists, of entities.
      ` E.G List<List<EntityClass>> myEntities;`

The entities are limited to only annotated classes. If I continue development, I would like to work with XML definitions as well. 

I used Apache POI since I generally like the apache API's. It doesn't have a whole lot in place for type conversion, so I only handled a few datatypes to numbers. The rest is all strings. 

I'm using the latest Hibernate (4.3.11.Final as of writing this), but there is no reason it can't work with prior versions. The only hibernate item it needs is the Column annotation class. 


**INSTALLATION:**


Install into local maven.
Clone the repo

` git clone https://github.com/kloverde1985/EntitiesToExcel.git`

Run maven to install into

   ` mvn install`
    
  Add the dependency to your pom.xml **(don't forget to check on the version #)**
```XML
      <dependency>
            <groupId>com.loverde</groupId>
            <artifactId>EntitiesToExcel</artifactId>
            <version>0.12</version>
        </dependency>
```



Source in project:

  Add the Apache POI dependency to your pom.xml
```xml
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi</artifactId>
            <version>3.13</version>
        </dependency>
```     
  You should already have the Hibernate depencies in your project. This project does not configure connections, or DAO's etc.
  

**USE**

**Defining entities**
```Java
  //Define columns in the entity file. Use the @Column annotation, only the name attributer is required. 
    @Column(name = "prod")
    public String getProd() {
    return prod;
    }
 ```   
    
**Output for just one worksheet**
```Java    
//In your business logic.  Obtain a list of Hibernate entities
List<Product> products = productsDAO.list();
//Create ExportToExcel object. Parameters are: List followed by the full path to the output file.
ExportToExcel exportToExcel = new ExportToExcel(products,"products.xlsx");
//Call the export function
exportToExcel.export();
```
**Output for multiple entities to worksheets.**
```Java
//In your business logic.  Obtain a list of Hibernate entities
List<Product> products = productDAO.list();
List<Price> prices = priceDAO.list();
//Create top level list
List<List> productPrices = new ArrayList<>();
productPrices.add(products);
productPrices.add(prices);

//Create ExportToExcel object. Parameters are: List followed by the full path to the output file.
ExportToExcel exportToExcel = new ExportToExcel(productPrices,"products.xlsx");
//Call the export function
exportToExcel.export();
```

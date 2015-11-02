# EntitiesToExcel

This project was designed to allow me to quickly dump a list of Hibernate entities to an Excel document. I have only had one use case so far. That use case was dumping some prototype data to an excel sheet. This has allowed for easy debugging a ETL process by dumping to excel instead of into a table. Excel has been easier to organize the data, and test for patterns. I didn't like importing the CSV every time. I used reflections to obtain column names, and data types. 

This will export to multiple sheets if you use a List of Lists, of entities. E.G List<List<EntityClass>> myEntities;

The entities are limited to only annotated classes. If I continue development, I would like to work with XML definitions as well. 

I used Apache POI since I generally like the apache API's. It doesn't have a whole lot in place for type conversion, so I only handled a few datatypes to numbers. The rest is all strings. 

I'm using the latest Hibernate (4.3.11.Final as of writing this), but there is no reason it can't work with prior versions. The only hibernate item it needs is the Column annotation class. 


INSTALLATION:


From local maven
  From the cloned directory, run maven to install the jar file into your local maven.
    mvn install
    
  Add the dependency to your pom.xml (don't forget to check on the version #)
  
            <dependency>
            <groupId>com.loverde</groupId>
            <artifactId>EntitiesToExcel</artifactId>
            <version>0.12</version>
        </dependency>



Source in project:

  Add the Apache POI dependency to your pom.xml
          <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi</artifactId>
            <version>3.13</version>
        </dependency>
  You should already have the Hibernate depencies in your project. 
  

USE

Defining entities

//Define columns in the entity file. Use the @Column annotation, only the name attributer is required. 
@Column(name = "prod", nullable = false, insertable = true, updatable = true, length = 6)
    public String getProd() {
        return prod;
    }
    Output for just one worksheet
    
//In your business logic.  Obtain a list of Hibernate entities
List<Product> products = productsDAO.list();
//Create ExportToExcel object. List followed by the full path to the output file.
ExportToExcel exportToExcel = new ExportToExcel(products,"products.xlsx");
//Call the export function
exportToExcel.export();

Output for multiple entities to worksheets.

//In your business logic.  Obtain a list of Hibernate entities
List<Product> products = productDAO.list();
List<Price> prices = priceDAO.list();
//Create top level list
List<List> productPrices = new ArrayList<>();
productPrices.add(products);
productPrices.add(prices);
//Create ExportToExcel object
//Create ExportToExcel object. List followed by the full path to the output file.
ExportToExcel exportToExcel = new ExportToExcel(productPrices,"products.xlsx");
//Call the export function
exportToExcel.export();


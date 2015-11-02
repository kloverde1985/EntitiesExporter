# EntitiesToExcel

This project was designed to allow me to quickly dump a list of Hibernate entities to an Excel document. I have only had one use case so far. That use case was dumping some prototype data to an excel sheet. This has allowed for easy debugging a ETL process by dumping to excel instead of into a table. Excel has been easier to organize the data, and test for patterns. I didn't like importing the CSV every time.

The entities are limited to only annotated classes. If I continue development, I would like to work with XML definitions as well. 

I used Apache POI since I generally like the apache API's. It doesn't have a whole lot in place for type conversion, so I only handled a few datatypes to numbers. The rest is all strings. 

I'm using the latest Hibernate (4.3.11.Final as of writing this), but there is no reason it can't work with prior versions. The only hibernate item it needs is the Column annotation class. 

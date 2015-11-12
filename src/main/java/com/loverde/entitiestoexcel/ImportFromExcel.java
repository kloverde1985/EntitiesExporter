package com.loverde.entitiestoexcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.persistence.Entity;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by kevin1 on 11/9/2015.
 */
public class ImportFromExcel {
private Workbook workbook;

    public ImportFromExcel(String fileName) throws IOException {
        workbook = new XSSFWorkbook(fileName);


    }
/*
Completely untested.
TODO: Make this work.
 */
public List importWorkbook()
    {

        Sheet sheet = workbook.getSheet("Tagprice");
        List<Entity> entities = new ArrayList<>();

        for(Row row: sheet){
            List entityItems = new ArrayList();
            for(Cell cell : row)
            {
                entityItems.add(cell.getStringCellValue());
            }
        }
        return entities;
    }

}

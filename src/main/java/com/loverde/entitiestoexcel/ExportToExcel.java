package com.loverde.entitiestoexcel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.persistence.Column;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Kevin Loverde on 11/2/2015.
 * This is a simple library to quickly save Hibernate Entities to excel files.
 * I created this to allow me to get data out of a system quickly for analysis.
 *
 * I have systems that use Hibernate heavily, and occasionally I would need to get
 * the data into a format that I could easily look at and analyze.
 *
 */
public class ExportToExcel {

    private List entityList;
    private String fileName;

    public ExportToExcel() {

    }

    /**
     *
     * @param entityList Accepts either a List<Entity> or List<List<Entity>> . If there is the nested
     *                   List of entities, it will create multiple sheets in the workbook.
     * @param fileName This is the full path for the file we will be saving. It will always save as an XLSX.
     */
    public ExportToExcel(List entityList, String fileName) {
        this.entityList = entityList;
        this.fileName = fileName;

    }


    public List getEntityList() {
        return entityList;
    }

    public void setEntityList(List entityList) {
        this.entityList = entityList;
    }

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    private Workbook saveEntityListToWorkBook(Workbook workbook, List entityList) throws IllegalAccessException, InvocationTargetException {

        Class entityClass = entityList.get(0).getClass();
        Sheet tagpricesSheet = workbook.createSheet(entityClass.getSimpleName());
        int rowIndex = 0;
        int cellIndex = 0;
        Row row = tagpricesSheet.createRow(rowIndex);
        List<Method> objectMethods = new ArrayList<>();


        for (Method method : entityClass.getMethods()) {
            if (method.isAnnotationPresent(Column.class)) {
                objectMethods.add(method);
                Column annotation = method.getAnnotation(Column.class);
                Cell cell = row.createCell(cellIndex);
                cell.setCellValue(annotation.name());
                cell.getCellStyle().setLocked(true);
                cellIndex++;
            }
        }
        rowIndex++;
        cellIndex = 0;
        for (Object object : entityList) {
            row = tagpricesSheet.createRow(rowIndex);
            for (Method method : objectMethods) {
                try {
                    Object value = method.invoke(object, null);
                    if (value != null) {

                        Cell cell = row.createCell(cellIndex);
                        if (method.getReturnType() == short.class || method.getReturnType() == Integer.class || method.getReturnType() == BigDecimal.class) {
                            cell.setCellValue(Double.valueOf(value.toString()));
                            cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                        }else
                        {
                            cell.setCellValue(value.toString());
                        }
                    }else
                    {

                        Cell cell = row.createCell(cellIndex);
                        cell.setCellValue("");

                    }
                    cellIndex++;
                } catch (IllegalAccessException e) {
                    throw e;
                } catch (InvocationTargetException e) {
                    throw e;
                }
            }
            cellIndex = 0;
            rowIndex++;

        }

        return workbook;
    }

    public void export() throws IOException, InvocationTargetException, IllegalAccessException {
        if (entityList == null) return;
        if (entityList.size() < 1) return;

        Workbook workbook = new XSSFWorkbook();
        if (List.class.isAssignableFrom(entityList.get(0).getClass()) ) {

            for(int x=0;x<=entityList.size()-1;x++)
            {
                List innerList = (List) entityList.get(x);
                saveEntityListToWorkBook(workbook,innerList);
            }
        }
        else
        {
            saveEntityListToWorkBook(workbook,entityList);
        }


        try(FileOutputStream fos = new FileOutputStream(fileName);)
        {
            workbook.write(fos);
            fos.close();
        }
    }

}

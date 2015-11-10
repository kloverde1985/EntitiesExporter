package com.loverde.entitiestoexcel;

import javax.persistence.Column;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by kevin1 on 11/10/2015.
 */
public class ExportToCsv {

    private List entityList;
    private String fileName;

    public ExportToCsv(List entityList, String fileName) {
        this.entityList = entityList;
        this.fileName = fileName;
    }

    public void export() throws IOException, InvocationTargetException, IllegalAccessException {
        if (entityList == null) return;
        if (entityList.size() < 1) return;

        File csvFile = new File(fileName);
        saveEntityListCsv(csvFile, entityList);


    }

    private void saveEntityListCsv(File csvFile, List entityList) throws IOException, IllegalAccessException, InvocationTargetException {
        Class entityClass = entityList.get(0).getClass();
        List<Method> objectMethods = new ArrayList<>();

        try (BufferedWriter bufferedWriter = new BufferedWriter(new FileWriter(csvFile))) {
            StringBuilder stringBuilder = new StringBuilder();
            for (Method method : entityClass.getMethods()) {
                if (method.isAnnotationPresent(Column.class)) {
                    objectMethods.add(method);
                    Column annotation = method.getAnnotation(Column.class);
                    stringBuilder.append(annotation.name());
                    stringBuilder.append(",");
                }
            }
            //Get rid of the trailing comma
            stringBuilder.deleteCharAt(stringBuilder.lastIndexOf(","));
            //start a new line
            stringBuilder.append("\n");
            //write to the file
            bufferedWriter.write(stringBuilder.toString());
            //flush out the first row.
            bufferedWriter.flush();
            for (Object object : entityList) {
                //Start with an empty stringBuilder.
                stringBuilder = new StringBuilder();
                for (Method method : objectMethods) {
                    try {
                        Object value = method.invoke(object, null);
                        if (value != null) {

                            if (method.getReturnType() == short.class || method.getReturnType() == Integer.class || method.getReturnType() == BigDecimal.class) {
                                stringBuilder.append(value + ",");
                            } else {
                                String outputValue = value.toString().replaceAll("\"", "\\\"");
                                //If it's a string, we'll want to quote the value.
                                stringBuilder.append("\"" + value + "\",");
                            }
                        } else {
                            stringBuilder.append(",");

                        }
                    } catch (IllegalAccessException e) {
                        throw e;
                    } catch (InvocationTargetException e) {
                        throw e;
                    }
                }
                //Get rid of the trailing comma
                stringBuilder.deleteCharAt(stringBuilder.lastIndexOf(","));
                //start a new line
                stringBuilder.append("\n");
                //write to the file
                bufferedWriter.write(stringBuilder.toString());
                //flush out the row.
                bufferedWriter.flush();
            }
        }
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
}

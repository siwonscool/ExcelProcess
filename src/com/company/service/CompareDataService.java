package com.company.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

public class CompareDataService {
    private List<List<String>> compareDifferentData(List<List<String>> originData, List<List<String>> resultData){
        List<List<String>> comparedData = new ArrayList<>();
        double errorCount = 0;
        //originData.get(i).get(4).equals("전자문서") &&
        for (int i = 0; i < originData.size(); i++) {
            if (!originData.get(i).get(4).equals(resultData.get(i).get(4))){
                comparedData.add(originData.get(i));
                comparedData.add(resultData.get(i));
                errorCount++;
                comparedData.add(new ArrayList<>());
            }
        }
        double errorRate = errorCount / originData.size() * 100;

        System.out.println("오차 데이터 크기 : " + errorCount);
        System.out.println("원본 데이터와 오차율 : " + errorRate + " %");

        return comparedData;
    }

    public void createCompareResultExcel(List<List<String>> originData, List<List<String>> resultData, String path, String filename){
        List<List<String>> compareData = compareDifferentData(originData, resultData);

        try{
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("오차 데이터");

            for (int i = 0; i < compareData.size(); i++) {
                Row row = sheet.createRow(i);
                for (int j = 0; j < compareData.get(i).size(); j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(compareData.get(i).get(j));
                }
            }

            String localFile = path + filename +".xlsx";
            File file = new File(localFile);
            FileOutputStream fileOutputStream = new FileOutputStream(file);
            workbook.write(fileOutputStream);
            workbook.close();

        }catch (FileNotFoundException e){
            e.printStackTrace();
        }catch(IOException e) {
            e.printStackTrace();
        }
    }


}

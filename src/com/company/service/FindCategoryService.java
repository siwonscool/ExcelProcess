package com.company.service;

import com.company.dto.ExcelDto;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class FindCategoryService {

    public List<String> calculateCategoryScore(List<List<String>> inputData, List<List<String>> keywordResults,String[] categories) {
        List<List<Double>> categoryScore = new ArrayList<>();
        double[] scoreBoard = initScoreBoard(categories, 0.1);

        for (int i = 0; i < inputData.size(); i++) {
            categoryScore.add(new ArrayList<>());
            for (int j = 0; j < keywordResults.size(); j++) {
                double count = 0;
                for (int k = 0; k < keywordResults.get(j).size(); k++) {
                    if (inputData.get(i).get(7).contains(keywordResults.get(j).get(k))){
                        count++;
                    }
                }
                categoryScore.get(i).add(count * scoreBoard[j]);
            }
        }

        return peekMaxCategory(categoryScore,categories);
    }

    private static double[] initScoreBoard(String[] categories, double increment){
        double[] scoreBoard = new double[categories.length];
        double startValue = 1.0;

        for (int i = scoreBoard.length -1 ; i >= 0 ; i--) {
            scoreBoard[i] = startValue;
            startValue += increment;
        }
        return scoreBoard;
    }

    private static List<String> peekMaxCategory(List<List<Double>> categoryScore, String[] categories){
        List<String> resultCategory = new ArrayList<>();

        for (int i = 0; i < categoryScore.size(); i++) {
            double max = 0;
            String category = "기타";
            for (int j = 0; j < categoryScore.get(i).size(); j++) {
                if (categoryScore.get(i).get(j) > max){
                    category = categories[j];
                }
            }
            resultCategory.add(category);
        }

        return resultCategory;
    }

    public void updateInputData(ExcelDto excelDto, List<String> resultCategory){
        String fullPath = excelDto.getDataPath()+excelDto.getDataFileName();

        try{
            File file = new File(fullPath);
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file));
            XSSFSheet sheet = workbook.getSheetAt(excelDto.getSheetIndex());

            for (int i = 0; i < resultCategory.size(); i++) {
                sheet.getRow(i+1)
                        .createCell(4)
                        .setCellValue(resultCategory.get(i));
            }

            FileOutputStream outputStream = new FileOutputStream(fullPath);
            workbook.write(outputStream);

        }catch (FileNotFoundException e){
            e.printStackTrace();
        }catch(IOException e) {
            e.printStackTrace();
        }
    }


}

package com.company;

import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.stream.Collectors;

public class HelpDeskKeyword {
    public static void main(String[] args) {
        String path = "C:\\Users\\urp시스템\\Desktop\\helpDesk 분석\\";	//파일 경로 설정
        String filename = "HelpDesk 분석_2019_2022_0310.xlsx";	//파일명 설정
        List<List<String>> list = readExcel(path,filename);
        createDistinctWordSet(list);
        create7ColumnList(list);
        checkSimilarity(create7ColumnList(list),createDistinctWordSet(list));
    }
    public static List<List<String>> readExcel(String path,String filename) {
        List<List<String>> list = new ArrayList<List<String>>();

        try {
            FileInputStream fi = new FileInputStream(path+filename);
            XSSFWorkbook workbook = new XSSFWorkbook(fi);
            XSSFSheet sheet = workbook.getSheetAt(1);

            for(int i=1; i<sheet.getLastRowNum() + 1; i++) {
                XSSFRow row = sheet.getRow(i);
                if(row != null) {
                    List<String> cellList = new ArrayList<String>();
                    for(int j=0; j<row.getLastCellNum(); j++) {
                        XSSFCell cell = row.getCell(j);
                        String value="";
                        //셀이 빈값일경우를 위한 널체크
                        if(cell==null){
                            //continue;
                        }else{
                            //타입별로 내용 읽기
                            switch (cell.getCellType()){
                                case XSSFCell.CELL_TYPE_FORMULA:
                                    value=cell.getCellFormula();
                                    break;
                                case XSSFCell.CELL_TYPE_NUMERIC:
                                    value=cell.getNumericCellValue()+"";
                                    break;
                                case XSSFCell.CELL_TYPE_STRING:
                                    value=cell.getStringCellValue()+"";
                                    break;
                                case XSSFCell.CELL_TYPE_BLANK:
                                    value=cell.getBooleanCellValue()+"";
                                    break;
                                case XSSFCell.CELL_TYPE_ERROR:
                                    value=cell.getErrorCellValue()+"";
                                    break;
                            }
                        }
                        value = value.replaceAll("\\R", " ");

                        cellList.add(value);
                    }
                    list.add(cellList); // 추가된 로우List를 List에 추가
                }
            }
        }catch(FileNotFoundException e) {
            e.printStackTrace();
        }catch(IOException e) {
            e.printStackTrace();
        }

        return list;
    }

    public static Set<String> createDistinctWordSet(List<List<String>> list){

        List<String> wordList = list.stream()
                .map(it -> it.get(7))
                .collect(Collectors.toList());

        //wordList.stream().forEach(System.out::println);
        //System.out.println(wordList.size());

        Set<String> distinctWordSet = wordList.stream()
                .map(it -> it.split(" "))
                .flatMap(Arrays::stream)
                .map(it -> it.replaceAll("[^\uAC00-\uD7A30-9a-zA-Z//>/]"," "))
                .filter(it -> it.length() >1)
                //.map(it -> it.substring(it.length()-2,it.length()))
                .collect(Collectors.toSet());

        //distinctWordSet.stream().forEach(System.out::println);
        return distinctWordSet;
    }

    public static List<String> create7ColumnList(List<List<String>> list){
        List<String> wordList = list.stream()
                .map(it -> it.get(7))
                .collect(Collectors.toList());
        return wordList;
    }

    public static void checkSimilarity(List<String> wordList, Set<String> distinctWordSet){

        Map<String,Integer> wordSimilarityMap = new HashMap<>();

        //wordSimilarityMap =

        for (String str : distinctWordSet) {
            int count = 1;
            for (int i = 0; i < wordList.size(); i++) {
                if (wordList.get(i).contains(str)){
                    if(wordSimilarityMap.get(str)!=null){
                        count++;
                    }
                    wordSimilarityMap.put(str,count);
                }
            }
        }

        wordSimilarityMap.entrySet().stream()

                .forEach(System.out::println);

    }
}

package com.company;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

public class HelpDeskKeyword {
    public static String path = "C:\\Users\\urp시스템\\Desktop\\helpDesk 분석\\";	//파일 경로 설정
    public static String filename = "연도별 분할 데이터.xlsx";	//파일명 설정
    public static String[] categories = {"PC 환경","전자문서","과제관리","메모보고"};

    public static List<List<String>> data = readExcel(path,filename,1);

    public static void main(String[] args) {
        List<List<String>> results = new ArrayList<>();
        List<List<String>> data7Lists = new ArrayList<>();

        // PC 환경 / 전자문서 / 시스템 연계 / 각종 기능 / 과제관리 / 기타 / 메모보고 / 성능 / 웹 기안기
        //for : category
        for (int i = 0; i < categories.length; i++) {
            List<String> data7List = createColumnList(data,7,categories[i]);

            results.add(startKeywordProcess(data7List));
            data7Lists.add(data7List);
        }

        createExcel(data, data7Lists, results, "2021년도분석", path);

    }

    public static List<String> startKeywordProcess(List<String> data7List){

        List<Map.Entry<String, Double>> similarityList = checkSimilarity(data7List, initDistinctWordSet(data7List));

        Map<String,Double> detailOneSimilarityMap = detailSimilarity(data7List,similarityList);
        List<Map.Entry<String,Double>> similarity2List = detailOneSimilarityMap.entrySet().stream()
                .sorted(Collections.reverseOrder(Map.Entry.comparingByValue()))
                .collect(Collectors.toList());

        Map<String,Double> detailTwoSimilarityMap = detailSimilarity(data7List,similarity2List);
        List<Map.Entry<String,Double>> similarity3List = detailTwoSimilarityMap.entrySet().stream()
                .sorted(Collections.reverseOrder(Map.Entry.comparingByValue()))
                .collect(Collectors.toList());

        Map<String,Double> detailThreeSimilarityMap = detailSimilarity(data7List,similarity3List);

        List<String> result = detailTwoSimilarityMap.entrySet().stream()
                .sorted(Collections.reverseOrder(Map.Entry.comparingByValue()))
                .map(it -> it.getKey())
                .collect(Collectors.toList());

        detailTwoSimilarityMap.entrySet().stream()
                .sorted(Collections.reverseOrder(Map.Entry.comparingByValue()))
                .forEach(System.out::println);

        System.out.println("======================================");

        return result;
    }

    public static List<List<String>> readExcel(String path,String filename,int sheetIndex) {
        List<List<String>> list = new ArrayList<List<String>>();

        try {
            FileInputStream fi = new FileInputStream(path+filename);
            XSSFWorkbook workbook = new XSSFWorkbook(fi);
            XSSFSheet sheet = workbook.getSheetAt(sheetIndex);

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

    public static void createExcel(List<List<String>> originData, List<List<String>> column7List, List<List<String>> result, String filename, String path){

        try{
            Workbook workbook = new XSSFWorkbook();
            for (int init = 0; init < categories.length; init++) {
                ArrayList<String> filterList = new ArrayList<>();
                Sheet sheet = workbook.createSheet(categories[init]);
                for (int i = 0; i < result.get(init).size(); i++) {
                    for (int j = 0; j < column7List.get(init).size(); j++) {
                        if (column7List.get(init).get(j).contains(result.get(init).get(i))) {
                            filterList.add(column7List.get(init).get(j));
                        }
                    }
                }

                ArrayList<String> distinctFilterList = (ArrayList<String>) filterList.stream().distinct().collect(Collectors.toList());

                for (int k = 0; k < distinctFilterList.size(); k++) {
                    for (int i = 0; i < originData.size(); i++) {
                        if (originData.get(i).get(7).equals(distinctFilterList.get(k))){
                            Row row = sheet.createRow(k);
                            for (int j = 0; j < originData.get(i).size(); j++) {
                                Cell cell = row.createCell(j);
                                cell.setCellValue(originData.get(i).get(j));
                            }
                        }
                    }
                }

                /*loop1 : for (int k = 0; k < filterList.size(); k++) {
                    for (int i = 0; i < originData.size(); i++) {
                        if (originData.get(i).get(7).equals(filterList.get(k))){
                            Row row = sheet.createRow(k);
                            System.out.println(originData.get(i).get(6) + " : "+originData.get(i).get(7));
                            for (int j = 0; j < originData.get(i).size(); j++) {
                                Cell cell = row.createCell(j);
                                cell.setCellValue(originData.get(i).get(j));
                            }
                            originData.remove(i);
                            continue loop1;
                        }
                    }
                }*/
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

    public static Set<String> initDistinctWordSet(List<String> columnList){


        Set<String> distinctWordSet = columnList.stream()
                .map(it -> it.split(" "))
                .flatMap(Arrays::stream)
                .map(it -> it.replaceAll("[^\uAC00-\uD7A30-9a-zA-Z//>/]"," "))
                .map(it -> it.trim())
                .filter(it -> it.length() > 1)
                .collect(Collectors.toSet());

        return removeZosa(distinctWordSet);
    }

    public static Set<String> detailDistinctWordSet(List<String> columnList, String word){

        Set<String> detailDistinctWordSet = columnList.stream()
                .map(it -> it.split(" "))
                .flatMap(Arrays::stream)
                .map(it -> it.replaceAll("[^\uAC00-\uD7A30-9a-zA-Z//>/]"," "))
                .map(it -> it.trim())
                .filter(it -> it.contains(word))
                .filter(it -> !it.equals(word))
                .collect(Collectors.toSet());

        if (CountWordInExcel(word,columnList) == 1){
            detailDistinctWordSet.add(word);
        }

        return removeZosa(detailDistinctWordSet);
    }

    public static Set<String> removeZosa(Set<String> distinctWordSet){

        String oneZosa ="이/가/을/를/에/는/은";
        String twoZosa ="이나/이고/이라/이랑/에서/가면/에서/에게/한테/으로/않음/안됨/되지/인데";
        String threeZosa ="이라고/이라는/에서는/에서만/에서는/에서도/한테는/한테도/한테만/됩니다/";
        String fourZosa="않습니다/나옵니다/";


        Iterator<String> iterator = distinctWordSet.iterator();

        while (iterator.hasNext()){
            String str = iterator.next();
            if (str.length() <= 1) {
                if (oneZosa.contains(str)) {
                    iterator.remove();
                }
            } else if (str.length() <= 2) {
                if (twoZosa.contains(str)) {
                    iterator.remove();
                }else if (oneZosa.contains(str.substring(str.length() - 1))){
                    iterator.remove();
                }
            } else if (str.length() <= 3){
                if(threeZosa.contains(str.substring(str.length()-3))){
                    iterator.remove();
                }else if (twoZosa.contains(str.substring(str.length()-2))){
                    iterator.remove();
                }else if (oneZosa.contains(str.substring(str.length()-1))){
                    iterator.remove();
                }
            } else {
                if (fourZosa.contains(str.substring(str.length()-4))){
                    iterator.remove();
                }else if(threeZosa.contains(str.substring(str.length()-3))){
                    iterator.remove();
                }else if (twoZosa.contains(str.substring(str.length()-2))){
                    iterator.remove();
                }else if (oneZosa.contains(str.substring(str.length()-1))){
                    iterator.remove();
                }
            }
        }

        return distinctWordSet;
    }

    public static List<String> createColumnList(List<List<String>> list, int column, String word){

        List<String> columnList = list.stream()
                .filter(it -> it.get(4).contains(word))
                .map(it -> it.get(column))
                .collect(Collectors.toList());
        System.out.println(columnList.size());
        return columnList;
    }

    public static List<Map.Entry<String, Double>> checkSimilarity(List<String> columnList, Set<String> distinctWordSet){

        Map<String,Double> wordSimilarityMap = new HashMap<>();

        for (String str : distinctWordSet) {
            double count = 1;

            for (int i = 0; i < columnList.size(); i++) {
                if (columnList.get(i).contains(str)){

                    if(wordSimilarityMap.get(str)!=null){
                        count++;
                    }

                    //double result = Math.round(count/ data.size() * 1000)/1000.0;
                    double result = count / columnList.size();

                    wordSimilarityMap.put(str, result);
                }
            }
        }

        List<Map.Entry<String,Double>> similarityWordList = wordSimilarityMap.entrySet().stream()
                //.sorted(Comparator.comparing(Map.Entry::getKey)) // 단어기준 정렬
                .sorted(Collections.reverseOrder(Map.Entry.comparingByValue())) // 유사도기준 정렬
                .collect(Collectors.toList());

        return similarityWordList;
    }

    public static Map<String,Double> detailSimilarity(List<String> columnList, List<Map.Entry<String,Double>> beforeSimilarityWordList){
        Map<String, Double> detailMap = new HashMap<>();

        for (int i = 0; i < beforeSimilarityWordList.size(); i++) {
            Set<String> set = detailDistinctWordSet(columnList,beforeSimilarityWordList.get(i).getKey());
            Double wordCount = CountWordInExcel(beforeSimilarityWordList.get(i).getKey(),columnList);

            List<Map.Entry<String,Double>> list = checkDetailSimilarity(columnList,set,wordCount);

            for (int j = 0; j < list.size(); j++) {
                if (detailMap.get(list.get(j).getKey()) != null){
                    double beforeValue = detailMap.get(list.get(j).getKey());
                    double afterValue = list.get(j).getValue();

                    double finalValue = beforeValue + afterValue;

                    detailMap.put(list.get(j).getKey(),finalValue);
                }else {
                    double value = list.get(j).getValue() * beforeSimilarityWordList.get(i).getValue();
                    detailMap.put(list.get(j).getKey(),value);
                }
            }
        }

        return detailMap;
    }

    private static Double CountWordInExcel(String key, List<String> columnList) {
        double count = 0;
        for (int i = 0; i < columnList.size(); i++) {
            if (columnList.get(i).contains(key));
            count ++;
        }

        return count;
    }

    private static List<Map.Entry<String, Double>> checkDetailSimilarity(List<String> columnList, Set<String> detailDistinctWordSet, Double wordCount) {
        Map<String,Double> detailSimilarityMap = new HashMap<>();

        for (String str : detailDistinctWordSet) {
            double count = 1;

            for (int i = 0; i < columnList.size(); i++) {
                if (columnList.get(i).contains(str)){

                    if(detailSimilarityMap.get(str)!=null){
                        count++;
                    }

                    //double result = Math.round(count / wordCount * 1000)/1000.0;
                    double result = count / wordCount;

                    detailSimilarityMap.put(str, result);
                }
            }
        }

        return detailSimilarityMap.entrySet().stream()
                .sorted(Collections.reverseOrder(Map.Entry.comparingByValue()))
                .collect(Collectors.toList());
    }


    //2차 수정 logic 인데 걍 버릴거임 ㅋㅋ
    @Deprecated
    public static List<String> makeDictionary(Set<String> distinctWordSet){
        List<String> includedList = distinctWordSet.stream()
                .collect(Collectors.toList());
        for (String strSet : distinctWordSet){
            for (int i =0 ;i<includedList.size();i++){
                if (!strSet.equals(includedList.get(i)) && includedList.get(i).contains(strSet)){
                    includedList.remove(i);
                }
            }
        }
        return includedList;
    }

}
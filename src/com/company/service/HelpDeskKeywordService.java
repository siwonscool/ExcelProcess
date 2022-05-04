package com.company.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

public class HelpDeskKeywordService {

    private static final String oneZosa = "이/가/을/를/에/는/은";
    private static final String twoZosa = "이나/이고/이라/이랑/에서/가면/에서/에게/한테/으로/않음/안됨/되지/인데";
    private static final String threeZosa = "이라고/이라는/에서는/에서만/에서는/에서도/한테는/한테도/한테만/됩니다/";
    private static final String fourZosa = "않습니다/나옵니다/";

    public List<String> startKeywordProcess(List<String> data7List, int numberProcess) {

        List<Map.Entry<String, Double>> similarityList = checkSimilarity(data7List,
            initDistinctWordSet(data7List));
        Map<String, Double> detailSimilarityMap = new HashMap<>();

        for (int i = 0; i < numberProcess; i++) {
            detailSimilarityMap = detailSimilarity(data7List, similarityList);
            similarityList = detailSimilarityMap.entrySet().stream()
                .sorted(Collections.reverseOrder(Map.Entry.comparingByValue()))
                .collect(Collectors.collectingAndThen(Collectors.toList(),
                    Collections::unmodifiableList));
        }

        return detailSimilarityMap.entrySet().stream()
            .sorted(Collections.reverseOrder(Map.Entry.comparingByValue()))
            .map(Map.Entry::getKey)
            .collect(
                Collectors.collectingAndThen(Collectors.toList(), Collections::unmodifiableList));

    }

    private Set<String> initDistinctWordSet(List<String> columnList) {

        return removeZosa(columnList.stream()
            .map(it -> it.split(" "))
            .flatMap(Arrays::stream)
            .map(it -> it.replaceAll("[^\uAC00-\uD7A30-9a-zA-Z//>/]", " "))
            .map(String::trim)
            .filter(it -> it.length() > 1)
            .collect(Collectors.toSet()));
    }

    private Set<String> detailDistinctWordSet(List<String> columnList, String word) {

        return removeZosa(columnList.stream()
            .map(it -> it.split(" "))
            .flatMap(Arrays::stream)
            .map(it -> it.replaceAll("[^\uAC00-\uD7A30-9a-zA-Z//>/]", " "))
            .map(String::trim)
            .filter(it -> it.contains(word))
            .filter(it -> !it.equals(word))
            .collect(Collectors.toSet()));
    }

    private Set<String> removeZosa(Set<String> distinctWordSet) {

        Iterator<String> iterator = distinctWordSet.iterator();

        //while 문 지양하기
        while (iterator.hasNext()) {
            String str = iterator.next();
            if (str.length() <= 1) {
                if (oneZosa.contains(str)) {
                    iterator.remove();
                }
            } else if (str.length() <= 2) {
                if (twoZosa.contains(str)) {
                    iterator.remove();
                } else if (oneZosa.contains(str.substring(str.length() - 1))) {
                    iterator.remove();
                }
            } else if (str.length() <= 3) {
                if (threeZosa.contains(str.substring(str.length() - 3))) {
                    iterator.remove();
                } else if (twoZosa.contains(str.substring(str.length() - 2))) {
                    iterator.remove();
                } else if (oneZosa.contains(str.substring(str.length() - 1))) {
                    iterator.remove();
                }
            } else {
                if (fourZosa.contains(str.substring(str.length() - 4))) {
                    iterator.remove();
                } else if (threeZosa.contains(str.substring(str.length() - 3))) {
                    iterator.remove();
                } else if (twoZosa.contains(str.substring(str.length() - 2))) {
                    iterator.remove();
                } else if (oneZosa.contains(str.substring(str.length() - 1))) {
                    iterator.remove();
                }
            }
        }

        return distinctWordSet;
    }

    public List<String> createColumnList(List<List<String>> list, int column, String category) {

        return list.stream()
            .filter(it -> it.get(4).contains(category))
            .map(it -> it.get(column))
            .collect(
                Collectors.collectingAndThen(Collectors.toList(), Collections::unmodifiableList));
    }

    private List<Map.Entry<String, Double>> checkSimilarity(List<String> columnList,
        Set<String> distinctWordSet) {
        Map<String, Double> wordSimilarityMap = new HashMap<>();

        for (String str : distinctWordSet) {
            double count = 1;
            for (int i = 0; i < columnList.size(); i++) {
                if (columnList.get(i).contains(str)) {
                    if (wordSimilarityMap.get(str) != null) {
                        count++;
                    }
                }
            }
            wordSimilarityMap.put(str, count / columnList.size());
        }

        return wordSimilarityMap.entrySet().stream()
            .sorted(Collections.reverseOrder(Map.Entry.comparingByValue())) // 유사도기준 정렬
            .collect(
                Collectors.collectingAndThen(Collectors.toList(), Collections::unmodifiableList));
    }

    private List<Map.Entry<String, Double>> checkDetailSimilarity(List<String> columnList,
        Set<String> detailDistinctWordSet, Double wordCount) {
        Map<String, Double> detailSimilarityMap = new HashMap<>();

        for (String str : detailDistinctWordSet) {
            double count = 1;
            for (int i = 0; i < columnList.size(); i++) {
                if (columnList.get(i).contains(str)) {
                    if (detailSimilarityMap.get(str) != null) {
                        count++;
                    }
                    detailSimilarityMap.put(str, count / wordCount);
                }
            }
        }

        return detailSimilarityMap.entrySet().stream()
            .sorted(Collections.reverseOrder(Map.Entry.comparingByValue()))
            .collect(
                Collectors.collectingAndThen(Collectors.toList(), Collections::unmodifiableList));
    }

    private Map<String, Double> detailSimilarity(List<String> columnList,
        List<Map.Entry<String, Double>> beforeSimilarityWordList) {
        Map<String, Double> detailMap = new HashMap<>();

        for (int i = 0; i < beforeSimilarityWordList.size(); i++) {
            Set<String> set = detailDistinctWordSet(columnList,
                beforeSimilarityWordList.get(i).getKey());
            Double wordCount = CountWordInExcel(beforeSimilarityWordList.get(i).getKey(),
                columnList);

            List<Map.Entry<String, Double>> list = checkDetailSimilarity(columnList, set,
                wordCount);

            for (int j = 0; j < list.size(); j++) {
                if (detailMap.get(list.get(j).getKey()) != null) {
                    double beforeValue = detailMap.get(list.get(j).getKey());
                    double afterValue = list.get(j).getValue();
                    double finalValue = beforeValue + afterValue;

                    detailMap.put(list.get(j).getKey(), finalValue);
                } else {
                    double value =
                        list.get(j).getValue() * beforeSimilarityWordList.get(i).getValue();
                    detailMap.put(list.get(j).getKey(), value);
                }
            }
        }

        return detailMap;
    }

    private Double CountWordInExcel(String key, List<String> columnList) {
        double count = 0;
        for (int i = 0; i < columnList.size(); i++) {
            if (columnList.get(i).contains(key))
                ;
            count++;
        }

        return count;
    }

}
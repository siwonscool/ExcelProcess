package com.company;

import com.company.dto.ExcelDto;
import com.company.service.CompareDataService;
import com.company.service.FindCategoryService;
import com.company.service.HelpDeskKeywordService;

import java.util.ArrayList;
import java.util.List;

public class Main {
    static String path = "C:\\Users\\urp시스템\\Desktop\\helpDesk 분석\\";   //파일 경로 설정
    static String originFilename = "HelpDesk 분석_2019_2022_0310.xlsx";	//파일명 설정
    static String inputFilename= "문의유형없는_데이터.xlsx";

    // 전자문서 / 메모보고 / 과제관리 / PC 환경 / 시스템 연계 / 웹 기안기 / 기타 /
    static String[] categories = {"전자문서","PC 환경","웹 기안기","시스템 연계","메모보고","과제관리"};
    static ExcelDto originDataDto = new ExcelDto(path, originFilename,1);
    static ExcelDto inputDataDto = new ExcelDto(path, inputFilename,0);

    static HelpDeskKeywordService helpDeskKeywordService = new HelpDeskKeywordService();
    static FindCategoryService findCategoryService = new FindCategoryService();
    static CompareDataService compareDataService = new CompareDataService();

    public static void main(String[] args) {

        List<List<String>> keywordResults = new ArrayList<>();
        List<List<String>> data7Lists = new ArrayList<>();

        for (int i = 0; i < categories.length; i++) {
            List<String> data7List = helpDeskKeywordService.createColumnList(originDataDto.readExcel(),7, categories[i]);
            System.out.println("'" + categories[i] + "'" + "(이)가 포함된 row 의 개수 : " + data7List.size());
            System.out.println("'" + categories[i] + "'" + " 카테고리 에서 키워드 추출중...");

            keywordResults.add(helpDeskKeywordService.startKeywordProcess(data7List,1));
            data7Lists.add(data7List);
        }

        //helpDeskKeywordService.createExcel(data, data7Lists, keywordResults, "2021년도분석", path, categories);

        List<String> resultCategory = findCategoryService.calculateCategoryScore(inputDataDto.readExcel(), keywordResults, categories);
        findCategoryService.updateInputData(inputDataDto,resultCategory);
        compareDataService.createCompareResultExcel(originDataDto.readExcel(),inputDataDto.readExcel(),path,"오차 분석");
    }
}

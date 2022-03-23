package com.company;

import java.io.FileInputStream;
import java.util.*;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExelTest {
    public static void main(String[] args) {
        String path = "C:\\Users\\urp시스템\\Desktop\\발표장표\\";	//파일 경로 설정
        String filename = "납품 품목_v1.3.xlsx";	//파일명 설정
        List<List<String>> list=readExcel(path,filename);
        countAgency(list);
    }
    public static List<List<String>> readExcel(String path,String filename) {
        List<List<String>> list = new ArrayList<List<String>>();

        try {
            FileInputStream fi = new FileInputStream(path+filename);
            XSSFWorkbook workbook = new XSSFWorkbook(fi);
            XSSFSheet sheet = workbook.getSheetAt(1);

            for(int i=0; i<sheet.getLastRowNum(); i++) {
                XSSFRow row = sheet.getRow(i);
                if(row != null) {
                    List<String> cellList = new ArrayList<String>();
                    for(int j=0; j<row.getLastCellNum(); j++) {
                        XSSFCell cell = row.getCell(j);
                        String value="";
                        //셀이 빈값일경우를 위한 널체크
                        if(cell==null){
                            continue;
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

    public static void countAgency(List<List<String>> list){
        Map<String,String> map = new HashMap<>();
        for (int i = 0; i < list.size(); i++) {
            for (int j = 0; j < list.get(i).size(); j++) {
                if (list.get(i).get(j).equals("서버 가상화")){
                    //if (Server32Core(list,i) || Server40Core(list,i)){
                        if (map.get(list.get(i).get(1))!=null){
                            double beforeValue = Double.parseDouble(map.get(list.get(i).get(1)));
                            double afterValue = Double.parseDouble(list.get(i).get(5));
                            int intValue = (int) (afterValue+beforeValue);
                            String finalValue = intValue+"";
                            map.put(list.get(i).get(1),finalValue);
                        }else {
                            double beforeValue = Double.parseDouble(list.get(i).get(5));
                            int intValue = (int) beforeValue;
                            map.put(list.get(i).get(1), intValue+"");
                        }
                    //}
                }
            }
        }

        Iterator<Map.Entry<String,String>> entries = map.entrySet().iterator();
        Double testValue = 0.0;
        while (entries.hasNext()){
            Map.Entry<String,String> entry = entries.next();
            testValue += Double.parseDouble(entry.getValue());
            System.out.print(entry.getKey()+"청 "+entry.getValue()+"식, ");
        }
        System.out.println();
        System.out.println("기관수 : "+map.size());
        System.out.println("총 기기 수 : "+testValue);
    }

    public static boolean Server80Core(List<List<String>> list,int rowNum){
        return list.get(rowNum).get(4).equals("Intel Xeon 2세대 2.5GHz*4, 80Core, 512GB, HDD(SSD) 480GB*2EA, HBA*2Ports, 10GbE*4Ports");
    }

    public static boolean Server40Core(List<List<String>> list,int rowNum){
        return list.get(rowNum).get(4).equals("Intel Xeon 2세대 3.1GHz*2, 40Core, 256GB, HDD(SSD) 480GB*2EA, HBA*2Ports, 10GbE*2Ports")||
                list.get(rowNum).get(4).equals("Intel Xeon 2세대 3.1GHz*2, 40Core, 512GB, HDD(SSD) 480GB*2EA, HBA*2Ports, 10GbE*4Ports")||
                list.get(rowNum).get(4).equals("Intel Xeon 2세대 3.1GHz*2, 40Core, 256GB, HDD(SSD) 480GB*2EA, 10GbE*2Ports");
    }

    public static boolean Server32Core(List<List<String>> list,int rowNum){
        return list.get(rowNum).get(4).equals("Intel Xeon 2세대 2.9GHz*2, 32Core, 128GB, HDD(SSD) 480GB*2EA, HBA*2Ports, 10GbE*2Ports")||
                list.get(rowNum).get(4).equals("Intel Xeon 2세대 2.9GHz*2, 32Core, 256GB, HDD(SSD) 480GB*2EA, 10GbE*2Ports")||
                list.get(rowNum).get(4).equals("Intel Xeon 2세대 2.9GHz*2, 32Core, 256GB, HDD(SSD) 480GB*2EA, HBA*2Ports, 10GbE*2Ports")||
                list.get(rowNum).get(4).equals("Intel Xeon 2세대 2.9GHz*2, 32Core, 256GB, HDD(SSD) 480GB*2EA, HBA*4Ports, 10GbE*2Ports");
    }

}


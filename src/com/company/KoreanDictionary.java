package com.company;

import org.json.JSONObject;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;

public class KoreanDictionary {
    public static void main(String[] args) {
        getJson();
    }

    public static void getJson() {
        try {
            URL url = new URL("https://opendict.korean.go.kr/api/view");
            HttpURLConnection conn = (HttpURLConnection)url.openConnection();

            conn.setRequestMethod("GET"); // http 메서드
            conn.setRequestProperty("Content-Type", "application/json"); // header Content-Type 정보
            conn.setRequestProperty("key", "D67054846C182842A9C8E4A256FA6D1D"); // header의 auth 정보
            conn.setRequestProperty("q","UTF-8");
            conn.setRequestProperty("pos","8");
            conn.setDoOutput(true); // 서버로부터 받는 값이 있다면 true

            // 서버로부터 데이터 읽어오기
            BufferedReader br = new BufferedReader(new InputStreamReader(conn.getInputStream()));
            StringBuilder sb = new StringBuilder();
            String line = null;

            while((line = br.readLine()) != null) { // 읽을 수 있을 때 까지 반복
                sb.append(line);
            }

            JSONObject obj = new JSONObject(sb.toString()); // json으로 변경 (역직렬화)
            System.out.println("code= " + obj.getInt("code") + " / message= " + obj.getString("message"));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

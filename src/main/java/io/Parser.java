package io;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

public class Parser {
    public static void main(String[] args) throws Exception {
        parse("Java", "C:\\parsed.xlsx");
    }

    public static void parse(String query, String outputPath) throws Exception {

        // создание таблицы
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet spreadsheet = workbook.createSheet(" Вакансии ");

        // записываем название столбцов в первую строку
        int cellid = 0;
        XSSFRow row = spreadsheet.createRow(0);
        for (String item : new String[]{"НАЗВАНИЕ ВАКАНСИИ", "ЗАРПЛАТА ОТ", "ЗАРПЛАТА ДО", "ОПЫТ", "ТРЕБОВАНИЯ"}) {
            Cell cell = row.createCell(cellid++);
            cell.setCellValue(item);
        }

        // вычисление количества страниц
        String url = "https://api.hh.ru/vacancies?area=1&text=" + query;
        int pages = pages(url);

        // постраничная запись в таблицу
        for (int p = 0; p < pages; p++) {
            url = "https://api.hh.ru/vacancies?area=1&page=" + p + "&per_page=10&text=" + query;
            getPageAndWrite(url, spreadsheet, p*10+1);
        }

        // запись таблицы в файл
        FileOutputStream out;
        try {
            out = new FileOutputStream(outputPath);
        } catch (FileNotFoundException e) {
            throw new RuntimeException("File " + outputPath +  " Not Found");
        }
        workbook.write(out);
        out.close();
    }

    /**
     * Processes one page
     * */
    private static void getPageAndWrite(String link, XSSFSheet spreadsheet, int rowid) throws Exception {
        JSONArray vacancies = new JSONObject(callApi(link)).getJSONArray("items");

        String[] parsedVacancy;
        XSSFRow row;
        int cellid;
        // извлекаем нужную информацию из каждой вакансии
        for(int i = 0; i < vacancies.length(); i++) {
            JSONObject vacancy = vacancies.getJSONObject(i);
            parsedVacancy = new String[5];

            // достаём название вакансии
            if (hasNonNullKey(vacancy, "name"))
                parsedVacancy[0] = vacancy.get("name").toString();

            // достаём информацию о зарплате
            if (hasNonNullKey(vacancy, "salary")) {
                JSONObject salary = vacancy.getJSONObject("salary");
                parsedVacancy[1] = hasNonNullKey(salary, "from") ? salary.get("from").toString() : "";
                parsedVacancy[2] = hasNonNullKey(salary, "to") ? salary.get("to").toString() : "";
            }

            // достаём информацию об опыте
            parsedVacancy[3] = hasNonNullKey(vacancy, "experience") ?
                    vacancy.getJSONObject("experience").get("name").toString() : "";

            // достаём информацию о требованиях
            if (hasNonNullKey(vacancy, "snippet")) {
                JSONObject snippet = vacancy.getJSONObject("snippet");
                if (hasNonNullKey(snippet, "requirement"))
                    parsedVacancy[4] = snippet.get("requirement").toString();
            }

            // запись информации о вакансии в таблицу
            row = spreadsheet.createRow(rowid++);
            cellid = 0;
            for (String item : parsedVacancy) {
                Cell cell = row.createCell(cellid++);
                cell.setCellValue(item);
            }
        }

    }

    /**
     * Makes HTTP request and returns json response
     * */
    private static String callApi(String link) throws Exception {
        // создаём и отправляем GET запрос
        URL url = new URL(link);
        HttpURLConnection con = (HttpURLConnection) url.openConnection();
        con.setRequestMethod("GET");
        con.setConnectTimeout(2500);
        con.setReadTimeout(2500);
        con.connect();

        // если код запроса не успешный выбрасываем исключение
        if (HttpURLConnection.HTTP_OK != con.getResponseCode())
            throw new RuntimeException("HTTP error: " + con.getResponseCode() + ". Try running again.");

        // считываем JSON строку
        BufferedReader bf = new BufferedReader(new InputStreamReader(con.getInputStream()));
        String line;
        StringBuilder sb = new StringBuilder();
        while((line = bf.readLine()) != null) {
            sb.append(line);
            sb.append("\n");
        }
        return sb.toString();
    }

    /**
     * Returns number of pages
     * */
    private static int pages(String link) throws Exception{
        return new JSONObject(callApi(link)).getInt("pages");
    }
    public static boolean hasNonNullKey(JSONObject obj, String key) {
        return obj.has(key) && !obj.isNull(key);
    }
}

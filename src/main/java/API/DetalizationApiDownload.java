package API;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.databind.node.*;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import enumsMaps.Translator;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.*;

public class DetalizationApiDownload {
    public static void downloadDetalization(String directoryReports, String dateFrom, String dateTo, String token) throws IOException {
        String apiUrl = "https://statistics-api.wildberries.ru/api/v5/supplier/reportDetailByPeriod?dateFrom=" + dateFrom + "&dateTo=" + dateTo + "&rrdid=0&limit=100000";

        try {
            // Выполняем HTTP GET запрос
            URL url = new URL(apiUrl);
            HttpURLConnection connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");
            connection.setRequestProperty("Authorization", "Bearer " + token);
            connection.setConnectTimeout(5000);
            connection.setReadTimeout(5000);

            int responseCode = connection.getResponseCode();
            if (responseCode == HttpURLConnection.HTTP_OK) {
                InputStream inputStream = connection.getInputStream();
                String jsonResponse = new Scanner(inputStream, "UTF-8").useDelimiter("\\A").next();
                inputStream.close();

                //сохраняем json файл
                String jsonPath = "e:\\Proga\\AnaliticsWBMain\\detalization.json";
                Gson gson = new GsonBuilder().setPrettyPrinting().create();
                String prettyJson = gson.toJson(gson.fromJson(jsonResponse, Object.class));
                try (BufferedWriter writer = new BufferedWriter(new FileWriter(jsonPath))){
                    writer.write(prettyJson);
                }

                ObjectMapper objectMapper = new ObjectMapper();
                objectMapper.enable(SerializationFeature.ORDER_MAP_ENTRIES_BY_KEYS); // Включаем порядок ключей
                JsonNode jsonNode = objectMapper.readTree(jsonResponse);
                ArrayNode jsonArray = (ArrayNode) jsonNode;

                // Создаем новый Excel-файл
                Workbook workbook = new XSSFWorkbook();
                Sheet sheet = workbook.createSheet("Report");

                //создаем ряд для заголовков столбцов
                Row headRow = sheet.createRow(0);

                //записываем заголовки столбцов и все данные
                ArrayList<String> headlines = new ArrayList<>();
                int j = 1;
                int k = 0;
                for (JsonNode node: jsonArray) {
                    Iterator<Map.Entry<String, JsonNode>> nodeFields = node.fields();
                    Row nextRow = sheet.createRow(j);
                    while (nodeFields.hasNext()) {
                        Map.Entry<String, JsonNode> entry = nodeFields.next();
                        String key = entry.getKey();
                        int cellNumber = headlines.indexOf(key);

                        //если такого заголовка не существует, то добавляем его
                        if(cellNumber == -1) {
                            Cell nextHeadlineCell = headRow.createCell(k);
                            headlines.add(key);
                            cellNumber = headlines.indexOf(key);
                            if(Translator.translateMap.containsKey(key)) {
                                key = Translator.translateMap.get(key);
                            }
                            nextHeadlineCell.setCellValue(key);
                            k++;
                        }

                        Cell nextCell = nextRow.createCell(cellNumber);
                        JsonNode value = entry.getValue();
                        if (value instanceof IntNode) {
                            nextCell.setCellValue(value.asInt());
                        }
                        if (value instanceof DoubleNode) {
                            nextCell.setCellValue(value.asDouble());
                        }
                        if (value instanceof TextNode) {
                            nextCell.setCellValue(value.asText());
                        }
                        if (value instanceof LongNode) {
                            nextCell.setCellValue(value.asLong());
                        }
                        if (value instanceof NullNode) {
                            nextCell.setCellValue("");
                        }
                    }
                    j++;
                }

                // Записываем в файл Excel
                try (FileOutputStream outputStream = new FileOutputStream(directoryReports + "Детализация.xlsx")) {
                    workbook.write(outputStream);
                }
                workbook.close();
            } else {
                System.out.println("Ошибка HTTP детализация: " + responseCode);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
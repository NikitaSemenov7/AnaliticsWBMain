package API;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.databind.node.*;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import enumsMaps.Translator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;
import java.util.Scanner;

public class StorageApiDownload {
    public static void downloadStorage(String directoryReports, String dateFrom, String dateTo, String token) throws IOException, InterruptedException {

        //создаем отчет
        String apiUrl = "https://seller-analytics-api.wildberries.ru/api/v1/paid_storage?dateFrom=" + dateFrom + "&dateTo=" + dateTo;
        String taskId = "";

        try {
            URL url = new URL(apiUrl);
            HttpURLConnection connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");
            connection.setRequestProperty("Authorization", "Bearer " + token);
            connection.setConnectTimeout(5000);
            connection.setReadTimeout(5000);

            int responseCode = connection.getResponseCode();
            if (responseCode == HttpURLConnection.HTTP_OK) {
                System.out.println("Отправлено задание на создание отчета по хранению. Статус: " + responseCode);

                //получаем taskId сгенерированного отчета
                InputStream inputStream = connection.getInputStream();
                String jsonResponse = new Scanner(inputStream, "UTF-8").useDelimiter("\\A").next();
                inputStream.close();

                //сохраняем json файл
                //раскомментировать, если нужно будет посмотреть json
                String jsonPath = "e:\\Proga\\AnaliticsWBMain\\storageTaskId.json";
                Gson gson = new GsonBuilder().setPrettyPrinting().create();
                String prettyJson = gson.toJson(gson.fromJson(jsonResponse, Object.class));
                try (BufferedWriter writer = new BufferedWriter(new FileWriter(jsonPath))) {
                    writer.write(prettyJson);
                }

                //получаем taskId
                //в этой строке поменял prettyJson на jsonResponse
                JSONObject jsonObject = new JSONObject(jsonResponse);
                JSONObject dataObject = jsonObject.getJSONObject("data");
                taskId = dataObject.getString("taskId");
            } else {
                System.out.println("Отчет по хранению не создался. Статус: " + responseCode);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

        //проверяем статус задания на генерацию
        apiUrl = "https://seller-analytics-api.wildberries.ru/api/v1/paid_storage/tasks/{" + taskId + "}/status";

        String status = "";
        int responseCode = 0;
        do {
            try {
                URL url = new URL(apiUrl);
                HttpURLConnection connection = (HttpURLConnection) url.openConnection();
                connection.setRequestMethod("GET");
                connection.setRequestProperty("Authorization", "Bearer " + token);
                connection.setConnectTimeout(5000);
                connection.setReadTimeout(5000);

                responseCode = connection.getResponseCode();
                if (responseCode == HttpURLConnection.HTTP_OK) {
                    Thread.sleep(2000);
                    InputStream inputStream = connection.getInputStream();
                    String jsonResponse = new Scanner(inputStream, "UTF-8").useDelimiter("\\A").next();
                    inputStream.close();

                    //сохраняем json файл
                    String jsonPath = "e:\\Proga\\AnaliticsWBMain\\storageTaskStatus.json";
                    Gson gson = new GsonBuilder().setPrettyPrinting().create();
                    String prettyJson = gson.toJson(gson.fromJson(jsonResponse, Object.class));
                    try (BufferedWriter writer = new BufferedWriter(new FileWriter(jsonPath))) {
                        writer.write(prettyJson);
                    }

                    //получаем taskId
                    JSONObject jsonObject = new JSONObject(prettyJson);
                    JSONObject dataObject = jsonObject.getJSONObject("data");
                    status = dataObject.getString("status");
                    System.out.println("Статус создания отчета по хранению: " + status);
                } else {
                    System.out.println("Со статусом задания на генерацию отчета что-то не так. Статус: " + responseCode);
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        } while (!status.equals("done") && responseCode != 429);



        //получаем отчет
        apiUrl = "https://seller-analytics-api.wildberries.ru/api/v1/paid_storage/tasks/{" + taskId + "}/download";
        try {
            URL url = new URL(apiUrl);
            HttpURLConnection connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");
            connection.setRequestProperty("Authorization", "Bearer " + token);
            connection.setConnectTimeout(5000);
            connection.setReadTimeout(5000);

            responseCode = connection.getResponseCode();
            if (responseCode == HttpURLConnection.HTTP_OK) {
                InputStream inputStream = connection.getInputStream();
                String jsonResponse = new Scanner(inputStream, "UTF-8").useDelimiter("\\A").next();
                inputStream.close();

                //сохраняем json файл
                String jsonPath = "e:\\Proga\\AnaliticsWBMain\\storage.json";
                Gson gson = new GsonBuilder().setPrettyPrinting().create();
                String prettyJson = gson.toJson(gson.fromJson(jsonResponse, Object.class));
                try (BufferedWriter writer = new BufferedWriter(new FileWriter(jsonPath))) {
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
                for (JsonNode node : jsonArray) {
                    Iterator<Map.Entry<String, JsonNode>> nodeFields = node.fields();
                    Row nextRow = sheet.createRow(j);
                    while (nodeFields.hasNext()) {
                        Map.Entry<String, JsonNode> entry = nodeFields.next();
                        String key = entry.getKey();
                        int cellNumber = headlines.indexOf(key);

                        //если такого заголовка не существует, то добавляем его
                        if (cellNumber == -1) {
                            Cell nextHeadlineCell = headRow.createCell(k);
                            headlines.add(key);
                            cellNumber = headlines.indexOf(key);
                            if (Translator.translateMap.containsKey(key)) {
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
                try (FileOutputStream outputStream = new FileOutputStream(directoryReports + "Хранение.xlsx")) {
                    workbook.write(outputStream);
                }
                workbook.close();

                System.out.println("Отчет по хранению сохранен. Статус: " + responseCode);
            } else {
                System.out.println("Отчет по хранению не получили. Статус: " + responseCode);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

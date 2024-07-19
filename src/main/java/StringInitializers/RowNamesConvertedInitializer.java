package StringInitializers;

import enumsMaps.Alphabet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;
import java.util.Collections;

public class RowNamesConvertedInitializer {
    private static ArrayList<String> categories = new ArrayList();
    private static ArrayList<String> articlesInCategory = new ArrayList();
    private static ArrayList<String> articlesInSebestoimost = new ArrayList<>();
    private static ArrayList<String> articlesNotInSebestoimost = new ArrayList<>();
    private static ArrayList<String> articlesInDetalization = new ArrayList<>();
    private static ArrayList<String> rowNamesConverted = new ArrayList<>();

    public static ArrayList<String> initRowNamesConverted(ArrayList<String> columnNamesDetalization, ArrayList<String> columnNamesStorage, XSSFSheet sheetReportFromWB, XSSFSheet sheetReportConverted, XSSFSheet sheetHranenie, XSSFSheet sheetSebestoimost) {
        //инициализируем Общий и Неизвестно
        initGeneralAndUnknown(sheetReportConverted);

        //считываем все артикулы из ВводныеСебестоимость
        if (articlesInSebestoimost.isEmpty()) {
            for (Row row : sheetSebestoimost) {
                String article = row.getCell(Alphabet.B.ordinal()).getStringCellValue();
                articlesInSebestoimost.add(article);
            }
        }


        for (Row row : sheetReportFromWB) {
            //проходим по всем категориям и добавляем в список уникальные категории
            Cell cellCategory = row.getCell(columnNamesDetalization.indexOf("Предмет"));
            if (cellCategory != null) {
                String category = cellCategory.getStringCellValue();
                if (!categories.contains(category) && !category.equals("Предмет") && !category.isEmpty()) {
                    categories.add(category);
                }
            }

            //проходим по артикулам и вносим их в список, чтобы понимать, что их категория уже внесена, чтобы при проходе по
            //таблице с хранением, мы не внесли эту же категорию с другим названием
            Cell cellArticle = row.getCell(columnNamesDetalization.indexOf("Артикул поставщика"));
            if (cellArticle != null) {
                String article = cellArticle.getStringCellValue();
                if(!articlesInDetalization.contains(article)) {
                    articlesInDetalization.add(article);
                }
            }
        }

        //проходим по всем категориям в хранении и добавляем в список уникальные категории, которые не добавили раньше
        //одна и та же категория в Детализации и в Хранении может называться по-разному
        for (Row row : sheetHranenie) {
            Cell cellCategory = row.getCell(columnNamesStorage.indexOf("Предмет"));
            Cell cellArticle = row.getCell(columnNamesStorage.indexOf("Артикул продавца"));
            if (cellCategory != null && cellArticle != null) {
                String category = cellCategory.getStringCellValue();
                String article = cellArticle.getStringCellValue().toLowerCase();
                if (!categories.contains(category) && !category.equals("Предмет")
                        && !category.isEmpty() && !articlesInDetalization.contains(article)) {
                    categories.add(category);
                }
            }
        }

        //сортируем категории по алфавиту
        Collections.sort(categories);

        for (String category : categories) {
            //вставляем каждую категорию в таблицу
            sheetReportConverted.createRow(sheetReportConverted.getLastRowNum() + 1).createCell(0).setCellValue(category);
            rowNamesConverted.add(category);

            //собираем артикулы по добавленной категории и сортируем их по алфавиту
            initArticlesReportFromWB(columnNamesDetalization, sheetReportFromWB, sheetReportConverted, category);
            initArticlesHranenie(category, columnNamesStorage, sheetHranenie, sheetReportConverted);
            Collections.sort(articlesInCategory);

            //добавляем артикулы вставленной категории в таблицу
            for (String article : articlesInCategory) {
                sheetReportConverted.createRow(sheetReportConverted.getLastRowNum() + 1).createCell(0).setCellValue(article);
                rowNamesConverted.add(article);
            }
        }

        return rowNamesConverted;
    }


    private static void initArticlesReportFromWB(ArrayList<String> columnNamesDetalization, XSSFSheet sheetReportFromWB, XSSFSheet sheetReportConverted, String category) {
        articlesInCategory.clear();
        for (Row row : sheetReportFromWB) {
            Cell cell = row.getCell(columnNamesDetalization.indexOf("Предмет"));
            if (cell != null) {
                String categoryToCheck = cell.getStringCellValue();
                if (categoryToCheck.equals(category)) {
                    Cell cellArticle = row.getCell(columnNamesDetalization.indexOf("Артикул поставщика"));
                    String article = cellArticle.getStringCellValue().toLowerCase();
                    //делаем проверку, что артикул есть в ВводныеСебестоимость
                    if (!articlesInSebestoimost.contains(article)) {
                        if (!articlesNotInSebestoimost.contains(article)) {
                            System.out.println("Артикул " + article + " отсутствует в ВводныеСебестоимость");
                            articlesNotInSebestoimost.add(article);
                        }
                    }

                    if (!articlesInCategory.contains(article)) {
                        articlesInCategory.add(article);
                    }
                }
            }
        }
    }

    private static void initArticlesHranenie(String category, ArrayList<String> columnNamesStorage, XSSFSheet sheetHranenie, XSSFSheet sheetReportConverted) {
        for (Row row : sheetHranenie) {
            Cell cell = row.getCell(columnNamesStorage.indexOf("Предмет"));
            if (cell != null) {
                String categoryToCheck = cell.getStringCellValue();
                if (categoryToCheck.equals(category)) {
                    String article = row.getCell(columnNamesStorage.indexOf("Артикул продавца")).getStringCellValue().toLowerCase();
                    //делаем проверку, что артикул есть в ВводныеСебестоимость
                    if (!articlesInSebestoimost.contains(article)) {
                        if(!articlesNotInSebestoimost.contains(article)){
                            System.out.println("Артикул " + article + " отсутствует в ВводныеСебестоимость");
                            articlesNotInSebestoimost.add(article);
                        }
                    }

                    boolean containsIgnoreCase = articlesInCategory.stream().anyMatch(a -> a.equalsIgnoreCase(article));
                    if (!containsIgnoreCase) {
                        articlesInCategory.add(article);
                    }
                }
            }
        }
    }

    public static void initGeneralAndUnknown(XSSFSheet sheetReportConverted) {
        sheetReportConverted.createRow(2).createCell(0).setCellValue("Общий");
        sheetReportConverted.createRow(3).createCell(0).setCellValue("Неизвестно");
    }
}

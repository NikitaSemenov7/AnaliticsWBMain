package MatrixInitializer;

import enumsMaps.ColumnNames;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;

import static MatrixInitializer.Part1Initializer.*;


public class Part2Initializer {
    private static double sum = 0;

    public static void initialize2Part(double[][] matrix, ArrayList<String> rowNamesConverted, ArrayList<String> columnNamesDetalization, ArrayList<String> columnNamesStorage, XSSFSheet sheetHranenie, XSSFSheet sheetReportFromWB, XSSFSheet sheetSebestoimost) {
        //инициализируем хранение по категориям и товарам
        initHranenieCategoriesPositions(matrix, rowNamesConverted, columnNamesStorage, sheetHranenie, sheetSebestoimost);

        //добавляем пересчет Хранения в Неизвестно
        matrix[3][ColumnNames.Хранение.ordinal()] = initOneCondition(columnNamesDetalization.indexOf("Обоснование для оплаты"), "Пересчет хранения", columnNamesDetalization.indexOf("Хранение"), sheetReportFromWB);

        //инициализируем Хранение Общий
        initHranenieObshiy(matrix, sheetReportFromWB, columnNamesDetalization);
    }

    public static void initHranenieCategoriesPositions(double[][] matrix, ArrayList<String> rowNamesConverted, ArrayList<String> columnNamesStorage, XSSFSheet sheetHranenie, XSSFSheet sheetSebestoimost) {
        ArrayList<String> articlesHranennieInitialized = new ArrayList<>();
        for (int i = 0; i < rowNamesConverted.size(); i++) {
            sum = 0;
            String rowName = rowNamesConverted.get(i);

            //проверяем, что rowName является артикулом, а также проверяем, что по нему себестоимость еще не вставляли,
            //так как один и тот же артикул может присутствовать дважды в категориях с разными названиями, если категорию переименовывали
            if (articlesSebestoimost.contains(rowName) && !articlesHranennieInitialized.contains(rowName)) {
                //если товар
                for (Row row : sheetHranenie) {
                    String article = row.getCell(columnNamesStorage.indexOf("Артикул продавца")).getStringCellValue().toLowerCase();
                    if (article.equals(rowName)) {
                        Cell cellToSum = row.getCell(columnNamesStorage.indexOf("Сумма хранения, руб"));
                        if (cellToSum != null) {
                            double storage = cellToSum.getNumericCellValue();
                            sum += storage;
                        }
                    }
                }
                articlesHranennieInitialized.add(rowName);
            } else {
                //если категория, то в списке с себестоимостью находим всем артикулы по данной категории и суммируем хранение по артикулам
                //может быть такое, что в Детализации и в Хранении одна и та же категория подписана по-разному
                for (Row rowSebestoimost : sheetSebestoimost) {
                    Cell cellCategory = rowSebestoimost.getCell(0);
                    String categoryInSebestoimos = cellCategory.getStringCellValue().toLowerCase();
                    if (categoryInSebestoimos.equalsIgnoreCase(rowName)) {
                        Cell cellArticle = rowSebestoimost.getCell(1);
                        String articleInSebestoimost = cellArticle.getStringCellValue().toLowerCase();
                        for (Row rowHranenie : sheetHranenie) {
                            String articleHranenie = rowHranenie.getCell(columnNamesStorage.indexOf("Артикул продавца")).getStringCellValue();
                            if (articleHranenie.equalsIgnoreCase(articleInSebestoimost)) {
                                Cell cellToSum = rowHranenie.getCell(columnNamesStorage.indexOf("Сумма хранения, руб"));
                                if (cellToSum != null) {
                                    double storage = cellToSum.getNumericCellValue();
                                    sum += storage;
                                }
                            }
                        }
                    }

                }
            }
            //делаем проверку, так как бывает очень маленькое число с буквой е
            if (sum > 0.001) {
                matrix[i + 4][ColumnNames.Хранение.ordinal()] = sum;
            }
        }
    }

    public static void initHranenieObshiy(double[][] matrix, Sheet sheetReportFromWB, ArrayList<String> columnNamesDetalization) {
        sum = 0;
        //суммируем все в столбце хранение и вставляем в общий
        for (Row row: sheetReportFromWB) {
            Cell cellHranenie = row.getCell(columnNamesDetalization.indexOf("Хранение"));
            if (cellHranenie != null && cellHranenie.getCellType() == CellType.NUMERIC) {
                double hranenie = cellHranenie.getNumericCellValue();
                if (hranenie != 0) {
                    sum += hranenie;
                }
            }
        }
        matrix[2][ColumnNames.Хранение.ordinal()] = sum;
    }
}

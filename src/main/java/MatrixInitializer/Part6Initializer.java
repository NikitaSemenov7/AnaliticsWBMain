package MatrixInitializer;

import enumsMaps.Alphabet;
import enumsMaps.ColumnNames;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;

import static MatrixInitializer.Part1Initializer.articlesSebestoimost;

public class Part6Initializer {
    static double sum;


    public static void initialize6Part(double[][] matrix, ArrayList<String> rowNamesConverted, ArrayList<String> columnNamesSales, XSSFSheet sheetProdazhi, XSSFSheet sheetSebestoimost) {
        //инциализируем налог 7% по категориям и товарам и суммируем удвоенный общий налог
        ArrayList<String> articlesNalogInitialized = new ArrayList<>();
        for (int i = 0; i < rowNamesConverted.size(); i++) {
            String rowName = rowNamesConverted.get(i).toLowerCase();

            //проверяем, что только 1 раз посчитали налог по артикулу, так как один и тот же артикул может быть 2 раза в таблице, если
            //его категорию переименовывали
            if (articlesSebestoimost.contains(rowName) && !articlesNalogInitialized.contains(rowName)) {
                //если товар
                matrix[i + 4][ColumnNames.Налог7.ordinal()] = 0.07 * Part1Initializer.initOneCondition(columnNamesSales.indexOf("Артикул продавца"), rowName, columnNamesSales.indexOf("К перечислению за товар, руб."), sheetProdazhi);
                articlesNalogInitialized.add(rowName);
            } else {
                //если категория
                double categorySum = 0;
                for (Row rowSales: sheetProdazhi) {
                    Cell cellArticleSales = rowSales.getCell(columnNamesSales.indexOf("Артикул продавца"));
                    if(cellArticleSales != null) {
                        String articleSales = cellArticleSales.getStringCellValue();
                        for (Row rowSebestoimost: sheetSebestoimost) {
                            Cell cellArticleSebestoimost = rowSebestoimost.getCell(Alphabet.B.ordinal());
                            if (cellArticleSebestoimost != null) {
                                String articleSebestoimost = cellArticleSebestoimost.getStringCellValue();
                                if (articleSebestoimost.equalsIgnoreCase(articleSales)) {
                                    Cell cellCategorySebestoimost = rowSebestoimost.getCell(Alphabet.A.ordinal());
                                    String category = cellCategorySebestoimost.getStringCellValue();
                                    if(category.equalsIgnoreCase(rowName)) {
                                        Cell cellRevenue = rowSales.getCell(columnNamesSales.indexOf("К перечислению за товар, руб."));
                                        categorySum += cellRevenue.getNumericCellValue();
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
                matrix[i + 4][ColumnNames.Налог7.ordinal()] = 0.07 * categorySum;
            }
            sum += matrix[i + 4][ColumnNames.Налог7.ordinal()];
        }

        //инициализируем общий налог 7%
        matrix[2][ColumnNames.Налог7.ordinal()] = sum / 2;
    }
}

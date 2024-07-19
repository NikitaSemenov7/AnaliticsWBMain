package MatrixInitializer;

import enumsMaps.Alphabet;
import enumsMaps.ColumnNames;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;

import static MatrixInitializer.Part1Initializer.articlesSebestoimost;

public class Part4Initializer {
    public static void initialize4Part(double[][] matrix, ArrayList<String> rowNamesConverted, XSSFSheet sheetSebestoimost) {
        //инициализируем ПолученыДеньгиШтук, чтобы на основе их посчитать себестоимость
        for (int i = 0; i < rowNamesConverted.size() + 2; i++) {
            matrix[2+i][ColumnNames.ПолученыДеньгиШтук.ordinal()] =
                    matrix[2+i][ColumnNames.ПродажаШтук.ordinal()] -
                            matrix[2+i][ColumnNames.ВозвратШтук.ordinal()];
        }

        //инициализируем Себестоимость
        initialize(Alphabet.C.ordinal(), ColumnNames.Себестоимость.ordinal(), matrix, rowNamesConverted, sheetSebestoimost);

        //инициализируем фулфилмент
        initialize(Alphabet.D.ordinal(), ColumnNames.Фулфилмент.ordinal(), matrix, rowNamesConverted, sheetSebestoimost);
    }

    private static void initialize(int columnSebestoimost, int columnConverted, double[][] matrix, ArrayList<String> rowNamesConverted, XSSFSheet sheetSebestoimost) {
        //проходим по rowNamesConverted, инициализируем сначала позиции по артикулу и категории, так как один артикул может быть в 2
        //категориях, если категорию переименовывали
        String categoryRowNameConveted = "";
        for (int i = 0; i < rowNamesConverted.size(); i++) {
            double sebestoimost = 0;
            String rowName = rowNamesConverted.get(i);
            if (!articlesSebestoimost.contains(rowName)) {
                //если категория, то запоминаем в какой категории сейчас будем заполнять артикулы
                categoryRowNameConveted = rowName;
            } else {
                //если товар, то находим его в таблице с себестоимостью по артикулу и категории и вставляем себестоимость
                String articleRowNameConverted = rowName;
                for (Row rowSebestoimost: sheetSebestoimost) {
                    String categorySebestoimost = rowSebestoimost.getCell(Alphabet.A.ordinal()).getStringCellValue();
                    String articleSebestoimost = rowSebestoimost.getCell(Alphabet.B.ordinal()).getStringCellValue();
                    if (categorySebestoimost.equals(categoryRowNameConveted) && articleSebestoimost.equals(articleRowNameConverted)) {
                        sebestoimost = rowSebestoimost.getCell(columnSebestoimost).getNumericCellValue();
                    }
                }
            }
            double sum = sebestoimost * matrix[4 + i][ColumnNames.ПолученыДеньгиШтук.ordinal()];
            matrix[4 + i][columnConverted] = sum;
        }

        //проходим по rowNamesConverted, чтобы просуммировать себестоимость по категориям из себестоимости по позициям из матрицы
        double sumSebestoimostCategory = 0;
        double sumSebestoimostObshiy = 0;
        int categoryIndex = 0;
        for (int i = 0; i < rowNamesConverted.size(); i++) {
            String rowName = rowNamesConverted.get(i);
            if (!articlesSebestoimost.contains(rowName)) {
                //если категория
                //вставляем сумму себестоимости по категории и прибавляем ее к общей себестоимости
                matrix[4 + categoryIndex][columnConverted] = sumSebestoimostCategory;
                sumSebestoimostObshiy += sumSebestoimostCategory;
                //обнуляем переменную, в которую будем считать сумму себестоимости по категории
                sumSebestoimostCategory = 0;
                //запоминаем индекс строки, в которой категория
                categoryIndex = i;
            } else {
                //если позиция, то суммируем себестоимость по ней в переменную
                sumSebestoimostCategory += matrix[4 + i][columnConverted];
            }
        }
        //вставляем сумму себестоиомсти по последней категории и прибавляем ее к общей категории
        matrix[4 + categoryIndex][columnConverted] = sumSebestoimostCategory;
        sumSebestoimostObshiy += sumSebestoimostCategory;

        //инициализируем Себестоимость Общий
        matrix[2][columnConverted] = sumSebestoimostObshiy;
    }
}

package MatrixInitializer;

import enumsMaps.ColumnNames;

import java.util.ArrayList;

public class Part7Initializer {
    public static void initialize7Part(double[][] matrix, ArrayList<String> rowNamesConverted){
        for (int i = 0; i < rowNamesConverted.size() + 2; i++) {
            matrix[2 + i][ColumnNames.ВыручкаДоВычетаКомиссии.ordinal()] =
                    matrix[2 + i][ColumnNames.ПродажаДоВычетаКомиссии.ordinal()] -
                            matrix[2 + i][ColumnNames.ВозвратДоВычетаКомиссии.ordinal()];

            matrix[2 + i][ColumnNames.ВыручкаПослеВычетаКомиссии.ordinal()] =
                    matrix[2 + i][ColumnNames.ПродажаПослеВычетаКомиссии.ordinal()] -
                            matrix[2 + i][ColumnNames.ВозвратПослеВычетаКомиссии.ordinal()];


            //выручка до вычета комиссии на 1 штуку
            double полученыДеньгиШтук = matrix[2 + i][ColumnNames.ПолученыДеньгиШтук.ordinal()];
            if(полученыДеньгиШтук != 0) {
                matrix[2 + i][ColumnNames.ВыручкаДоВычетаКомиссииНа1шт.ordinal()] =
                        matrix[2 + i][ColumnNames.ВыручкаДоВычетаКомиссии.ordinal()] / полученыДеньгиШтук;
            } else {
                matrix[2 + i][ColumnNames.ВыручкаДоВычетаКомиссииНа1шт.ordinal()] = 0;
            }

            //выручка после вычета комиссии на 1 штуку
            if(полученыДеньгиШтук != 0) {
                matrix[2 + i][ColumnNames.ВыручкаПослеВычетаКомиссииНа1шт.ordinal()] =
                        matrix[2 + i][ColumnNames.ВыручкаПослеВычетаКомиссии.ordinal()] / полученыДеньгиШтук;
            } else {
                matrix[2 + i][ColumnNames.ВыручкаПослеВычетаКомиссииНа1шт.ordinal()] = 0;
            }

            matrix[2 + i][ColumnNames.Логистика.ordinal()] =
                    matrix[2 + i][ColumnNames.ДоставкаККлиенту.ordinal()] +
                            matrix[2 + i][ColumnNames.ДоставкаОтКлиента.ordinal()] +
                            matrix[2 + i][ColumnNames.КоррекцияЛогистики.ordinal()];

            matrix[2 + i][ColumnNames.ПриходНаСчетВб.ordinal()] =
                    matrix[2 + i][ColumnNames.ВыручкаПослеВычетаКомиссии.ordinal()] -
                            matrix[2 + i][ColumnNames.Логистика.ordinal()] -
                            matrix[2 + i][ColumnNames.Штрафы.ordinal()] -
                            matrix[2 + i][ColumnNames.Хранение.ordinal()] -
                            matrix[2 + i][ColumnNames.ПлатнаяПриемка.ordinal()] -
                            matrix[2 + i][ColumnNames.ПрочиеУдержания.ordinal()];

            matrix[2 + i][ColumnNames.ПрибыльНаМаркетинг.ordinal()] =
                    matrix[2 + i][ColumnNames.ПриходНаСчетВб.ordinal()] -
                            matrix[2 + i][ColumnNames.Себестоимость.ordinal()] -
                            matrix[2 + i][ColumnNames.Фулфилмент.ordinal()] -
                            matrix[2 + i][ColumnNames.Налог7.ordinal()];

            if(полученыДеньгиШтук != 0) {
                matrix[2 + i][ColumnNames.ПрибыльНаМаркетингНа1штуку.ordinal()] =
                        matrix[2 + i][ColumnNames.ПрибыльНаМаркетинг.ordinal()] / полученыДеньгиШтук;
            } else {
                matrix[2 + i][ColumnNames.ПрибыльНаМаркетингНа1штуку.ordinal()] = 0;
            }
        }
    }
}

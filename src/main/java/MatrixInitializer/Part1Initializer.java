package MatrixInitializer;

import enumsMaps.Alphabet;
import enumsMaps.ColumnNames;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;

public class Part1Initializer {
    private static double sum;
    public static ArrayList<String> articlesSebestoimost = new ArrayList<>();

    public static void initialize1Part(double[][] matrix, ArrayList<String> rowNamesConverted, ArrayList<String> columnNamesDetalization, XSSFSheet sheetReportFromWB, XSSFSheet sheetSebestoimost, XSSFSheet sheetReportConverted) {

        //наполняем articlesSebestoimost
        for(Row row: sheetSebestoimost) {
            Cell cell = row.getCell(Alphabet.B.ordinal());
            if(cell != null && cell.getCellType() == CellType.STRING && !cell.getStringCellValue().equals("Артикул поставщика")) {
                articlesSebestoimost.add(cell.getStringCellValue());
            }
        }

        //ПродажаДоВычетаКомиссии
        initialize(matrix, rowNamesConverted, articlesSebestoimost, columnNamesDetalization, sheetReportFromWB, ColumnNames.ПродажаДоВычетаКомиссии.ordinal(), columnNamesDetalization.indexOf("Вайлдберриз реализовал Товар (Пр)"), columnNamesDetalization.indexOf("Тип документа"), "Продажа");
        //ПродажаШтук
        initialize(matrix, rowNamesConverted, articlesSebestoimost, columnNamesDetalization, sheetReportFromWB, ColumnNames.ПродажаШтук.ordinal(), columnNamesDetalization.indexOf("Кол-во"), columnNamesDetalization.indexOf("Тип документа"), "Продажа");
        //ВозвратДоВычетаКомиссии
        initialize(matrix, rowNamesConverted, articlesSebestoimost, columnNamesDetalization, sheetReportFromWB, ColumnNames.ВозвратДоВычетаКомиссии.ordinal(), columnNamesDetalization.indexOf("Вайлдберриз реализовал Товар (Пр)"), columnNamesDetalization.indexOf("Тип документа"), "Возврат");
        //ВозвратШтук
        initialize(matrix, rowNamesConverted, articlesSebestoimost, columnNamesDetalization, sheetReportFromWB, ColumnNames.ВозвратШтук.ordinal(), columnNamesDetalization.indexOf("Кол-во"), columnNamesDetalization.indexOf("Тип документа"), "Возврат");
        //ПродажаПослеВычетаКомиссии
        initialize(matrix, rowNamesConverted, articlesSebestoimost, columnNamesDetalization, sheetReportFromWB, ColumnNames.ПродажаПослеВычетаКомиссии.ordinal(), columnNamesDetalization.indexOf("К перечислению Продавцу за реализованный Товар"), columnNamesDetalization.indexOf("Тип документа"), "Продажа");
        //ВозвратПослеВычетаКомиссии
        initialize(matrix, rowNamesConverted, articlesSebestoimost, columnNamesDetalization, sheetReportFromWB, ColumnNames.ВозвратПослеВычетаКомиссии.ordinal(), columnNamesDetalization.indexOf("К перечислению Продавцу за реализованный Товар"), columnNamesDetalization.indexOf("Тип документа"), "Возврат");
        //КоррекцияПродаж
        initialize(matrix, rowNamesConverted, articlesSebestoimost, columnNamesDetalization, sheetReportFromWB, ColumnNames.КоррекцияПродаж.ordinal(), columnNamesDetalization.indexOf("К перечислению Продавцу за реализованный Товар"), columnNamesDetalization.indexOf("Обоснование для оплаты"), "Коррекция продаж");
        //ДоставкаККлиенту
        initialize(matrix, rowNamesConverted, articlesSebestoimost, columnNamesDetalization, sheetReportFromWB, ColumnNames.ДоставкаККлиенту.ordinal(), columnNamesDetalization.indexOf("Услуги по доставке товара покупателю"), columnNamesDetalization.indexOf("Количество доставок"), "1");
        //ДоставкаККлиентуШтук
        initialize(matrix, rowNamesConverted, articlesSebestoimost, columnNamesDetalization, sheetReportFromWB, ColumnNames.ДоставкаККлиентуШтук.ordinal(), columnNamesDetalization.indexOf("Количество доставок"), columnNamesDetalization.indexOf("Количество доставок"), "1");
        //ДоставкаОтКлиента
        initialize(matrix, rowNamesConverted, articlesSebestoimost, columnNamesDetalization, sheetReportFromWB, ColumnNames.ДоставкаОтКлиента.ordinal(), columnNamesDetalization.indexOf("Услуги по доставке товара покупателю"), columnNamesDetalization.indexOf("Количество возврата"), "1");
        //ДоставкаОтКлиентаШтук
        initialize(matrix, rowNamesConverted, articlesSebestoimost, columnNamesDetalization, sheetReportFromWB, ColumnNames.ДоставкаОтКлиентаШтук.ordinal(), columnNamesDetalization.indexOf("Количество возврата"), columnNamesDetalization.indexOf("Количество возврата"), "1");
        //КоррекцияЛогистики
        initialize(matrix, rowNamesConverted, articlesSebestoimost, columnNamesDetalization, sheetReportFromWB, ColumnNames.КоррекцияЛогистики.ordinal(), columnNamesDetalization.indexOf("Услуги по доставке товара покупателю"), columnNamesDetalization.indexOf("Обоснование для оплаты"), "Коррекция логистики");
    }

    public static double initOneCondition(int columnToCheck, String condition, int columnToSum, Sheet sheetFrom) {
        sum = 0;

        //получаем значение для проверки
        String cellToCheckValue;
        for (Row row : sheetFrom) {
            Cell cellToChek = row.getCell(columnToCheck);
            if (cellToChek != null) {
                CellType cellType = cellToChek.getCellType();

                if (cellType == CellType.STRING) {
                    cellToCheckValue = cellToChek.getStringCellValue();
                } else {
                    cellToCheckValue = (int) cellToChek.getNumericCellValue() + "";
                }

                //проверяем верно ли условие
                //startsWith написано, потому что могут быть "Штраф" и "Штрафы"
                if (cellToCheckValue.equalsIgnoreCase(condition) || cellToCheckValue.startsWith(condition)) {
                    Cell cellToSum = row.getCell(columnToSum);
                    if (cellToSum != null && cellToSum.getCellType() == CellType.NUMERIC) {
                        sum += cellToSum.getNumericCellValue();
                    }
                }
            }
        }
        return sum;
    }


    public static double initTwoConditions(int columnToCheck1, String condition1, int columnToCheck2, String condition2, int columnToSum, ArrayList<String> columnNamesDetalization, int initializationColumn, XSSFSheet sheetReportFromWB) {
        sum = 0;

        //получаем значение для проверки 1
        String cellToCheckValue1;
        for (Row row : sheetReportFromWB) {
            Cell cellToCheck1 = row.getCell(columnToCheck1);
            if (cellToCheck1 == null)
                continue;
            else {
                CellType cellType1 = cellToCheck1.getCellType();

                if (cellType1 == CellType.STRING) {
                    cellToCheckValue1 = cellToCheck1.getStringCellValue();
                } else {
                    cellToCheckValue1 = (int) cellToCheck1.getNumericCellValue() + "";
                }
            }

            //получаем значение для проверки 2
            String cellToCheckValue2;
            Cell cellToCheck2 = row.getCell(columnToCheck2);
            if (cellToCheck2 == null)
                continue;
            else {
                CellType cellType2 = cellToCheck2.getCellType();

                if (cellType2 == CellType.STRING) {
                    cellToCheckValue2 = cellToCheck2.getStringCellValue();
                } else {
                    cellToCheckValue2 = (int) cellToCheck2.getNumericCellValue() + "";
                }
            }

            boolean condition3boolean = true;
            if(initializationColumn == ColumnNames.ПродажаШтук.ordinal() || initializationColumn == ColumnNames.ВозвратШтук.ordinal()) {
                int index = columnNamesDetalization.indexOf("Обоснование для оплаты");
                Cell cellPaymentReason = row.getCell(index);
                if (cellPaymentReason != null) {
                    String paymentReason = cellPaymentReason.getStringCellValue();
                    if(paymentReason.equals("Частичная компенсация брака")) {
                        condition3boolean = false;
                    }
                }
            }

            Cell cellToSum = row.getCell(columnToSum);
            if (cellToCheckValue1.equalsIgnoreCase(condition1)
                    && cellToCheckValue2.equalsIgnoreCase(condition2)
                    && cellToSum != null && cellToSum.getCellType() == CellType.NUMERIC
                    && condition3boolean) {
                sum += cellToSum.getNumericCellValue();
            }
        }
        return sum;
    }

    public static double initThreeConditions(int columnToCheck1, String condition1, int columnToCheck2, String condition2, int columnToCheck3, String condition3, int columnToSum, ArrayList<String> columnNamesDetalization, int initializationColumn, XSSFSheet sheetReportFromWB) {
        sum = 0;
        String cellToCheckValue1;
        String cellToCheckValue2;
        String cellToCheckValue3;
        for (Row row : sheetReportFromWB) {
            //получаем значение для проверки 1
            Cell cellToCheck1 = row.getCell(columnToCheck1);
            if (cellToCheck1 == null)
                continue;
            else {
                CellType cellType1 = cellToCheck1.getCellType();

                if (cellType1 == CellType.STRING) {
                    cellToCheckValue1 = cellToCheck1.getStringCellValue();
                } else {
                    cellToCheckValue1 = (int) cellToCheck1.getNumericCellValue() + "";
                }
            }

            //получаем значение для проверки 2
            Cell cellToCheck2 = row.getCell(columnToCheck2);
            if (cellToCheck2 == null)
                continue;
            else {
                CellType cellType2 = cellToCheck2.getCellType();

                if (cellType2 == CellType.STRING) {
                    cellToCheckValue2 = cellToCheck2.getStringCellValue();
                } else {
                    cellToCheckValue2 = (int) cellToCheck2.getNumericCellValue() + "";
                }
            }

            //получаем значение для проверки 3
            Cell cellToCheck3 = row.getCell(columnToCheck3);
            if (cellToCheck3 == null)
                continue;
            else {
                CellType cellType3 = cellToCheck3.getCellType();

                if (cellType3 == CellType.STRING) {
                    cellToCheckValue3 = cellToCheck3.getStringCellValue();
                } else {
                    cellToCheckValue3 = (int) cellToCheck3.getNumericCellValue() + "";
                }
            }

            //если частичная компенсация брака, то ПродажаШтук или ВозвратШтук прибавлять не будем
            boolean condition4boolean = true;
            if(initializationColumn == ColumnNames.ПродажаШтук.ordinal() || initializationColumn == ColumnNames.ВозвратШтук.ordinal()) {
                int index = columnNamesDetalization.indexOf("Обоснование для оплаты");
                Cell cellPaymentReason = row.getCell(index);
                if (cellPaymentReason != null) {
                    String paymentReason = cellPaymentReason.getStringCellValue();
                    if(paymentReason.equals("Частичная компенсация брака")) {
                        condition4boolean = false;
                    }
                }
            }

            Cell cellToSum = row.getCell(columnToSum);
            if (cellToCheckValue1.equalsIgnoreCase(condition1)
                    && cellToCheckValue2.equalsIgnoreCase(condition2)
                    && cellToCheckValue3.equalsIgnoreCase(condition3)
                    && cellToSum != null && cellToSum.getCellType() == CellType.NUMERIC
                    && condition4boolean) {
                sum += cellToSum.getNumericCellValue();
            }
        }
        return sum;
    }

    public static void initialize(double[][] matrix, ArrayList<String> rowNamesConverted, ArrayList<String> articles, ArrayList<String> columnNamesDetalization,XSSFSheet sheetReportFromWB, int initializationColumn, int columnToSum, int columnToCheck, String condition) {
        //инициализируем по категориям и товарам
        String category = "";
        for (int i = 0; i < rowNamesConverted.size(); i++) {
            String rowName = rowNamesConverted.get(i).toLowerCase();
            //если товар
            if (articles.contains(rowName)) {
                //проверяем по артикулу и категории, а не только по артикулу, так как, если категорию переименуют, может быть 1 и тот же артикул
                //с разными названиями категорий
                matrix[4 + i][initializationColumn] = initThreeConditions(columnNamesDetalization.indexOf("Артикул поставщика"), rowName, columnToCheck, condition, columnNamesDetalization.indexOf("Предмет"), category, columnToSum, columnNamesDetalization, initializationColumn, sheetReportFromWB);
            } else {
                //если категория
                matrix[4 + i][initializationColumn] = initTwoConditions(columnNamesDetalization.indexOf("Предмет"), rowName, columnToCheck, condition, columnToSum, columnNamesDetalization, initializationColumn, sheetReportFromWB);
                category = rowName;
            }
        }

        initializeObshiySum(rowNamesConverted, matrix, initializationColumn);
    }

    public static void initializeObshiySum(ArrayList<String> rowNamesConverted, double[][] matrix, int initializationColumn) {
        double sum = 0;
        for (int i = 0; i < rowNamesConverted.size(); i++) {
            sum += matrix[4 + i][initializationColumn];
        }
        matrix[2][initializationColumn] = sum / 2;
    }
}

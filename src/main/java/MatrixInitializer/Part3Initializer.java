package MatrixInitializer;


import enumsMaps.Alphabet;
import enumsMaps.ColumnNames;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;

import static MatrixInitializer.Part1Initializer.articlesSebestoimost;

public class Part3Initializer {

    private static double sum = 0;

    public static void initialize3Part(double[][] matrix, ArrayList<String> columnNamesConverted, ArrayList<String> columnNamesDetalization, ArrayList<String> columnNamesProchieUderzhaniya, XSSFSheet sheetReportFromWB, XSSFSheet sheetProchieUderzhaniya, XSSFSheet sheetNazvaniyaRK, String directoryIntroductories) {
        //получаем список номеров рекламных кампаний
        ArrayList<Integer> listActualDocuments = getListActualDocumentsRK(columnNamesDetalization, columnNamesProchieUderzhaniya, sheetReportFromWB, sheetProchieUderzhaniya);
        //проверяем, что не появились новые рекламные кампании, которых нет в ВводныеНазванияРК
        checkNewRK(listActualDocuments, columnNamesProchieUderzhaniya, sheetNazvaniyaRK, sheetProchieUderzhaniya, directoryIntroductories);

        initProchieUderzhaniyaObshiyUnknown(matrix, columnNamesDetalization, sheetReportFromWB);
        initProchieUderzhaniyaCategories(matrix, columnNamesConverted, listActualDocuments, columnNamesProchieUderzhaniya, sheetProchieUderzhaniya, sheetReportFromWB, sheetNazvaniyaRK);
        initProchieUderzhaniyaPositions(matrix, columnNamesConverted);
    }

    public static void initProchieUderzhaniyaPositions(double[][] matrix, ArrayList<String> columnNamesConverted) {
        //проходи по списку columnNamesConverted
        for (int i = 0; i < columnNamesConverted.size(); i++) {
            String columnName = columnNamesConverted.get(i);
            //если категория
            if (!articlesSebestoimost.contains(columnName)) {
                //считываем из матрицы сумму рекламных расходов в этой категории
                double categorySum = matrix[i + 4][ColumnNames.ПрочиеУдержания.ordinal()];
                //считаем количество артикулов в этой категории из columnNamesConverted
                int articlesNumber = 0;
                String nextColumnName = "";
                int j = i;
                while (true) {
                    if (j + 1 >= columnNamesConverted.size()) {
                        break;
                    }
                    nextColumnName = columnNamesConverted.get(j + 1);
                    if (articlesSebestoimost.contains(nextColumnName)) {
                        articlesNumber++;
                        j++;
                    } else {
                        break;
                    }
                }
                //считаем сумму прочих удержаний по категории / Продажа штук
                double prodazhaShtuk = matrix[i + 4][ColumnNames.ПродажаШтук.ordinal()];
                if (prodazhaShtuk != 0) {
                    double prochieUderzhaniyaNa1Shtuku = categorySum / prodazhaShtuk;
                    //расставляем взвешенные прочие удержания по товарам
                    for (int k = 1; k <= articlesNumber; k++) {
                        matrix[i + 4 + k][ColumnNames.ПрочиеУдержания.ordinal()] = prochieUderzhaniyaNa1Shtuku * matrix[i + 4 + k][ColumnNames.ПродажаШтук.ordinal()];
                    }
                } else {
                    for (int k = 1; k <= articlesNumber; k++) {
                        matrix[i + 4 + k][ColumnNames.ПрочиеУдержания.ordinal()] = categorySum / articlesNumber;
                    }
                }
            }
        }
    }

    public static void initProchieUderzhaniyaCategories(double[][] matrix, ArrayList<String> rowNamesConverted, ArrayList<Integer> listActualDocuments, ArrayList<String> columnNamesProchieUderzhaniya, XSSFSheet sheetProchieUderzhaniya, XSSFSheet sheetReportFromWB, XSSFSheet sheetNazvaniyaRK) {
        for (int i = 0; i < rowNamesConverted.size(); i++) {
            sum = 0;
            String rowName = rowNamesConverted.get(i);
            //если rowName категория
            if (!articlesSebestoimost.contains(rowName)) {
                //проходим по всей таблице с рекламой и проверяем 2 условия
                for (Row row : sheetProchieUderzhaniya) {
                    //проверяем, что категория рекламной кампании такая же как rowName
                    boolean condition1 = checkCategory(rowName, sheetProchieUderzhaniya, row, sheetNazvaniyaRK);
                    if (condition1) {
                        //проверяем, что номер РК есть в listActualDocuments
                        boolean condition2 = checkRkNumber(row, listActualDocuments, columnNamesProchieUderzhaniya);
                        if (condition2) {
                            //прибавляем к сумме, если 2 условия выполнились
                            double numericCellValue = row.getCell(columnNamesProchieUderzhaniya.indexOf("Сумма")).getNumericCellValue();
                            sum += numericCellValue;
                        }
                    }
                }
                matrix[i + 4][ColumnNames.ПрочиеУдержания.ordinal()] = sum;
            }
        }
    }

    private static boolean checkRkNumber(Row row, ArrayList<Integer> listActualDocuments, ArrayList<String> columnNamesProchieUderzhaniya) {
        Cell cellRkNumber = row.getCell(columnNamesProchieUderzhaniya.indexOf("Номер документа"));
        if (cellRkNumber != null) {
            if (cellRkNumber.getCellType() == CellType.NUMERIC) {
                int rkNumber = (int) cellRkNumber.getNumericCellValue();
                if (listActualDocuments.contains(rkNumber)) {
                    return true;
                }
            }
        }
        return false;
    }

    private static boolean checkCategory(String rowName, XSSFSheet sheetProchieUderzhaniya, Row row, XSSFSheet sheetNazvaniyaRK) {
        boolean firstCondition = false;
        //в переданном ряде берем название рекламой кампании из столбца B
        Cell cellRkName = row.getCell(Alphabet.B.ordinal());
        if (cellRkName != null) {
            String rkName = cellRkName.getStringCellValue();
            //проходим по всей таблице ВводныеНазванияРК
            for (Row rowNazvaniyaRK : sheetNazvaniyaRK) {
                String rkNameVvodnye = rowNazvaniyaRK.getCell(Alphabet.A.ordinal()).getStringCellValue();
                //в таблице ВводныеНазванияРК находим строку, в которой находится рассматриваемая РК
                if (rkNameVvodnye.equalsIgnoreCase(rkName)) {
                    //узнаем категорию соответствующую рассматриваемой РК
                    String category = rowNazvaniyaRK.getCell(Alphabet.B.ordinal()).getStringCellValue();
                    firstCondition = category.equalsIgnoreCase(rowName);
                    if (firstCondition == true) {
                        return firstCondition;
                    }
                }
            }
        }
        return firstCondition;
    }


    public static ArrayList<Integer> getListActualDocumentsRK(ArrayList<String> columnNamesDetalization, ArrayList<String> columnNamesProchieUderzhaniya, XSSFSheet sheetReportFromWB, XSSFSheet sheetProchieUderzhaniya) {

        //получаем суммы по прочим удержаниям из детализации
        ArrayList<Integer> listActualDocumentsSums = new ArrayList<>();
        for (Row row : sheetReportFromWB) {
            Cell cellUderzhaniya = row.getCell(columnNamesDetalization.indexOf("Удержания"));
            if (row.getRowNum() != 0 && cellUderzhaniya != null && cellUderzhaniya.getCellType() == CellType.NUMERIC) {
                sum = cellUderzhaniya.getNumericCellValue();
                Cell cellReasonUderzhaniya = row.getCell(columnNamesDetalization.indexOf("Виды логистики, штрафов и доплат"));
                //дописал cellReasonUderzhaniya == null, так как в ранних периодах не проставлено ничего
                //sum > 0, так как бывают отрицательные числа, которые не относятся к рекламе
                if (sum > 0 && (cellReasonUderzhaniya == null ||
                        cellReasonUderzhaniya.getStringCellValue().equals("Оказание услуг «ВБ.Продвижение»"))) {
                    listActualDocumentsSums.add((int) sum);
                }
            }
        }

        //получаем все номера документов из таблицы с рекламой
        ArrayList<Integer> listAllDocuments = new ArrayList<>();
        for (Row row : sheetProchieUderzhaniya) {
            if (row.getRowNum() != 0) {
                Cell cellDocNum = row.getCell(columnNamesProchieUderzhaniya.indexOf("Номер документа"));
                if (cellDocNum.getCellType() == CellType.STRING) {
                    break;
                }
                int docNumber = (int) cellDocNum.getNumericCellValue();
                if (!listAllDocuments.contains(docNumber)) {
                    listAllDocuments.add(docNumber);
                }
            }
        }

        //в таблице с рекламой высчитываем суммы соответствующие всем номерам документов
        ArrayList<Integer> listAllDocumentsSums = new ArrayList<>();
        for (Integer listAllDocument : listAllDocuments) {
            sum = 0;
            for (Row row : sheetProchieUderzhaniya) {
                if (row.getRowNum() != 0) {
                    Cell cellToCheck = row.getCell(columnNamesProchieUderzhaniya.indexOf("Номер документа"));
                    int docNum = (int) cellToCheck.getNumericCellValue();
                    if (listAllDocument.equals(docNum)) {
                        Cell cellToSum = row.getCell(columnNamesProchieUderzhaniya.indexOf("Сумма"));
                        int rashod = (int) cellToSum.getNumericCellValue();
                        sum += rashod;
                    }
                }
            }
            listAllDocumentsSums.add((int) sum);
        }

        //получаем список с номерами документов, суммы по которым соответствуют суммам из детализации
        ArrayList<Integer> listActualDocuments = new ArrayList<>();
        for (Integer actualDocumentsSum : listActualDocumentsSums) {
            for (int j = 0; j < listAllDocumentsSums.size(); j++) {
                if (actualDocumentsSum.equals(listAllDocumentsSums.get(j))) {
                    listActualDocuments.add(listAllDocuments.get(j));
                    break;
                }
            }
        }

        //если определилось недостаточное количество actualDocuments, то выводим сообщение об ошибке
        if (listActualDocuments.size() < listActualDocumentsSums.size()) {
            System.out.println("Определилось недостаточное количество номеров документов по рекламе");
        }

        return listActualDocuments;
    }

    public static void checkNewRK(ArrayList<Integer> listActualDocuments, ArrayList<String> columnNamesProchieUderzhaniya, XSSFSheet sheetNazvaniyaRK, XSSFSheet sheetProchieUderzhaniya, String directoryIntroductories) {
        //получаем список РК из вводных
        ArrayList<String> listRKNamesVvodnye = new ArrayList<>();
        int counter = 0;
        while (sheetNazvaniyaRK.getRow(counter) != null) {
            String rkNameVvodnye = sheetNazvaniyaRK.getRow(counter).getCell(Alphabet.A.ordinal()).getStringCellValue();
            listRKNamesVvodnye.add(rkNameVvodnye);
            counter++;
        }
        //получаем список РК из отчета
        ArrayList<String> listRkNamesReport = new ArrayList<>();
        counter = 1;
        while (sheetProchieUderzhaniya.getRow(counter) != null) {
            Cell cell1 = sheetProchieUderzhaniya.getRow(counter).getCell(columnNamesProchieUderzhaniya.indexOf("Номер документа"));
            if (cell1.getCellType() != CellType.STRING) {
                int documentNumber = (int) cell1.getNumericCellValue();
                if (listActualDocuments.contains(documentNumber)) {
                    String rkNameReport = sheetProchieUderzhaniya.getRow(counter).getCell(columnNamesProchieUderzhaniya.indexOf("Кампания")).getStringCellValue();
                    if (!listRkNamesReport.contains(rkNameReport)) {
                        listRkNamesReport.add(rkNameReport);
                    }
                }
            }
            counter++;
        }
        //сравниваем список РК из вводных и список РК из отчета
        for (String rk : listRkNamesReport) {
            if (!listRKNamesVvodnye.contains(rk)) {
                System.out.println("Рекламной кампании нет в ВводныеНазванияРК " + rk);
            }
        }
    }

    public static void initProchieUderzhaniyaObshiyUnknown(double[][] matrix, ArrayList<String> columnNamesDetalization, XSSFSheet sheetReportFromWB) {
        sum = 0;
        double unknownSum = 0;
        for (Row row : sheetReportFromWB) {
            Cell cell = row.getCell(columnNamesDetalization.indexOf("Удержания"));
            double uderzhaniya = 0;
            if (cell != null && cell.getCellType() == CellType.NUMERIC
                    && (uderzhaniya = cell.getNumericCellValue()) != 0) {
                sum += uderzhaniya;
                //если значение в столбце Виды логистики, штрафов и доплат != Оказание услуг «ВБ.Продвижение» или удержание
                //отрицательное, то складываем в unknownSum
                Cell cellReasonUderzhaniya = row.getCell(columnNamesDetalization.indexOf("Виды логистики, штрафов и доплат"));
                if ((cellReasonUderzhaniya != null
                        && !cellReasonUderzhaniya.getStringCellValue().equals("Оказание услуг «ВБ.Продвижение»"))
                        || uderzhaniya < 0)
                    unknownSum += uderzhaniya;
            }
        }
        matrix[2][ColumnNames.ПрочиеУдержания.ordinal()] = sum;
        matrix[3][ColumnNames.ПрочиеУдержания.ordinal()] = unknownSum;
    }
}

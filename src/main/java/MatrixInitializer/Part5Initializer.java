package MatrixInitializer;

import enumsMaps.Alphabet;
import enumsMaps.ColumnNames;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.ArrayList;

public class Part5Initializer {
    public static void initialize5Part(double[][] matrix, ArrayList<String> columnNamesDetalization, XSSFSheet sheetReportFromWB){
        initShtrafi(matrix, columnNamesDetalization, sheetReportFromWB);
        initPlatnayaPriemka(matrix, columnNamesDetalization, sheetReportFromWB);
    }

    public static void initShtrafi(double[][] matrix, ArrayList<String> columnNamesDetalization, XSSFSheet sheetReportFromWB){
        matrix[2][ColumnNames.Штрафы.ordinal()] = Part1Initializer.initOneCondition(columnNamesDetalization.indexOf("Обоснование для оплаты"), "Штраф", columnNamesDetalization.indexOf("Общая сумма штрафов"), sheetReportFromWB);
        matrix[3][ColumnNames.Штрафы.ordinal()] = matrix[2][ColumnNames.Штрафы.ordinal()];
    }

    public static void initPlatnayaPriemka(double[][] matrix, ArrayList<String> columnNamesDetalization, XSSFSheet sheetReportFromWB){
        double sumPlatnayaPriemka = 0;
        for(Row row: sheetReportFromWB) {
            Cell cellPlatnayaPriemka = row.getCell(columnNamesDetalization.indexOf("Платная приемка"));
            if(cellPlatnayaPriemka != null) {
                if(cellPlatnayaPriemka.getCellType() == CellType.NUMERIC) {
                    double platnayaPriemka = cellPlatnayaPriemka.getNumericCellValue();
                    sumPlatnayaPriemka += platnayaPriemka;
                }
            }
        }
        matrix[2][ColumnNames.ПлатнаяПриемка.ordinal()] = sumPlatnayaPriemka;
        matrix[3][ColumnNames.ПлатнаяПриемка.ordinal()] = sumPlatnayaPriemka;
    }
}

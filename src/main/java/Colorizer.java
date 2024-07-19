import enumsMaps.ColumnNames;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;

import static MatrixInitializer.Part1Initializer.articlesSebestoimost;

public class Colorizer {
    public static CellStyle styleGreen;
    public static CellStyle styleLightGrey;
    public static CellStyle styleLightBlue;

    public static void colorize(ArrayList<String> columnNames, XSSFSheet sheetReportConverted, XSSFWorkbook reportConverted) {
        //создание зеленого стиля для заливки ячейки
        createStyleGreen(reportConverted);
        createStyleGrey(reportConverted);
        createStyleLightGrey(reportConverted);

        //раскрашиваем в голубой строку Общий
        fillRowBlue(2, sheetReportConverted);

        //раскрашиваем в светло-серый строки с категориями
        for (int i = 0; i < columnNames.size(); i++) {
            String columnName = columnNames.get(i);
            if (!articlesSebestoimost.contains(columnName)) {
                fillRowLightGrey(4 + i, sheetReportConverted);
            }
        }

        //раскрашиваем в зеленый важные столбцы
        fillColumnGreen(ColumnNames.ВыручкаДоВычетаКомиссии.ordinal(), sheetReportConverted);
        fillColumnGreen(ColumnNames.ВыручкаПослеВычетаКомиссии.ordinal(), sheetReportConverted);
        fillColumnGreen(ColumnNames.ПолученыДеньгиШтук.ordinal(), sheetReportConverted);
        fillColumnGreen(ColumnNames.ПриходНаСчетВб.ordinal(), sheetReportConverted);
        fillColumnGreen(ColumnNames.ПрибыльНаМаркетинг.ordinal(), sheetReportConverted);
    }

    public static void createStyleGreen(XSSFWorkbook reportConverted) {
        styleGreen = reportConverted.createCellStyle();
        styleGreen.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
        styleGreen.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    public static void createStyleLightGrey(XSSFWorkbook reportConverted) {
        styleLightGrey = reportConverted.createCellStyle();
        styleLightGrey.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        styleLightGrey.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    public static void createStyleGrey(XSSFWorkbook reportConverted) {
        styleLightBlue = reportConverted.createCellStyle();
        styleLightBlue.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());
        styleLightBlue.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    public static void fillColumnGreen(int columnNum, XSSFSheet sheetReportConverted) {
        for (Row row: sheetReportConverted){
            Cell cell = row.getCell(columnNum);
            if (cell != null) {
                cell.setCellStyle(styleGreen);
            }
        }
    }

    public static void fillRowBlue(int rowNum, XSSFSheet sheetReportConverted) {
        Row row = sheetReportConverted.getRow(rowNum);
        for (Cell cell: row) {
            if (cell != null) {
                cell.setCellStyle(styleLightBlue);
            }
        }
    }

    public static void fillRowLightGrey(int rowNum, XSSFSheet sheetReportConverted) {
        Row row = sheetReportConverted.getRow(rowNum);
        for (Cell cell : row) {
            if (cell != null) {
                cell.setCellStyle(styleLightGrey);
            }
        }
    }
}

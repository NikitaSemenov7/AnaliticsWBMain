import API.AdsApiDownload;
import API.DetalizationApiDownload;
import API.SalesApiDownload;
import API.StorageApiDownload;
import MatrixInitializer.*;
import StringInitializers.RowNamesConvertedInitializer;
import StringInitializers.DatesInitializer;
import StringInitializers.ColumnNamesInitializer;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.ParseException;
import java.util.ArrayList;
import java.time.temporal.WeekFields;
import java.util.Locale;
import java.time.LocalDate;


public class ExcelConverter {
    public static int counter = 0;
    public static double sum;
    public static XSSFSheet sheetReportConverted;
    public static XSSFSheet sheetNazvaniyaRK;
    public static XSSFSheet sheetSebestoimost;
    public static ArrayList<String> rowNamesConverted;


    public static void main(String[] args) throws IOException, InterruptedException, ParseException {
        int year = 2024;

        //задаем номера начальной и конечной недели по номеру в году
        int weekNumberFrom = 4;
        int weekNumberTo = 4;

        //настраиваем weekFields по дефолтной локали (страны)
        WeekFields weekFields = WeekFields.of(Locale.getDefault());

        //запускаем генерацию отчетов для каждой недели
        for (int i = weekNumberFrom; i <= weekNumberTo; i++) {
            LocalDate dateFromLocalDate = LocalDate.of(year, 1, 1).with(weekFields.weekOfYear(), i)
                    .with(weekFields.dayOfWeek(), 1);
            LocalDate dateToLocalDate = LocalDate.of(year, 1, 1).with(weekFields.weekOfYear(), i)
                    .with(weekFields.dayOfWeek(), 7);

            String dateFrom = dateFromLocalDate.toString();
            String dateTo = dateToLocalDate.toString();

            System.out.println("Будет сгенерирован отчет за даты " + dateFrom + " - " + dateTo);

            mainShadow(dateFrom, dateTo);

            //засыпаем на 1 минуту, так как отчет по хранению можно выгружать раз в минуту
            if(i < weekNumberTo) {
                //System.out.println("Засыпаем на 60 секунд");
                //Thread.sleep(60000);
            }
        }
    }


    public static void mainShadow(String dateFrom, String dateTo) throws IOException, InterruptedException, ParseException {

        //Атрашкевич
        //String directoryMain = "e:\\Proga\\AnaliticsWBAtrashkevich\\";
        //String token = "eyJhbGciOiJFUzI1NiIsImtpZCI6IjIwMjQwOTA0djEiLCJ0eXAiOiJKV1QifQ.eyJlbnQiOjEsImV4cCI6MTc0Mjk1Njc2NywiaWQiOiIwMTkyMjQ3OS0xMWFlLTdjZjQtODg1OS0zZmQ1NzBiYzE1MGEiLCJpaWQiOjM2NTAyMTI3LCJvaWQiOjY3Nzk4LCJzIjoxMDAsInNpZCI6IjhjZTkwMGI3LTJhMzQtNTQwMS04OTcwLWIyN2M1M2ZkZGM4ZCIsInQiOmZhbHNlLCJ1aWQiOjM2NTAyMTI3fQ.PjaAwloqWncl2ohvwMWF808r2Dq8IArhbdRvgLMClKgic4fQw4bDaApdST24SleW1XywXqdXUUmUbpdBm9A5Ow";
        //Седов
        String directoryMain = "e:\\Proga\\AnaliticsWBSedov\\";
        String token = "eyJhbGciOiJFUzI1NiIsImtpZCI6IjIwMjQxMDAxdjEiLCJ0eXAiOiJKV1QifQ.eyJlbnQiOjEsImV4cCI6MTc0NDQ5Njg4MSwiaWQiOiIwMTkyODA0NS01YTYwLTc3MzAtYjg4Ny0yMmQ0MjZkMGUwM2IiLCJpaWQiOjIzMjg3NTIyLCJvaWQiOjEzODU1MDIsInMiOjEwMCwic2lkIjoiNWZmNThlYTgtMGU4Ny00NjkwLWExNDQtN2NiYWVkOWQwMmUyIiwidCI6ZmFsc2UsInVpZCI6MjMyODc1MjJ9.QyJLsJioIZxfSKx7hKAjKljWIBFW7o6HzX-syw0UoXWdBm581vsVkkv0kkrdZefa6RdQobLLFBqLf9UkkuUqTQ";

        String directoryIntroductories = directoryMain + "Вводные\\";
        String directoryReports = directoryMain + "Отчеты" + dateFrom + "_" + dateTo + "\\";

        //создаем папку и загружаем по Api необходимые таблицы
        //Files.createDirectories(Path.of(directoryReports));
        //DetalizationApiDownload.downloadDetalization(directoryReports, dateFrom, dateTo, token);
        //AdsApiDownload.downloadApiAds(directoryReports, dateTo, token);
        //StorageApiDownload.downloadStorage(directoryReports, dateFrom, dateTo, token);
        //SalesApiDownload.downloadSales(directoryReports, dateFrom, dateTo, token);

        //открываем потоки для чтения скаченных таблиц и 1 поток для записи таблицы с итоговым отчетом
        try (FileOutputStream fileOutputStreamUnitEconomic = new FileOutputStream(directoryReports + "UnitЭкономика.xlsx");
             FileInputStream fileInputStreamSebestoimost = new FileInputStream(directoryIntroductories + "ВводныеСебестоимость.xlsx");
             FileInputStream fileInputStreamNazvaniyaRK = new FileInputStream(directoryIntroductories + "ВводныеНазванияРК.xlsx");
             FileInputStream fileInputStreamProchieUderzhaniya = new FileInputStream(directoryReports + "Реклама.xlsx");
             FileInputStream fileInputStreamReportFromWB = new FileInputStream(directoryReports + "Детализация.xlsx");
             FileInputStream fileInputStreamHranenie = new FileInputStream(directoryReports + "Хранение.xlsx");
             FileInputStream fileInputStreamProdazhi = new FileInputStream(directoryReports + "Продажи.xlsx")) {

            //создаем новый WorkBook и Sheet для конечного отчета
            XSSFWorkbook workbookReportConverted = new XSSFWorkbook();
            sheetReportConverted = workbookReportConverted.createSheet();

            //инициализируем Excel файлы
            XSSFWorkbook workbookNazvaniyaRK = new XSSFWorkbook(fileInputStreamNazvaniyaRK);
            XSSFWorkbook workbookSebestoimost = new XSSFWorkbook(fileInputStreamSebestoimost);
            XSSFWorkbook workbookHranenie = new XSSFWorkbook(fileInputStreamHranenie);
            XSSFWorkbook workbookReportFromWB = new XSSFWorkbook(fileInputStreamReportFromWB);
            XSSFWorkbook workbookProchieUderzhaniya = new XSSFWorkbook(fileInputStreamProchieUderzhaniya);
            XSSFWorkbook workbookProdazhi = new XSSFWorkbook(fileInputStreamProdazhi);


            //инициализируем листы в Excel файлах
            sheetSebestoimost = workbookSebestoimost.getSheetAt(0);
            sheetNazvaniyaRK = workbookNazvaniyaRK.getSheetAt(0);
            XSSFSheet sheetProchieUderzhaniya = workbookProchieUderzhaniya.getSheetAt(0);
            XSSFSheet sheetReportFromWB = workbookReportFromWB.getSheetAt(0);
            XSSFSheet sheetHranenie = workbookHranenie.getSheetAt(0);
            XSSFSheet sheetProdazhi = workbookProdazhi.getSheetAt(0);

            //инициализируем даты
            DatesInitializer.initDates(directoryReports, sheetReportConverted);

            //инициализируем названия столбцов
            ColumnNamesInitializer.initAllColumnNames(sheetReportConverted);

            //создаем списоки названий столбцов из всех файлов
            ArrayList<String> columnNamesDetalization = ColumnNamesCreator.createColumnNames(sheetReportFromWB);
            ArrayList<String> columnNamesProchieUderzhaniya = ColumnNamesCreator.createColumnNames(sheetProchieUderzhaniya);
            ArrayList<String> columnNamesStorage = ColumnNamesCreator.createColumnNames(sheetHranenie);
            ArrayList<String> columnNamesSales = ColumnNamesCreator.createColumnNames(sheetProdazhi);

            //инициализируем названия строк в Converted и получаем их список. Делаем очистку, чтобы на новом круге был пустой
            if(rowNamesConverted != null)
                rowNamesConverted.clear();
            rowNamesConverted = RowNamesConvertedInitializer.initRowNamesConverted(columnNamesDetalization, columnNamesStorage, sheetReportFromWB, sheetReportConverted, sheetHranenie, sheetSebestoimost);

            //создаем пустую матрицу необходимой длины и ширины
            double[][] matrix = MatrixCreator.createMatrix(sheetReportConverted);

            //инициализируем матрицу
            //основные показатели
            Part1Initializer.initialize1Part(matrix, rowNamesConverted, columnNamesDetalization, sheetReportFromWB, sheetSebestoimost, sheetReportConverted);
            //хранение
            Part2Initializer.initialize2Part(matrix, rowNamesConverted, columnNamesDetalization, columnNamesStorage, sheetHranenie, sheetReportFromWB, sheetSebestoimost);
            //прочие удержания
            Part3Initializer.initialize3Part(matrix, rowNamesConverted, columnNamesDetalization, columnNamesProchieUderzhaniya, sheetReportFromWB, sheetProchieUderzhaniya, sheetNazvaniyaRK, directoryIntroductories);
            //себестоимость и фулфиллмент
            Part4Initializer.initialize4Part(matrix, rowNamesConverted, sheetSebestoimost);
            //штрафы и платная приемка
            Part5Initializer.initialize5Part(matrix, columnNamesDetalization, sheetReportFromWB);
            //налог 7%
            Part6Initializer.initialize6Part(matrix, rowNamesConverted, columnNamesSales, sheetProdazhi, sheetSebestoimost);
            //вычисляемые показатели из существующих
            Part7Initializer.initialize7Part(matrix, rowNamesConverted);

            //переписываем данные из матрицы в таблицу
            MatrixRewriter.rewriteMatrix(matrix, sheetReportConverted);
            Colorizer.colorize(rowNamesConverted, sheetReportConverted, workbookReportConverted);

            workbookReportConverted.write(fileOutputStreamUnitEconomic);
            System.out.println("Файл создан");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Collections;

public class ExcelConverter {
    public static int counter = 0;

    public static double sum;

    public static CellStyle styleGreen;

    public static final String directoryIntroductories = "e:\\Proga\\AnaliticsWB\\Вводные\\";

    public static XSSFSheet sheetReportConverted;
    public static XSSFSheet sheetNazvaniyaRK;
    public static XSSFSheet sheetReportFromWB;

    public static void main(String[] args) {
        mainShadow("e:\\Proga\\AnaliticsWB\\Отчеты0807_1407\\");
    }

    public static void mainShadow(String directoryReports) {
        try (FileOutputStream fileOutputStream = new FileOutputStream(directoryReports + "UnitЭкономика.xlsx");
             FileInputStream fileInputStreamReportFromWB = new FileInputStream(directoryReports + "ОтчетДетализация.xlsx");
             FileInputStream fileInputStreamSebestoimost = new FileInputStream(directoryIntroductories + "ВводныеСебестоимость.xlsx");
             FileInputStream fileInputStreamHranenie = new FileInputStream(directoryReports + "ОтчетХранение.xlsx");
             FileInputStream fileInputStreamProchieUderzhaniya = new FileInputStream(directoryReports + "ОтчетРеклама.xlsx");
             FileInputStream fileInputStreamNazvaniyaRK = new FileInputStream(directoryIntroductories + "ВводныеНазванияРК.xlsx")) {

            XSSFWorkbook reportConverted = new XSSFWorkbook();
            sheetReportConverted = reportConverted.createSheet();

            //инициализируем Excel файлы
            XSSFWorkbook reportFromWB = new XSSFWorkbook(fileInputStreamReportFromWB);
            XSSFWorkbook sebestoimost = new XSSFWorkbook(fileInputStreamSebestoimost);
            XSSFWorkbook workbookHranenie = new XSSFWorkbook(fileInputStreamHranenie);
            XSSFWorkbook workbookProchieUderzhaniya = new XSSFWorkbook(fileInputStreamProchieUderzhaniya);
            XSSFWorkbook workbookNazvaniyaRK = new XSSFWorkbook(fileInputStreamNazvaniyaRK);

            //инициализируем листы в Excel файлах
            sheetReportFromWB = reportFromWB.getSheetAt(0);
            XSSFSheet sheetSebestoimost = sebestoimost.getSheetAt(0);
            XSSFSheet sheetHranenie = workbookHranenie.getSheetAt(0);
            XSSFSheet sheetProchieUderzhaniya = workbookProchieUderzhaniya.getSheetAt(0);
            sheetNazvaniyaRK = workbookNazvaniyaRK.getSheetAt(0);

            //создание зеленого стиля для заливки ячейки
            styleGreen = reportConverted.createCellStyle();
            styleGreen.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
            styleGreen.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            //инициализируем 1-ый столбец
            initFirstColumn("Товар");
            initFirstColumn("Выручка до вычета комиссии");
            initFirstColumn("Продажа");
            initFirstColumn("Продажа штук");
            initFirstColumn("Возврат");
            initFirstColumn("Возврат штук");
            initFirstColumn("Выручка после вычета комиссии");
            initFirstColumn("Продажа");
            initFirstColumn("Возврат");
            initFirstColumn("Коррекция продаж");
            initFirstColumn("Получены деньги штук");
            initFirstColumn("Выручка до вычета комиссии на 1 шт.");
            initFirstColumn("Выручка после вычета комиссии на 1 шт.");
            initFirstColumn("Логистика");
            initFirstColumn("Доставка к клиенту");
            initFirstColumn("Доставка к клиенту штук");
            initFirstColumn("Доставка от клиента");
            initFirstColumn("Доставка от клиента штук");
            initFirstColumn("Коррекция логистики");
            initFirstColumn("Штрафы");
            initFirstColumn("Хранение");
            initFirstColumn("Платная приемка");
            initFirstColumn("Прочие удержания");
            initFirstColumn("Приход на счет вб");
            initFirstColumn("Себестоимость");
            initFirstColumn("Фулфилмент");
            initFirstColumn("Налог 7%");
            initFirstColumn("Прибыль на маркетинг");
            initFirstColumn("Прибыль на маркетинг на 1 штуку");

            //расставляем названия категорий
            ArrayList<String> list = new ArrayList<>();
            for (Row row : sheetReportFromWB) {
                Cell cell = row.getCell(Alphabet.C.ordinal());
                if (cell != null) {
                    String category = cell.getStringCellValue();
                    if (!list.contains(category) && !category.equals("Предмет") && !category.isEmpty()) {
                        list.add(category);
                    }
                }
            }
            Collections.sort(list);
            for (int i = 0; i < list.size(); i++) {
                sheetReportConverted.getRow(0).createCell(Alphabet.C.ordinal() + i).setCellValue(list.get(i));
            }

            //название столбца Общий
            sheetReportConverted.getRow(0).createCell(1).setCellValue("Общий");

            //продажа до вычета комиссии
            double prodazhaDoVichKomiss = initOneCondition(Alphabet.K.ordinal(), "Продажа", Alphabet.P.ordinal(), 2, 1, sheetReportFromWB);
            ArrayList<Double> listProdazhaDoVichKomiss = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.K.ordinal(), "Продажа", Alphabet.C.ordinal(), list.get(i), Alphabet.P.ordinal(), 2, i + 2);
                listProdazhaDoVichKomiss.add(sum);
            }

            //продажа штук
            double prodazhaShtuk = initOneCondition(Alphabet.J.ordinal(), "Продажа", Alphabet.N.ordinal(), 3, 1, sheetReportFromWB);
            ArrayList<Double> listProdazhaShtuk = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.J.ordinal(), "Продажа", Alphabet.C.ordinal(), list.get(i), Alphabet.N.ordinal(), 3, i + 2);
                listProdazhaShtuk.add(sum);
            }

            //возврат до вычета комиссии
            double vozvratDoVichKomiss = initOneCondition(Alphabet.K.ordinal(), "Возврат", Alphabet.P.ordinal(), 4, 1, sheetReportFromWB);
            ArrayList<Double> listVozvratDoVichKomiss = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.K.ordinal(), "Возврат", Alphabet.C.ordinal(), list.get(i), Alphabet.P.ordinal(), 4, i + 2);
                listVozvratDoVichKomiss.add(sum);
            }

            //возврат штук
            double vozvratShtuk = initOneCondition(Alphabet.K.ordinal(), "Возврат", Alphabet.N.ordinal(), 5, 1, sheetReportFromWB);
            ArrayList<Double> listVozvratShtuk = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.K.ordinal(), "Возврат", Alphabet.C.ordinal(), list.get(i), Alphabet.N.ordinal(), 5, i + 2);
                listVozvratShtuk.add(sum);
            }

            //выручка до вычета комиссии
            double viruchkaDoVichetaKomissii = prodazhaDoVichKomiss - vozvratDoVichKomiss;
            sheetReportConverted.getRow(1).createCell(1).setCellValue(viruchkaDoVichetaKomissii);
            ArrayList<Double> listViruchkaDoVichetaKomissii = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = listProdazhaDoVichKomiss.get(i) - listVozvratDoVichKomiss.get(i);
                sheetReportConverted.getRow(1).createCell(i + 2).setCellValue(sum);
                listViruchkaDoVichetaKomissii.add(sum);
            }

            //продажа
            double prodazhaPosleVichKomiss = initOneCondition(Alphabet.J.ordinal(), "Продажа", Alphabet.AG.ordinal(), 7, 1, sheetReportFromWB);
            ArrayList<Double> listProdazhaPosleVichKomiss = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.J.ordinal(), "Продажа", Alphabet.C.ordinal(), list.get(i), Alphabet.AG.ordinal(), 7, i + 2);
                listProdazhaPosleVichKomiss.add(sum);
            }

            //возврат
            double vozvratPosleVychKomiss = initOneCondition(Alphabet.K.ordinal(), "Возврат", Alphabet.AG.ordinal(), 8, 1, sheetReportFromWB);
            ArrayList<Double> listVozvratPosleVichKomiss = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.K.ordinal(), "Возврат", Alphabet.C.ordinal(), list.get(i), Alphabet.AG.ordinal(), 8, i + 2);
                listVozvratPosleVichKomiss.add(sum);
            }

            //коррекция продаж
            double korrekciyaProdazh = initOneCondition(Alphabet.K.ordinal(), "Коррекция продаж", Alphabet.AG.ordinal(), 9, 1, sheetReportFromWB);
            ArrayList<Double> listKorrekciyaProdazh = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.K.ordinal(), "Коррекция продаж", Alphabet.C.ordinal(), list.get(i), Alphabet.AG.ordinal(), 9, i + 2);
                listKorrekciyaProdazh.add(sum);
            }

            //выручка после вычета комисии
            double viruchkaPosleVichetaKomissii = prodazhaPosleVichKomiss - vozvratPosleVychKomiss - korrekciyaProdazh;
            sheetReportConverted.getRow(6).createCell(1).setCellValue(viruchkaPosleVichetaKomissii);
            ArrayList<Double> listViruchkaPosleVichetaKomissii = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = listProdazhaPosleVichKomiss.get(i) - listVozvratPosleVichKomiss.get(i) - listKorrekciyaProdazh.get(i);
                sheetReportConverted.getRow(6).createCell(i + 2).setCellValue(sum);
                listViruchkaPosleVichetaKomissii.add(sum);
            }

            //получены деньги штук
            double poluchenyDengiShtuk = prodazhaShtuk - vozvratShtuk;
            sheetReportConverted.getRow(10).createCell(1).setCellValue(poluchenyDengiShtuk);
            ArrayList<Double> listPolucheniDengiShtuk = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = listProdazhaShtuk.get(i) - listVozvratShtuk.get(i);
                sheetReportConverted.getRow(10).createCell(i + 2).setCellValue(sum);
                listPolucheniDengiShtuk.add(sum);
            }

            //доставка к клиенту
            double dostavkaKKlientu = initOneCondition(Alphabet.AH.ordinal(), "1", Alphabet.AJ.ordinal(), 14, 1, sheetReportFromWB);
            ArrayList<Double> listDostavkaKklientu = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.AH.ordinal(), "1", Alphabet.C.ordinal(), list.get(i), Alphabet.AJ.ordinal(), 14, i + 2);
                listDostavkaKklientu.add(sum);
            }

            //доставка к клиенту штук
            initSumInColumn(Alphabet.AH.ordinal(), 15,Alphabet.B.ordinal(), sheetReportFromWB);
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.AH.ordinal(), "1", Alphabet.C.ordinal(), list.get(i), Alphabet.AH.ordinal(), 15, i + 2);
            }

            //доставка от клиента
            double dostavkaOtKlienta = initOneCondition(Alphabet.AI.ordinal(), "1", Alphabet.AJ.ordinal(), 16, 1, sheetReportFromWB);
            ArrayList<Double> listDostavkaOtklienta = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.AI.ordinal(), "1", Alphabet.C.ordinal(), list.get(i), Alphabet.AJ.ordinal(), 16, i + 2);
                listDostavkaOtklienta.add(sum);
            }

            //доставка от клиента штук
            initSumInColumn(Alphabet.AI.ordinal(), 17,Alphabet.B.ordinal(), sheetReportFromWB);
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.AI.ordinal(), "1", Alphabet.C.ordinal(), list.get(i), Alphabet.AI.ordinal(), 17, i + 2);
            }

            //коррекция логистики
            double korekciyaLogistiki = initOneCondition(Alphabet.K.ordinal(), "Коррекция логистики", Alphabet.AJ.ordinal(), 18, 1, sheetReportFromWB);
            ArrayList<Double> listKorekciyaLogistiki = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.K.ordinal(), "Коррекция логистики", Alphabet.C.ordinal(), list.get(i), Alphabet.AJ.ordinal(), 18, i + 2);
                listKorekciyaLogistiki.add(sum);
            }




            //хранение
            double hranenie = initSumInColumn(Alphabet.BD.ordinal(), 20, 1, sheetReportFromWB);
            ArrayList<Double> listHranenie = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initOneCondition(Alphabet.N.ordinal(), list.get(i), Alphabet.T.ordinal(), 20, i + 2, sheetHranenie);
                listHranenie.add(sum);
            }

            //получаем список категорий, по которым были затраты по хранению для проверки, что их столько же сколько list.size()
            ArrayList<String> listHranenieCategories = new ArrayList<>();
            for (Row row: sheetHranenie) {
                String categoryName = row.getCell(Alphabet.N.ordinal()).getStringCellValue();
                if (!categoryName.equals("Предмет") && !listHranenieCategories.contains(categoryName)) {
                    listHranenieCategories.add(categoryName);
                }
            }
            //проверяем есть ли недостающие столбцы и добавляем загаловки столбцов с недостающими категориями
            counter = 0;
            for (String hranenieCategory: listHranenieCategories) {
                sum = 0;
                if (!list.contains(hranenieCategory)) {
                    sheetReportConverted.getRow(0).createCell(2 + list.size() + counter).setCellValue(hranenieCategory);
                    sum = initOneCondition(Alphabet.N.ordinal(), hranenieCategory, Alphabet.T.ordinal(), 20,2 + list.size() + counter, sheetHranenie);
                    for (int j = 1; j <= 19; j++) {
                        sheetReportConverted.getRow(j).createCell(2 + list.size() + counter).setCellValue("-");
                    }
                    sheetReportConverted.getRow(22).createCell(2 + list.size() + counter).setCellValue("-");
                    sheetReportConverted.getRow(23).createCell(2 + list.size() + counter).setCellValue(-sum);
                    sheetReportConverted.getRow(24).createCell(2 + list.size() + counter).setCellValue("-");
                    sheetReportConverted.getRow(25).createCell(2 + list.size() + counter).setCellValue("-");
                    sheetReportConverted.getRow(26).createCell(2 + list.size() + counter).setCellValue(-sum * 0.07);
                    sheetReportConverted.getRow(27).createCell(2 + list.size() + counter).setCellValue(-sum * 0.93);
                    sheetReportConverted.getRow(28).createCell(2 + list.size() + counter).setCellValue("-");
                    counter++;
                    listHranenie.add(sum);
                }
            }
            //проверяем, что сумма хранения по категориям сходится с общим хранением
            double sumListHranenie = 0;
            for (Double hranenieFromCategory : listHranenie) {
                sumListHranenie += hranenieFromCategory;
            }
            long hranenieInt = Math.round(hranenie);
            long sumListHranenieInt = Math.round(sumListHranenie);
            if(hranenieInt - sumListHranenieInt > 1) {
                System.out.println("Общая сумма хранения не сходится с суммой хранения по категориям");
            }


            //штрафы
            double shtrafi = initOneCondition(Alphabet.K.ordinal(), "Штраф", Alphabet.AK.ordinal(), 19, 1, sheetReportFromWB);
            for (int i = 0; i < list.size(); i++) {
                sheetReportConverted.getRow(19).createCell(Alphabet.C.ordinal() + i).setCellValue("-");
            }


            //платная приемка
            double paidAcceptance = initSumInColumn(Alphabet.BF.ordinal(), 21, Alphabet.B.ordinal(), sheetReportFromWB);
            for (int i = 0; i < list.size() + counter; i++) {
                sheetReportConverted.getRow(21).createCell(2 + i).setCellValue("-");
            }

            //если штрафы > 0 или платная приемка > 0, то создаем столбец Неизвестно
            if(shtrafi > 0 || paidAcceptance > 0) {
                sheetReportConverted.getRow(19).createCell(2 + list.size() + counter).setCellValue(shtrafi);
                sheetReportConverted.getRow(21).createCell(2 + list.size() + counter).setCellValue(paidAcceptance);
                sheetReportConverted.getRow(23).createCell(2 + list.size() + counter).setCellValue(-shtrafi -paidAcceptance);
                sheetReportConverted.getRow(26).createCell(2 + list.size() + counter).setCellValue((-shtrafi - paidAcceptance) * 0.07);
                sheetReportConverted.getRow(27).createCell(2 + list.size() + counter).setCellValue((-shtrafi - paidAcceptance) * 0.93);
                sheetReportConverted.getRow(0).createCell(2 + list.size() + counter).setCellValue("Неизвестно");
                for (int i = 1; i <= 18; i++) {
                    sheetReportConverted.getRow(i).createCell(2 + list.size() + counter).setCellValue("-");
                }
                sheetReportConverted.getRow(20).createCell(2 + list.size() + counter).setCellValue("-");
                sheetReportConverted.getRow(22).createCell(2 + list.size() + counter).setCellValue("-");
                sheetReportConverted.getRow(24).createCell(2 + list.size() + counter).setCellValue("-");
                sheetReportConverted.getRow(25).createCell(2 + list.size() + counter).setCellValue("-");
                sheetReportConverted.getRow(28).createCell(2 + list.size() + counter).setCellValue("-");
            }


            //прочие удержания
            double prochieUderjania = initSumInColumn(Alphabet.BE.ordinal(), 22, 1, sheetReportFromWB);


            ArrayList<Integer> listActualDocuments = getListActualDocumentsRK(sheetReportFromWB, sheetProchieUderzhaniya);

            //расставляем прочие удержания
            ArrayList<Double> listProchieUderzhaniya = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                ArrayList<String> listRKNames = getListRKNames(list.get(i));
                sum = initTwoConditionReklama(Alphabet.B.ordinal(), listRKNames, Alphabet.G.ordinal(), listActualDocuments, Alphabet.F.ordinal(), 22, i + 2, sheetProchieUderzhaniya);
                listProchieUderzhaniya.add(sum);
            }


            //выручка до вычета комиссии на 1 шт.
            sheetReportConverted.getRow(11).createCell(1).setCellValue(viruchkaDoVichetaKomissii / poluchenyDengiShtuk);
            for (int i = 0; i < list.size(); i++) {
                if (listPolucheniDengiShtuk.get(i) != 0)
                    sum = listViruchkaDoVichetaKomissii.get(i) / listPolucheniDengiShtuk.get(i);
                else
                    sum = 0;
                sheetReportConverted.getRow(11).createCell(i + 2).setCellValue(sum);
            }

            //выручка после вычета комиссии на 1 шт.
            sheetReportConverted.getRow(12).createCell(1).setCellValue(viruchkaPosleVichetaKomissii / poluchenyDengiShtuk);
            for (int i = 0; i < list.size(); i++) {
                if (listPolucheniDengiShtuk.get(i) != 0)
                    sum = listViruchkaPosleVichetaKomissii.get(i) / listPolucheniDengiShtuk.get(i);
                else sum = 0;
                sheetReportConverted.getRow(12).createCell(i + 2).setCellValue(sum);
            }

            //логистика
            double logistika = dostavkaKKlientu + dostavkaOtKlienta + korekciyaLogistiki;
            sheetReportConverted.getRow(13).createCell(1).setCellValue(logistika);
            ArrayList<Double> listLogistika = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = listDostavkaKklientu.get(i) + listDostavkaOtklienta.get(i) + listKorekciyaLogistiki.get(i);
                sheetReportConverted.getRow(13).createCell(i + 2).setCellValue(sum);
                listLogistika.add(sum);
            }

            //приход на счет WB
            double prihodNaSchetWB = viruchkaPosleVichetaKomissii - logistika - shtrafi - hranenie - paidAcceptance - prochieUderjania;
            sheetReportConverted.getRow(23).createCell(1).setCellValue(prihodNaSchetWB);
            ArrayList<Double> listPrihodNaSchetWb = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = 0;
                sum = listViruchkaPosleVichetaKomissii.get(i) - listLogistika.get(i) - listHranenie.get(i) - listProchieUderzhaniya.get(i);
                sheetReportConverted.getRow(23).createCell(i + 2).setCellValue(sum);
                listPrihodNaSchetWb.add(sum);
            }

            //налог 7%
            double nalog7Proc = prihodNaSchetWB * 0.07;
            sheetReportConverted.getRow(26).createCell(1).setCellValue(nalog7Proc);
            ArrayList<Double> listNalog7Proc = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = 0;
                sum = listPrihodNaSchetWb.get(i) * 0.07;
                sheetReportConverted.getRow(26).createCell(i + 2).setCellValue(sum);
                listNalog7Proc.add(sum);
            }

            //Себестоимость
            double sebes = 0;
            boolean isSebesoimostExistInVvodnyeSebestoimost = false;
            ArrayList<String> articlesWithoutSebestoimost = new ArrayList<>();
            counter = 0;
            for (Row row : sheetReportFromWB) {
                counter++;
                Cell cellToCheck = row.getCell(Alphabet.J.ordinal());
                if (cellToCheck != null) {
                    String serviceType = cellToCheck.getStringCellValue();
                    //если продажа, то добавляем себестоимость
                    if (serviceType.equals("Продажа")) {
                        String sellersArticle = row.getCell(Alphabet.F.ordinal()).getStringCellValue();
                        for (Row rowSebes : sheetSebestoimost) {
                            if (rowSebes.getCell(Alphabet.B.ordinal()).getStringCellValue().equals(sellersArticle)) {
                                Cell cell = rowSebes.getCell(Alphabet.C.ordinal());
                                if (cell == null) {
                                    System.out.println("Вы где-то не указали себестоимость товара, но артикул продавца есть.");
                                } else {
                                    sebes += cell.getNumericCellValue();
                                }
                                isSebesoimostExistInVvodnyeSebestoimost = true;
                            }
                        }
                        if (!isSebesoimostExistInVvodnyeSebestoimost) {
                            articlesWithoutSebestoimost.add(sellersArticle);
                        }
                        isSebesoimostExistInVvodnyeSebestoimost = false;
                    }
                    //если возврат, то отнимаем себестоимость
                    if (serviceType.equals("Возврат")) {
                        String sellersArticle = row.getCell(Alphabet.F.ordinal()).getStringCellValue();
                        for (Row rowSebes : sheetSebestoimost) {
                            if (rowSebes.getCell(Alphabet.B.ordinal()).getStringCellValue().equals(sellersArticle)) {
                                Cell cell = rowSebes.getCell(Alphabet.C.ordinal());
                                if (cell == null) {
                                    System.out.println("Вы где-то не указали себестоимость товара, но артикул продавца есть.");
                                } else {
                                    sebes -= cell.getNumericCellValue();
                                }
                                isSebesoimostExistInVvodnyeSebestoimost = true;
                            }
                        }
                        if (!isSebesoimostExistInVvodnyeSebestoimost) {
                            articlesWithoutSebestoimost.add(sellersArticle);
                        }
                        isSebesoimostExistInVvodnyeSebestoimost = false;
                    }
                }
            }
            for (String article : articlesWithoutSebestoimost) {
                System.out.println("Нет данных о себестоимости товара с артикулом продавца или проверьте заглавные и строчные буквы в написании артикула продавца " + article);
            }
            sheetReportConverted.getRow(24).createCell(1).setCellValue(sebes);

            ArrayList<Double> listSebestoimost = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = 0;
                for (Row row : sheetReportFromWB) {
                    Cell cellToCheck = row.getCell(Alphabet.J.ordinal());
                    if (cellToCheck != null) {
                        String serviceType = cellToCheck.getStringCellValue();
                        //если продажа, то добавляем себестоимость
                        if (serviceType.equals("Продажа")) {
                            String category = row.getCell(Alphabet.C.ordinal()).getStringCellValue();
                            if (category.equals(list.get(i))) {
                                String sellersArticle = row.getCell(Alphabet.F.ordinal()).getStringCellValue();
                                for (Row rowSebes : sheetSebestoimost) {
                                    if (rowSebes.getCell(Alphabet.B.ordinal()).getStringCellValue().equals(sellersArticle)) {
                                        sum += rowSebes.getCell(Alphabet.C.ordinal()).getNumericCellValue();
                                    }
                                }
                            }
                        }
                        //если возврат, то отнимаем себестоимость
                        if (serviceType.equals("Возврат")) {
                            String category = row.getCell(Alphabet.C.ordinal()).getStringCellValue();
                            if (category.equals(list.get(i))) {
                                String sellersArticle = row.getCell(Alphabet.F.ordinal()).getStringCellValue();
                                for (Row rowSebes : sheetSebestoimost) {
                                    if (rowSebes.getCell(Alphabet.B.ordinal()).getStringCellValue().equals(sellersArticle)) {
                                        sum -= rowSebes.getCell(Alphabet.C.ordinal()).getNumericCellValue();
                                    }
                                }
                            }
                        }
                    }
                }
                sheetReportConverted.getRow(24).createCell(i + 2).setCellValue(sum);
                listSebestoimost.add(sum);
            }


            //Фулфилмент
            double fulfilmentSum = 0;
            for (Row rowReport : sheetReportFromWB) {
                Cell cellToCheck = rowReport.getCell(Alphabet.J.ordinal());
                if (cellToCheck != null) {
                    String stringCellValue = cellToCheck.getStringCellValue();
                    //если продажа, то прибавляем фулфилмент
                    if (stringCellValue.equals("Продажа")) {
                        String sellersArticle = rowReport.getCell(Alphabet.F.ordinal()).getStringCellValue();
                        for (Row rowSebes : sheetSebestoimost) {
                            if (rowSebes.getCell(Alphabet.B.ordinal()).getStringCellValue().equals(sellersArticle)) {
                                fulfilmentSum += rowSebes.getCell(Alphabet.D.ordinal()).getNumericCellValue();
                            }
                        }
                    }
                    //если возврат, то отнимаем фулфилмент
                    if (stringCellValue.equals("Возврат")) {
                        String sellersArticle = rowReport.getCell(Alphabet.F.ordinal()).getStringCellValue();
                        for (Row rowSebes : sheetSebestoimost) {
                            if (rowSebes.getCell(Alphabet.B.ordinal()).getStringCellValue().equals(sellersArticle)) {
                                fulfilmentSum -= rowSebes.getCell(Alphabet.D.ordinal()).getNumericCellValue();
                            }
                        }
                    }
                }
            }
            sheetReportConverted.getRow(25).createCell(1).setCellValue(fulfilmentSum);

            ArrayList<Double> listFulfilment = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = 0;
                for (Row row : sheetReportFromWB) {
                    Cell cellToCheck = row.getCell(Alphabet.J.ordinal());
                    if (cellToCheck != null) {
                        String serviceType = cellToCheck.getStringCellValue();
                        //если продажа, то прибавляем фулфилмент
                        if (serviceType.equals("Продажа")) {
                            String category = row.getCell(Alphabet.C.ordinal()).getStringCellValue();
                            if (category.equals(list.get(i))) {
                                String sellersArticle = row.getCell(Alphabet.F.ordinal()).getStringCellValue();
                                for (Row rowSebes : sheetSebestoimost) {
                                    if (rowSebes.getCell(Alphabet.B.ordinal()).getStringCellValue().equals(sellersArticle)) {
                                        sum += rowSebes.getCell(Alphabet.D.ordinal()).getNumericCellValue();
                                    }
                                }
                            }
                        }
                        //если возврат, то отнимаем фулфилмент
                        if (serviceType.equals("Возврат")) {
                            String category = row.getCell(Alphabet.C.ordinal()).getStringCellValue();
                            if (category.equals(list.get(i))) {
                                String sellersArticle = row.getCell(Alphabet.F.ordinal()).getStringCellValue();
                                for (Row rowSebes : sheetSebestoimost) {
                                    if (rowSebes.getCell(Alphabet.B.ordinal()).getStringCellValue().equals(sellersArticle)) {
                                        sum -= rowSebes.getCell(Alphabet.D.ordinal()).getNumericCellValue();
                                    }
                                }
                            }
                        }
                    }
                }
                sheetReportConverted.getRow(25).createCell(i + 2).setCellValue(sum);
                listFulfilment.add(sum);
            }

            //прибыль на маркетинг
            double pribilNaMarketing = prihodNaSchetWB - sebes - fulfilmentSum - nalog7Proc;
            sheetReportConverted.getRow(27).createCell(1).setCellValue(pribilNaMarketing);
            ArrayList<Double> listPribilNaMarketing = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = listPrihodNaSchetWb.get(i) - listSebestoimost.get(i) - listFulfilment.get(i) - listNalog7Proc.get(i);
                sheetReportConverted.getRow(27).createCell(i + 2).setCellValue(sum);
                listPribilNaMarketing.add(sum);
            }

            //прибыль на маркетинг на 1 штуку
            sheetReportConverted.getRow(28).createCell(1).setCellValue(pribilNaMarketing / poluchenyDengiShtuk);
            for (int i = 0; i < list.size(); i++) {
                if (listPolucheniDengiShtuk.get(i) != 0)
                    sum = listPribilNaMarketing.get(i) / listPolucheniDengiShtuk.get(i);
                else
                    sum = 0;
                sheetReportConverted.getRow(28).createCell(i + 2).setCellValue(sum);
            }


            fillRowGreen(1);
            fillRowGreen(6);
            fillRowGreen(10);
            fillRowGreen(23);
            fillRowGreen(27);

            reportConverted.write(fileOutputStream);
            System.out.println("Файл создан");
        } catch (Exception e) {
            System.out.println(e);
        }
    }

    public static ArrayList<Integer> getListActualDocumentsRK(XSSFSheet sheetReportFromWB, XSSFSheet sheetProchieUderzhaniya) {
        ArrayList<Integer> listActualDocumentsSums = new ArrayList<>();
        for (Row row : sheetReportFromWB) {
            if (row.getRowNum() != 0) {
                Cell cell = row.getCell(Alphabet.BE.ordinal());
                if (cell.getCellType() == CellType.STRING) {
                    break;
                }
                sum = 0;
                sum = cell.getNumericCellValue();
                if (sum != 0)
                    listActualDocumentsSums.add((int) sum);
            }
        }

        ArrayList<Integer> listAllDocuments = new ArrayList<>();
        for (Row row : sheetProchieUderzhaniya) {
            if (row.getRowNum() != 0) {
                Cell cell = row.getCell(Alphabet.G.ordinal());
                if (cell.getCellType() == CellType.STRING) {
                    break;
                }
                int docNumber = (int) cell.getNumericCellValue();
                if (!listAllDocuments.contains(docNumber)) {
                    listAllDocuments.add(docNumber);
                }
            }
        }

        ArrayList<Integer> listAllDocumentsSums = new ArrayList<>();
        for (Integer listAllDocument : listAllDocuments) {
            sum = 0;
            for (Row row : sheetProchieUderzhaniya) {
                if (row.getRowNum() != 0) {
                    Cell cellToCheck = row.getCell(Alphabet.G.ordinal());
                    int docNum = (int) cellToCheck.getNumericCellValue();
                    if (listAllDocument.equals(docNum)) {
                        Cell cellToSum = row.getCell(Alphabet.F.ordinal());
                        int rashod = (int) cellToSum.getNumericCellValue();
                        sum += rashod;
                    }
                }
            }
            listAllDocumentsSums.add((int) sum);
        }

        //получаем список с актуальными номерами документов
        ArrayList<Integer> listActualDocuments = new ArrayList<>();
        for (Integer listActualDocumentsSum : listActualDocumentsSums) {
            for (int j = 0; j < listAllDocumentsSums.size(); j++) {
                if (listActualDocumentsSum.equals(listAllDocumentsSums.get(j))) {
                    listActualDocuments.add(listAllDocuments.get(j));
                }
            }
        }

        //если определилось недостаточное количество actualDocuments, то выводим сообщение об ошибке
        if(listActualDocuments.size() < listActualDocumentsSums.size()) {
            System.out.println("Недостаточно данных по рекламе. Выберите более раннюю начальную дату. Проверьте, что конечная дата соответствует дате конца периода отчета или более поздняя.");
        }

        return listActualDocuments;
    }

    public static void fillRowGreen(int rowNum) {
        Row row = sheetReportConverted.getRow(rowNum);
        if (row != null) {
            int lastCellNum = row.getLastCellNum();
            for (int i = 0; i < lastCellNum; i++) {
                Cell cell = row.getCell(i);
                cell.setCellStyle(styleGreen);
            }
        }
    }


    public static void initFirstColumn(String str) {
        Row row = sheetReportConverted.createRow(counter);
        row.createCell(0).setCellValue(str);
        counter++;
    }

    public static double initOneCondition(int columnToCheck, String condition, int columnToSum, int rowToSet, int columnToSet, Sheet sheetFrom) {
        sum = 0;
        for (Row row : sheetFrom) {
            Cell cellToSum = row.getCell(columnToSum);
            Cell cellToChek = row.getCell(columnToCheck);
            if (cellToChek != null) {
                CellType cellToCheckType = cellToChek.getCellType();
                String stringCellToCheckValue;

                if (cellToCheckType == CellType.STRING) {
                    stringCellToCheckValue = cellToChek.getStringCellValue();
                } else {
                    int intCellToCheckValue = (int) cellToChek.getNumericCellValue();
                    stringCellToCheckValue = intCellToCheckValue + "";
                }

                if (stringCellToCheckValue.equals(condition) && cellToSum != null && cellToSum.getCellType() == CellType.NUMERIC) {
                    sum += cellToSum.getNumericCellValue();
                }

                if (stringCellToCheckValue.equals(condition) && cellToSum != null && cellToSum.getCellType() == CellType.STRING) {
                    String hranenieDot = cellToSum.getStringCellValue();
                    double hranenieDouble = Double.parseDouble(hranenieDot);
                    sum += hranenieDouble;
                }
            }
        }
        Row row = sheetReportConverted.getRow(rowToSet);
        row.createCell(columnToSet).setCellValue(sum);
        return sum;
    }

    public static double initTwoCondition(int columnToCheck1, String condition1, int columnToCheck2, String condition2, int columnToSum, int rowToSet, int columnToSet) {
        sum = 0;
        for (Row row : sheetReportFromWB) {
            Cell cellToSum = row.getCell(columnToSum);

            Cell cellToCheck1 = row.getCell(columnToCheck1);
            if (cellToCheck1 == null)
                continue;
            CellType cellToCheck1Type = cellToCheck1.getCellType();
            String stringCellToCheckValue1;
            if (cellToCheck1Type == CellType.STRING) {
                stringCellToCheckValue1 = cellToCheck1.getStringCellValue();
            } else {
                int intCellToCheckValue1 = (int) cellToCheck1.getNumericCellValue();
                stringCellToCheckValue1 = intCellToCheckValue1 + "";
            }

            Cell cellToCheck2 = row.getCell(columnToCheck2);
            if (cellToCheck2 == null)
                continue;
            CellType cellToCheck2Type = cellToCheck2.getCellType();
            String stringCellToCheckValue2;
            if (cellToCheck2Type == CellType.STRING) {
                stringCellToCheckValue2 = cellToCheck2.getStringCellValue();
            } else {
                int intCellToCheckValue2 = (int) cellToCheck2.getNumericCellValue();
                stringCellToCheckValue2 = intCellToCheckValue2 + "";
            }


            if (stringCellToCheckValue1.equals(condition1) && stringCellToCheckValue2.equals(condition2) && cellToSum != null && cellToSum.getCellType() == CellType.NUMERIC) {
                sum += cellToSum.getNumericCellValue();
            }
        }
        sheetReportConverted.getRow(rowToSet).createCell(columnToSet).setCellValue(sum);
        return sum;
    }

    public static double initTwoConditionReklama(int columnRKNames, ArrayList<String> listRKNames, int columnDocNum, ArrayList<Integer> listDocNum, int columnToSum, int rowToSet, int columnToSet, Sheet sheetFrom) {
        sum = 0;
        for (Row row : sheetFrom) {
            if (row.getRowNum() != 0) {
                Cell cellToSum = row.getCell(columnToSum);

                Cell cellToCheck1 = row.getCell(columnRKNames);
                if (cellToCheck1 == null)
                    continue;
                CellType cellToCheck1Type = cellToCheck1.getCellType();
                String stringCellToCheckValue1;
                if (cellToCheck1Type == CellType.STRING) {
                    stringCellToCheckValue1 = cellToCheck1.getStringCellValue();
                } else {
                    int intCellToCheckValue1 = (int) cellToCheck1.getNumericCellValue();
                    stringCellToCheckValue1 = intCellToCheckValue1 + "";
                }

                Cell cellToCheck2 = row.getCell(columnDocNum);
                if (cellToCheck2 == null)
                    continue;
                CellType cellToCheck2Type = cellToCheck2.getCellType();
                int intCellToCheckValue2;
                if (cellToCheck2Type == CellType.STRING) {
                    String stringCellToCheckValue2 = cellToCheck2.getStringCellValue();
                    intCellToCheckValue2 = Integer.parseInt(stringCellToCheckValue2);
                } else {
                    intCellToCheckValue2 = (int) cellToCheck2.getNumericCellValue();
                }


                if (listRKNames.contains(stringCellToCheckValue1) && listDocNum.contains(intCellToCheckValue2) && cellToSum != null && cellToSum.getCellType() == CellType.NUMERIC) {
                    sum += cellToSum.getNumericCellValue();
                }
            }
        }
        sheetReportConverted.getRow(rowToSet).createCell(columnToSet).setCellValue(sum);
        return sum;
    }


    public static double initSumInColumn(int columnToSum, int rowToSet, int columnToSet, XSSFSheet sheetFrom) {
        sum = 0;
        for (Row row : sheetFrom) {
            Cell cell = row.getCell(columnToSum);
            if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                sum += cell.getNumericCellValue();
            }
        }
        sheetReportConverted.getRow(rowToSet).createCell(columnToSet).setCellValue(sum);
        return sum;
    }

    public static ArrayList<String> getListRKNames(String category) {
        ArrayList<String> list = new ArrayList<>();
        for (Row row : sheetNazvaniyaRK) {
            Cell cellCategory = row.getCell(Alphabet.B.ordinal());
            String rowCategory = cellCategory.getStringCellValue();
            if (rowCategory.equals(category)) {
                Cell cellRKName = row.getCell(Alphabet.A.ordinal());
                String rkName = cellRKName.getStringCellValue();
                list.add(rkName);
            }
        }
        return list;
    }
}

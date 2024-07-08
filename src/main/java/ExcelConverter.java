import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Collections;

public class ExcelConverter {
    private static int counter = 0;

    private static double sum;

    private static CellStyle styleGreen;

    private static String directory = "e:\\Бизнес\\Черновики";

    public static void main(String[] args) {
        try {
            //отрываем потоки
            FileOutputStream fileOutputStream = new FileOutputStream(directory + "\\UnitЭкономика.xlsx");
            FileInputStream fileInputStreamReportFromWB = new FileInputStream(new File(directory + "\\ОтчетДетализация.xlsx"));
            FileInputStream fileInputStreamSebestoimost = new FileInputStream(new File(directory + "\\ВводныеСебестоимость.xlsx"));

            //инициализируем Excel файлы
            XSSFWorkbook reportConverted = new XSSFWorkbook();
            XSSFWorkbook reportFromWB = new XSSFWorkbook(fileInputStreamReportFromWB);
            XSSFWorkbook sebestoimost = new XSSFWorkbook(fileInputStreamSebestoimost);

            //инициализируем листы в Excel файлах
            XSSFSheet sheetReportConverted = reportConverted.createSheet();
            XSSFSheet sheetReportFromWB = reportFromWB.getSheetAt(0);
            XSSFSheet sheetSebestoimost = sebestoimost.getSheetAt(0);


            //создание зеленого стиля для заливки ячейки
            styleGreen = reportConverted.createCellStyle();
            styleGreen.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
            styleGreen.setFillPattern(FillPatternType.SOLID_FOREGROUND);


            //инициализируем 1-ый столбец
            //initFirstColumn("Дата с ... до ...", sheetReportConverted);
            initFirstColumn("Товар", sheetReportConverted);
            initFirstColumn("Выручка до вычета комиссии", sheetReportConverted);
            initFirstColumn("Продажа", sheetReportConverted);
            initFirstColumn("Продажа штук", sheetReportConverted);
            initFirstColumn("Возврат", sheetReportConverted);
            initFirstColumn("Возврат штук", sheetReportConverted);
            initFirstColumn("Выручка после вычета комиссии", sheetReportConverted);
            initFirstColumn("Продажа", sheetReportConverted);
            initFirstColumn("Возврат", sheetReportConverted);
            initFirstColumn("Коррекция продаж", sheetReportConverted);
            initFirstColumn("Получены деньги штук", sheetReportConverted);
            initFirstColumn("Выручка до вычета комиссии на 1 шт.", sheetReportConverted);
            initFirstColumn("Выручка после вычета комиссии на 1 шт.", sheetReportConverted);
            initFirstColumn("Логистика", sheetReportConverted);
            initFirstColumn("Доставка к клиенту", sheetReportConverted);
            initFirstColumn("Доставка к клиенту штук", sheetReportConverted);
            initFirstColumn("Доставка от клиента", sheetReportConverted);
            initFirstColumn("Доставка от клиента штук", sheetReportConverted);
            initFirstColumn("Коррекция логистики", sheetReportConverted);
            initFirstColumn("Штрафы", sheetReportConverted);
            initFirstColumn("Хранение", sheetReportConverted);
            initFirstColumn("Прочие удержания", sheetReportConverted);
            initFirstColumn("Приход на счет вб", sheetReportConverted);
            initFirstColumn("Себестоимость", sheetReportConverted);
            initFirstColumn("Фулфилмент", sheetReportConverted);
            initFirstColumn("Налог 7%", sheetReportConverted);
            initFirstColumn("Прибыль на маркетинг", sheetReportConverted);
            initFirstColumn("Прибыль на маркетинг на 1 штуку", sheetReportConverted);

            //расставляем названия категорий
            ArrayList<String> list = new ArrayList<>();
            for (Row row : sheetReportFromWB) {
                Cell cell = row.getCell(Alphabet.C.ordinal());
                if (cell != null) {
                    String category = cell.getStringCellValue();
                    if (!list.contains(category) && !category.equals("Предмет")) {
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
            double prodazhaDoVichKomiss = initOneCondition(Alphabet.K.ordinal(), "Продажа", Alphabet.P.ordinal(), 2, 1, sheetReportFromWB, sheetReportConverted);
            ArrayList<Double> listProdazhaDoVichKomiss = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.K.ordinal(), "Продажа", Alphabet.C.ordinal(), list.get(i), Alphabet.P.ordinal(), 2, i + 2, sheetReportFromWB, sheetReportConverted);
                listProdazhaDoVichKomiss.add(sum);
            }

            //продажа штук
            double prodazhaShtuk = initOneCondition(Alphabet.K.ordinal(), "Продажа", Alphabet.N.ordinal(), 3, 1, sheetReportFromWB, sheetReportConverted);
            ArrayList<Double> listProdazhaShtuk = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.K.ordinal(), "Продажа", Alphabet.C.ordinal(), list.get(i), Alphabet.N.ordinal(), 3, i + 2, sheetReportFromWB, sheetReportConverted);
                listProdazhaShtuk.add(sum);
            }

            //возврат до вычета комиссии
            double vozvratDoVichKomiss = initOneCondition(Alphabet.K.ordinal(), "Возврат", Alphabet.T.ordinal(), 4, 1, sheetReportFromWB, sheetReportConverted);
            ArrayList<Double> listVozvratDoVichKomiss = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.K.ordinal(), "Возврат", Alphabet.C.ordinal(), list.get(i), Alphabet.P.ordinal(), 4, i + 2, sheetReportFromWB, sheetReportConverted);
                listVozvratDoVichKomiss.add(sum);
            }

            //возврат штук
            double vozvratShtuk = initOneCondition(Alphabet.K.ordinal(), "Возврат", Alphabet.N.ordinal(), 5, 1, sheetReportFromWB, sheetReportConverted);
            ArrayList<Double> listVozvratShtuk = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.K.ordinal(), "Возврат", Alphabet.C.ordinal(), list.get(i), Alphabet.N.ordinal(), 5, i + 2, sheetReportFromWB, sheetReportConverted);
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
            double prodazhaPosleVichKomiss = initOneCondition(Alphabet.K.ordinal(), "Продажа", Alphabet.AG.ordinal(), 7, 1, sheetReportFromWB, sheetReportConverted);
            ArrayList<Double> listProdazhaPosleVichKomiss = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.K.ordinal(), "Продажа", Alphabet.C.ordinal(), list.get(i), Alphabet.AG.ordinal(), 7, i + 2, sheetReportFromWB, sheetReportConverted);
                listProdazhaPosleVichKomiss.add(sum);
            }

            //возврат
            double vozvratPosleVychKomiss = initOneCondition(Alphabet.K.ordinal(), "Возврат", Alphabet.AG.ordinal(), 8, 1, sheetReportFromWB, sheetReportConverted);
            ArrayList<Double> listVozvratPosleVichKomiss = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.K.ordinal(), "Возврат", Alphabet.C.ordinal(), list.get(i), Alphabet.AG.ordinal(), 8, i + 2, sheetReportFromWB, sheetReportConverted);
                listVozvratPosleVichKomiss.add(sum);
            }

            //коррекция продаж
            double korrekciyaProdazh = initOneCondition(Alphabet.K.ordinal(), "Коррекция продаж", Alphabet.AG.ordinal(), 9, 1, sheetReportFromWB, sheetReportConverted);
            ArrayList<Double> listKorrekciyaProdazh = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.K.ordinal(), "Коррекция продаж", Alphabet.C.ordinal(), list.get(i), Alphabet.AG.ordinal(), 9, i + 2, sheetReportFromWB, sheetReportConverted);
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
            double dostavkaKKlientu = initOneCondition(Alphabet.AH.ordinal(), "1", Alphabet.AJ.ordinal(), 14, 1, sheetReportFromWB, sheetReportConverted);
            ArrayList<Double> listDostavkaKklientu = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.AH.ordinal(), "1", Alphabet.C.ordinal(), list.get(i), Alphabet.AJ.ordinal(), 14, i + 2, sheetReportFromWB, sheetReportConverted);
                listDostavkaKklientu.add(sum);
            }

            //доставка к клиенту штук
            double dostavkaKKlientuShtuk = initOneCondition(Alphabet.AH.ordinal(), "1", Alphabet.AH.ordinal(), 15, 1, sheetReportFromWB, sheetReportConverted);
            ArrayList<Double> listDostavkaKklientuShtuk = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.AH.ordinal(), "1", Alphabet.C.ordinal(), list.get(i), Alphabet.AH.ordinal(), 15, i + 2, sheetReportFromWB, sheetReportConverted);
                listDostavkaKklientuShtuk.add(sum);
            }

            //доставка от клиента
            double dostavkaOtKlienta = initOneCondition(Alphabet.AI.ordinal(), "1", Alphabet.AJ.ordinal(), 16, 1, sheetReportFromWB, sheetReportConverted);
            ArrayList<Double> listDostavkaOtklienta = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.AI.ordinal(), "1", Alphabet.C.ordinal(), list.get(i), Alphabet.AJ.ordinal(), 16, i + 2, sheetReportFromWB, sheetReportConverted);
                listDostavkaOtklienta.add(sum);
            }

            //доставка от клиента штук
            double dostavkaOtKlientaShtuk = initOneCondition(Alphabet.AI.ordinal(), "1", Alphabet.AI.ordinal(), 17, 1, sheetReportFromWB, sheetReportConverted);
            ArrayList<Double> listDostavkaOtKlientaShtuk = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.AI.ordinal(), "1", Alphabet.C.ordinal(), list.get(i), Alphabet.AI.ordinal(), 17, i + 2, sheetReportFromWB, sheetReportConverted);
                listDostavkaOtKlientaShtuk.add(sum);
            }

            //коррекция логистики
            double korekciyaLogistiki = initOneCondition(Alphabet.K.ordinal(), "Коррекция логистики", Alphabet.AJ.ordinal(), 18, 1, sheetReportFromWB, sheetReportConverted);
            ArrayList<Double> listKorekciyaLogistiki = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initTwoCondition(Alphabet.K.ordinal(), "Коррекция логистики", Alphabet.C.ordinal(), list.get(i), Alphabet.AJ.ordinal(), 18, i + 2, sheetReportFromWB, sheetReportConverted);
                listKorekciyaLogistiki.add(sum);
            }

            //штрафы
            double shtrafi = initOneCondition(Alphabet.K.ordinal(), "Штраф", Alphabet.AK.ordinal(), 19, 1, sheetReportFromWB, sheetReportConverted);

            //хранение
            double hranenie = initSumInColumn(Alphabet.BD.ordinal(), 20, 1, sheetReportFromWB, sheetReportConverted);
            FileInputStream fileInputStreamHranenie = new FileInputStream(new File("e:\\Бизнес\\Черновики\\ОтчетХранение.xlsx"));
            XSSFWorkbook workbookHranenie = new XSSFWorkbook(fileInputStreamHranenie);
            XSSFSheet sheetHranenie = workbookHranenie.getSheetAt(0);
            ArrayList<Double> listHranenie = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = initOneCondition(Alphabet.N.ordinal(), list.get(i), Alphabet.T.ordinal(), 20, i + 2, sheetHranenie, sheetReportConverted);
                listHranenie.add(sum);
            }
            fileInputStreamHranenie.close();


            //прочие удержания
            double prochieUderjania = initSumInColumn(Alphabet.BE.ordinal(), 21, 1, sheetReportFromWB, sheetReportConverted);
            FileInputStream fileInputStreamProchieUderzhaniya = new FileInputStream(new File(directory + "\\ОтчетРеклама.xlsx"));
            FileInputStream fileInputStreamNazvaniyaRK = new FileInputStream(new File(directory + "\\ВводныеНазванияРК.xlsx"));

            XSSFWorkbook workbookProchieUderzhaniya = new XSSFWorkbook(fileInputStreamProchieUderzhaniya);
            XSSFWorkbook workbookNazvaniyaRK = new XSSFWorkbook(fileInputStreamNazvaniyaRK);

            XSSFSheet sheetProchieUderzhaniya = workbookProchieUderzhaniya.getSheetAt(0);
            XSSFSheet sheetNazvaniyaRK = workbookNazvaniyaRK.getSheetAt(0);

            ArrayList<Integer> listActualDocumentsSums = new ArrayList();
            int counter = 0;
            for (Row row: sheetReportFromWB) {
                if (row.getRowNum() != 0) {
                    Cell cell = row.getCell(Alphabet.BE.ordinal());
                    if(cell.getCellType() == CellType.STRING) {
                        break;
                    }
                    sum = 0;
                    sum = cell.getNumericCellValue();
                    if (sum != 0)
                        listActualDocumentsSums.add((int)sum);
                    counter++;
                }
            }

            ArrayList<Integer> listAllDocuments = new ArrayList();
            for(Row row: sheetProchieUderzhaniya) {
                if (row.getRowNum() != 0) {
                    Cell cell = row.getCell(Alphabet.G.ordinal());
                    if (cell.getCellType() == CellType.STRING) {
                        break;
                    }
                    int docNumber = (int)cell.getNumericCellValue();
                    if(!listAllDocuments.contains(docNumber)) {
                        listAllDocuments.add(docNumber);
                    }
                }
            }

            ArrayList<Integer> listAllDocumentsSums = new ArrayList();
            for (int i = 0; i < listAllDocuments.size(); i++) {
                sum = 0;
                for(Row row: sheetProchieUderzhaniya) {
                    if (row.getRowNum() != 0) {
                        Cell cellToCheck = row.getCell(Alphabet.G.ordinal());
                        int docNum = (int)cellToCheck.getNumericCellValue();
                        if(listAllDocuments.get(i).equals(docNum)) {
                            Cell cellToSum = row.getCell(Alphabet.F.ordinal());
                            int rashod = (int)cellToSum.getNumericCellValue();
                            sum += rashod;
                        }
                    }
                }
                listAllDocumentsSums.add((int)sum);
            }

            //получаем список с актуальными номерами документов
            ArrayList<Integer> listActualDocuments = new ArrayList();
            for (int i = 0; i < listActualDocumentsSums.size(); i++) {
                for (int j = 0; j < listAllDocumentsSums.size(); j++) {
                    if (listActualDocumentsSums.get(i).equals(listAllDocumentsSums.get(j))) {
                        listActualDocuments.add(listAllDocuments.get(j));
                    }
                }
            }
            //расставляем прочие удержания
            ArrayList<Double> listProchieUderzhaniya = new ArrayList();
            for (int i = 0; i < list.size(); i++) {
                ArrayList listRKNames = getListRKNamesForCategory(list.get(i), sheetNazvaniyaRK);
                sum = initTwoConditionReklama(Alphabet.B.ordinal(), listRKNames, Alphabet.G.ordinal(), listActualDocuments, Alphabet.F.ordinal(), 21, i + 2, sheetProchieUderzhaniya, sheetNazvaniyaRK, sheetReportConverted);
                listProchieUderzhaniya.add(sum);
            }

            fileInputStreamProchieUderzhaniya.close();
            fileInputStreamNazvaniyaRK.close();


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
            double prihodNaSchetWB = viruchkaPosleVichetaKomissii - logistika - shtrafi - hranenie - prochieUderjania;
            sheetReportConverted.getRow(22).createCell(1).setCellValue(prihodNaSchetWB);
            ArrayList<Double> listPrihodNaSchetWb = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = 0;
                sum = listViruchkaPosleVichetaKomissii.get(i) - listLogistika.get(i) - listHranenie.get(i) - listProchieUderzhaniya.get(i);
                sheetReportConverted.getRow(22).createCell(i + 2).setCellValue(sum);
                listPrihodNaSchetWb.add(sum);
            }

            //налог 7%
            double nalog7Proc = prihodNaSchetWB * 0.07;
            sheetReportConverted.getRow(25).createCell(1).setCellValue(nalog7Proc);
            ArrayList<Double> listNalog7Proc = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = 0;
                sum = listPrihodNaSchetWb.get(i) * 0.07;
                sheetReportConverted.getRow(25).createCell(i + 2).setCellValue(sum);
                listNalog7Proc.add(sum);
            }

            //Себестоимость
            double sebes = 0;
            for (Row row: sheetReportFromWB) {
                String serviceType = row.getCell(Alphabet.K.ordinal()).getStringCellValue();
                if (serviceType.equals("Продажа")) {
                    String sellersArticle = row.getCell(Alphabet.F.ordinal()).getStringCellValue();
                    for (Row rowSebes : sheetSebestoimost) {
                        if (rowSebes.getCell(Alphabet.B.ordinal()).getStringCellValue().equals(sellersArticle)) {
                            sebes += rowSebes.getCell(Alphabet.C.ordinal()).getNumericCellValue();
                        }
                    }
                }
            }
            sheetReportConverted.getRow(23).createCell(1).setCellValue(sum);

            ArrayList<Double> listSebestoimost = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = 0;
                for (Row row: sheetReportFromWB) {
                    String serviceType = row.getCell(Alphabet.K.ordinal()).getStringCellValue();
                    if (serviceType.equals("Продажа")){
                        String category = row.getCell(Alphabet.C.ordinal()).getStringCellValue();
                        if(category.equals(list.get(i))) {
                            String sellersArticle = row.getCell(Alphabet.F.ordinal()).getStringCellValue();
                            for (Row rowSebes : sheetSebestoimost) {
                                if (rowSebes.getCell(Alphabet.B.ordinal()).getStringCellValue().equals(sellersArticle)) {
                                    sum += rowSebes.getCell(Alphabet.C.ordinal()).getNumericCellValue();
                                }
                            }
                        }
                    }
                }
                sheetReportConverted.getRow(23).createCell(i + 2).setCellValue(sum);
                listSebestoimost.add(sum);
            }




            //Фулфилмент
            double fulfilmentSum = 0;
            for (Row rowReport : sheetReportFromWB) {
                String stringCellValue = rowReport.getCell(Alphabet.K.ordinal()).getStringCellValue();
                if (stringCellValue.equals("Продажа")) {
                    String sellersArticle = rowReport.getCell(Alphabet.F.ordinal()).getStringCellValue();
                    for (Row rowSebes : sheetSebestoimost) {
                        if (rowSebes.getCell(Alphabet.B.ordinal()).getStringCellValue().equals(sellersArticle)) {
                            fulfilmentSum += rowSebes.getCell(Alphabet.D.ordinal()).getNumericCellValue();
                        }
                    }
                }
            }
            sheetReportConverted.getRow(24).createCell(1).setCellValue(fulfilmentSum);

            ArrayList<Double> listFulfilment = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = 0;
                for (Row row: sheetReportFromWB) {
                    String serviceType = row.getCell(Alphabet.K.ordinal()).getStringCellValue();
                    if (serviceType.equals("Продажа")){
                        String category = row.getCell(Alphabet.C.ordinal()).getStringCellValue();
                        if(category.equals(list.get(i))) {
                            String sellersArticle = row.getCell(Alphabet.F.ordinal()).getStringCellValue();
                            for (Row rowSebes : sheetSebestoimost) {
                                if (rowSebes.getCell(Alphabet.B.ordinal()).getStringCellValue().equals(sellersArticle)) {
                                    sum += rowSebes.getCell(Alphabet.D.ordinal()).getNumericCellValue();
                                }
                            }
                        }
                    }
                }
                sheetReportConverted.getRow(24).createCell(i + 2).setCellValue(sum);
                listFulfilment.add(sum);
            }

            //прибыль на маркетинг
            double pribilNaMarketing = prihodNaSchetWB - sebes - fulfilmentSum - nalog7Proc;
            sheetReportConverted.getRow(26).createCell(1).setCellValue(pribilNaMarketing);
            ArrayList<Double> listPribilNaMarketing = new ArrayList<>();
            for (int i = 0; i < list.size(); i++) {
                sum = listPrihodNaSchetWb.get(i) - listSebestoimost.get(i) - listFulfilment.get(i) - listNalog7Proc.get(i);
                sheetReportConverted.getRow(26).createCell(i + 2).setCellValue(sum);
                listPribilNaMarketing.add(sum);
            }

            //прибыль на маркетинг на 1 штуку
            sheetReportConverted.getRow(27).createCell(1).setCellValue(pribilNaMarketing / poluchenyDengiShtuk);
            for (int i = 0; i < list.size(); i++) {
                if (listPolucheniDengiShtuk.get(i) != 0)
                    sum = listPribilNaMarketing.get(i) / listPolucheniDengiShtuk.get(i);
                else
                    sum = 0;
                sheetReportConverted.getRow(27).createCell(i + 2).setCellValue(sum);
            }


            fillRowGreen(1, sheetReportConverted);
            fillRowGreen(6, sheetReportConverted);
            fillRowGreen(10, sheetReportConverted);
            fillRowGreen(22, sheetReportConverted);
            fillRowGreen(26, sheetReportConverted);

            reportConverted.write(fileOutputStream);
            System.out.println("Файл создан");
            fileOutputStream.close();
            fileInputStreamReportFromWB.close();
            fileInputStreamSebestoimost.close();
        } catch (Exception e) {
            System.out.println(e.toString());
        }
    }

    private static void fillRowGreen(int rowNum, XSSFSheet sheetReportConverted) {
        Row row = sheetReportConverted.getRow(rowNum);
        if (row != null) {
            int lastCellNum = row.getLastCellNum();
            for (int i = 0; i < lastCellNum; i++) {
                Cell cell = row.getCell(i);
                cell.setCellStyle(styleGreen);
            }
        }
    }


    public static void initFirstColumn(String str, Sheet sheet) {
        Row row = sheet.createRow(counter);
        row.createCell(0).setCellValue(str);
        counter++;
    }

    public static double initOneCondition(int columnToCheck, String condition, int columnToSum, int rowToSet, int columnToSet, Sheet sheetFrom, Sheet sheetReportConverted) {
        sum = 0;
        for (Row row : sheetFrom) {
            Cell cellToSum = row.getCell(columnToSum);
            Cell cellToChek = row.getCell(columnToCheck);
            CellType cellToCheckType = cellToChek.getCellType();
            String stringCellToCheckValue = "";

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
        Row row = sheetReportConverted.getRow(rowToSet);
        row.createCell(columnToSet).setCellValue(sum);
        return sum;
    }

    private static double initTwoCondition(int columnToCheck1, String condition1, int columnToCheck2, String condition2, int columnToSum, int rowToSet, int columnToSet, Sheet sheetFrom, Sheet sheetReportConverted) {
        sum = 0;
        for (Row row : sheetFrom) {
            Cell cellToSum = row.getCell(columnToSum);

            Cell cellToCheck1 = row.getCell(columnToCheck1);
            if (cellToCheck1 == null)
                continue;
            CellType cellToCheck1Type = cellToCheck1.getCellType();
            String stringCellToCheckValue1 = "";
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
            String stringCellToCheckValue2 = "";
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

    private static double initTwoConditionReklama(int columnRKNames, ArrayList listRKNames, int columnDocNum, ArrayList listDocNum, int columnToSum, int rowToSet, int columnToSet, Sheet sheetFrom, Sheet sheetNazvaniyaRK, Sheet sheetReportConverted) {
        sum = 0;
        for (Row row : sheetFrom) {
            if (row.getRowNum() != 0) {
                Cell cellToSum = row.getCell(columnToSum);

                Cell cellToCheck1 = row.getCell(columnRKNames);
                if (cellToCheck1 == null)
                    continue;
                CellType cellToCheck1Type = cellToCheck1.getCellType();
                String stringCellToCheckValue1 = "";
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
                int intCellToCheckValue2 = 0;
                if (cellToCheck2Type == CellType.STRING) {
                    String stringCellToCheckValue2 = cellToCheck2.getStringCellValue();
                    intCellToCheckValue2 = Integer.parseInt(stringCellToCheckValue2);
                } else {
                    intCellToCheckValue2 = (int)cellToCheck2.getNumericCellValue();
                }


                if (listRKNames.contains(stringCellToCheckValue1) && listDocNum.contains(intCellToCheckValue2) && cellToSum != null && cellToSum.getCellType() == CellType.NUMERIC) {
                    sum += cellToSum.getNumericCellValue();
                }
            }
        }
        sheetReportConverted.getRow(rowToSet).createCell(columnToSet).setCellValue(sum);
        return sum;
    }


    private static double initSumInColumn(int columnToSum, int rowToSet, int columnToSet, XSSFSheet sheetReportFromWB, XSSFSheet sheetReportConverted) {
        sum = 0;
        for (Row row : sheetReportFromWB) {
            Cell cell = row.getCell(columnToSum);
            if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                sum += cell.getNumericCellValue();
            }
        }
        sheetReportConverted.getRow(rowToSet).createCell(columnToSet).setCellValue(sum);
        return sum;
    }

    private static ArrayList<String> getListRKNamesForCategory(String category, Sheet sheetNazvaniyaRK) {
        ArrayList list = new ArrayList<>();
        for(Row row : sheetNazvaniyaRK) {
            Cell cellCategory = row.getCell(Alphabet.B.ordinal());
            String rowCategory = cellCategory.getStringCellValue();
            if(rowCategory.equals(category)) {
                Cell cellRKName = row.getCell(Alphabet.A.ordinal());
                String rkName = cellRKName.getStringCellValue();
                list.add(rkName);
            }
        }
        return list;
    }
}

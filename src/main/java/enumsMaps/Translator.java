package enumsMaps;

import java.util.HashMap;
import java.util.Map;

public class Translator {
    public static Map<String, String> translateMap = Map.ofEntries(
            Map.entry("retail_amount","Вайлдберриз реализовал Товар (Пр)"),
            Map.entry("supplier_oper_name","Обоснование для оплаты"),
            Map.entry("quantity","Кол-во"),
            Map.entry("ppvz_for_pay","К перечислению Продавцу за реализованный Товар"),
            Map.entry("doc_type_name","Тип документа"),
            Map.entry("delivery_rub","Услуги по доставке товара покупателю"),
            Map.entry("delivery_amount","Количество доставок"),
            Map.entry("return_amount","Количество возврата"),
            Map.entry("subject_name","Предмет"),
            Map.entry("sa_name","Артикул поставщика"),
            Map.entry("storage_fee","Хранение"),
            Map.entry("deduction","Удержания"),
            Map.entry("penalty","Общая сумма штрафов"),
            Map.entry("acceptance","Платная приемка"),
            Map.entry("updNum","Номер документа"),
            Map.entry("updSum","Сумма"),
            Map.entry("campName","Кампания"),
            Map.entry("warehousePrice","Сумма хранения, руб"),
            Map.entry("vendorCode","Артикул продавца"),
            Map.entry("subject","Предмет"),
            Map.entry("saleInvoiceCostPrice","К перечислению за товар, руб."),
            Map.entry("sa","Артикул продавца"),
            Map.entry("bonus_type_name","Виды логистики, штрафов и доплат")
    );
}

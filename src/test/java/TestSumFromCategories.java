import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;


public class TestSumFromCategories {
    private final TestsHelper testsHelper = new TestsHelper();

    @Test
    public void testPrihodNaSchetWBFromCategories0807_1407() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты0807_1407\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPrihodNaSchetWBSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPrihodNaSchetWBFromCategoriesExpected(path);
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPribilNaMarketingFromCategories0807_1407() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты0807_1407\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPribilNaMarketingSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPribilNaMarketingFromCategoriesExpected(path);
        if (Math.abs(sumInt - sumExpectedInt) <= 1) {
            sumInt = sumExpectedInt;
        }
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPrihodNaSchetWBFromCategories0107_0707() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты0107_0707\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPrihodNaSchetWBSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPrihodNaSchetWBFromCategoriesExpected(path);
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPribilNaMarketingFromCategories0107_0707() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты0107_0707\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPribilNaMarketingSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPribilNaMarketingFromCategoriesExpected(path);
        if (Math.abs(sumInt - sumExpectedInt) <= 1) {
            sumInt = sumExpectedInt;
        }
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPrihodNaSchetWBFromCategories2406_3006() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты2406_3006\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPrihodNaSchetWBSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPrihodNaSchetWBFromCategoriesExpected(path);
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPribilNaMarketingFromCategories2406_3006() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты2406_3006\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPribilNaMarketingSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPribilNaMarketingFromCategoriesExpected(path);
        if (Math.abs(sumInt - sumExpectedInt) <= 1) {
            sumInt = sumExpectedInt;
        }
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPrihodNaSchetWBFromCategories1706_2306() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты1706_2306\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPrihodNaSchetWBSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPrihodNaSchetWBFromCategoriesExpected(path);
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPribilNaMarketingFromCategories1706_2306() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты1706_2306\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPribilNaMarketingSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPribilNaMarketingFromCategoriesExpected(path);
        if (Math.abs(sumInt - sumExpectedInt) <= 1) {
            sumInt = sumExpectedInt;
        }
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPrihodNaSchetWBFromCategories1006_1606() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты1006_1606\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPrihodNaSchetWBSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPrihodNaSchetWBFromCategoriesExpected(path);
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPribilNaMarketingFromCategories1006_1606() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты1006_1606\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPribilNaMarketingSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPribilNaMarketingFromCategoriesExpected(path);
        if (Math.abs(sumInt - sumExpectedInt) <= 1) {
            sumInt = sumExpectedInt;
        }
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPrihodNaSchetWBFromCategories0306_0906() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты0306_0906\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPrihodNaSchetWBSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPrihodNaSchetWBFromCategoriesExpected(path);
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPribilNaMarketingFromCategories0306_0906() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты0306_0906\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPribilNaMarketingSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPribilNaMarketingFromCategoriesExpected(path);
        if (Math.abs(sumInt - sumExpectedInt) <= 1) {
            sumInt = sumExpectedInt;
        }
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPrihodNaSchetWBFromCategories2705_0206() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты2705_0206\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPrihodNaSchetWBSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPrihodNaSchetWBFromCategoriesExpected(path);
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPribilNaMarketingFromCategories2705_0206() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты2705_0206\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPribilNaMarketingSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPribilNaMarketingFromCategoriesExpected(path);
        if (Math.abs(sumInt - sumExpectedInt) <= 1) {
            sumInt = sumExpectedInt;
        }
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPrihodNaSchetWBFromCategories2005_2605() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты2005_2605\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPrihodNaSchetWBSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPrihodNaSchetWBFromCategoriesExpected(path);
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPribilNaMarketingFromCategories2005_2605() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты2705_0206\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPribilNaMarketingSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPribilNaMarketingFromCategoriesExpected(path);
        if (Math.abs(sumInt - sumExpectedInt) <= 1) {
            sumInt = sumExpectedInt;
        }
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPrihodNaSchetWBFromCategories1305_1905() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты1305_1905\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPrihodNaSchetWBSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPrihodNaSchetWBFromCategoriesExpected(path);
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPribilNaMarketingFromCategories1305_1905() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты1305_1905\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPribilNaMarketingSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPribilNaMarketingFromCategoriesExpected(path);
        if (Math.abs(sumInt - sumExpectedInt) <= 1) {
            sumInt = sumExpectedInt;
        }
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPrihodNaSchetWBFromCategories0605_1205() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты0605_1205\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPrihodNaSchetWBSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPrihodNaSchetWBFromCategoriesExpected(path);
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPribilNaMarketingFromCategories0605_1205() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты0605_1205\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPribilNaMarketingSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPribilNaMarketingFromCategoriesExpected(path);
        if (Math.abs(sumInt - sumExpectedInt) <= 1) {
            sumInt = sumExpectedInt;
        }
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPrihodNaSchetWBFromCategories2904_0505() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты2904_0505\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPrihodNaSchetWBSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPrihodNaSchetWBFromCategoriesExpected(path);
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPribilNaMarketingFromCategories2904_0505() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты2904_0505\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPribilNaMarketingSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPribilNaMarketingFromCategoriesExpected(path);
        if (Math.abs(sumInt - sumExpectedInt) <= 1) {
            sumInt = sumExpectedInt;
        }
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPrihodNaSchetWBFromCategories2204_2804() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты2204_2804\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPrihodNaSchetWBSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPrihodNaSchetWBFromCategoriesExpected(path);
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPribilNaMarketingFromCategories2204_2804() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты2204_2804\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPribilNaMarketingSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPribilNaMarketingFromCategoriesExpected(path);
        if (Math.abs(sumInt - sumExpectedInt) <= 1) {
            sumInt = sumExpectedInt;
        }
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPrihodNaSchetWBFromCategories1504_2104() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты1504_2104\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPrihodNaSchetWBSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPrihodNaSchetWBFromCategoriesExpected(path);
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPribilNaMarketingFromCategories1504_2104() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты1504_2104\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPribilNaMarketingSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPribilNaMarketingFromCategoriesExpected(path);
        if (Math.abs(sumInt - sumExpectedInt) <= 1) {
            sumInt = sumExpectedInt;
        }
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPrihodNaSchetWBFromCategories0804_1404() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты0804_1404\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPrihodNaSchetWBSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPrihodNaSchetWBFromCategoriesExpected(path);
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPribilNaMarketingFromCategories0804_1404() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты0804_1404\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPribilNaMarketingSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPribilNaMarketingFromCategoriesExpected(path);
        if (Math.abs(sumInt - sumExpectedInt) <= 1) {
            sumInt = sumExpectedInt;
        }
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPrihodNaSchetWBFromCategories0104_0704() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты0104_0704\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPrihodNaSchetWBSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPrihodNaSchetWBFromCategoriesExpected(path);
        assertEquals(sumExpectedInt, sumInt);
    }

    @Test
    public void testPribilNaMarketingFromCategories0104_0704() throws IOException {
        String path = "e:\\Proga\\AnaliticsWB\\Отчеты0104_0704\\UnitЭкономика.xlsx";
        int sumInt = testsHelper.getPribilNaMarketingSumFromCategories(path);
        int sumExpectedInt = testsHelper.getPribilNaMarketingFromCategoriesExpected(path);
        if (Math.abs(sumInt - sumExpectedInt) <= 1) {
            sumInt = sumExpectedInt;
        }
        assertEquals(sumExpectedInt, sumInt);
    }
}

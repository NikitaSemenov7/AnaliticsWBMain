import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

class TestCreation {
    private final TestsHelper testsHelper = new TestsHelper();

    @Test
    void testPrihodNaSchetWB0807_1407() throws IOException {
        double prihodNaSchetWB = testsHelper.createUnitEcomonicaAndGetPrihodNaScetWB("e:\\Proga\\AnaliticsWB\\Отчеты0807_1407\\");
        double prihodNaSchetWBExpected = 102476.6;
        assertEquals(prihodNaSchetWB, prihodNaSchetWBExpected);
    }

    @Test
    void testPrihodNaSchetWB0107_0707SOtrabotkoyMain() throws IOException {
        double prihodNaSchetWB = testsHelper.createUnitEcomonicaAndGetPrihodNaScetWB("e:\\Proga\\AnaliticsWB\\Отчеты0107_0707\\");
        double prihodNaSchetWBExpected = 84855.29;
        assertEquals(prihodNaSchetWB, prihodNaSchetWBExpected);
    }

    @Test
    void testPrihodNaSchetWB2406_3006SOtrabotkoyMain() throws IOException{
        double prihodNaSchetWB = testsHelper.createUnitEcomonicaAndGetPrihodNaScetWB("e:\\Proga\\AnaliticsWB\\Отчеты2406_3006\\");
        double prihodNaSchetWBExpected = 37726.24;
        assertEquals(prihodNaSchetWB, prihodNaSchetWBExpected);
    }

    @Test
    void testPrihodNaSchetWB1706_2306SOtrabotkoyMain() throws IOException{
        double prihodNaSchetWB = testsHelper.createUnitEcomonicaAndGetPrihodNaScetWB("e:\\Proga\\AnaliticsWB\\Отчеты1706_2306\\");
        double prihodNaSchetWBExpected = 64829.55;
        assertEquals(prihodNaSchetWB, prihodNaSchetWBExpected);
    }



    @Test
    void testPrihodNaSchetWB1006_1606() throws IOException {
        double prihodNaSchetWB = testsHelper.createUnitEcomonicaAndGetPrihodNaScetWB("e:\\Proga\\AnaliticsWB\\Отчеты1006_1606\\");
        double prihodNaSchetWBExpected = 47073.83;
        assertEquals(prihodNaSchetWB, prihodNaSchetWBExpected);
    }

    @Test
    void testPrihodNaSchetWB0306_0906() throws IOException {
        double prihodNaSchetWB = testsHelper.createUnitEcomonicaAndGetPrihodNaScetWB("e:\\Proga\\AnaliticsWB\\Отчеты0306_0906\\");
        double prihodNaSchetWBExpected = 62729.55;
        assertEquals(prihodNaSchetWB, prihodNaSchetWBExpected);
    }

    @Test
    void testPrihodNaSchetWB2705_0206() throws IOException {
        double prihodNaSchetWB = testsHelper.createUnitEcomonicaAndGetPrihodNaScetWB("e:\\Proga\\AnaliticsWB\\Отчеты2705_0206\\");
        double prihodNaSchetWBExpected = 31188.59;
        assertEquals(prihodNaSchetWB, prihodNaSchetWBExpected);
    }

    @Test
    void testPrihodNaSchetWB2005_2605() throws IOException {
        double prihodNaSchetWB = testsHelper.createUnitEcomonicaAndGetPrihodNaScetWB("e:\\Proga\\AnaliticsWB\\Отчеты2005_2605\\");
        double prihodNaSchetWBExpected = 18772.95;
        assertEquals(prihodNaSchetWB, prihodNaSchetWBExpected);
    }

    @Test
    void testPrihodNaSchetWB1305_1905() throws IOException {
        double prihodNaSchetWB = testsHelper.createUnitEcomonicaAndGetPrihodNaScetWB("e:\\Proga\\AnaliticsWB\\Отчеты1305_1905\\");
        double prihodNaSchetWBExpected = 28362.38;
        assertEquals(prihodNaSchetWB, prihodNaSchetWBExpected);
    }

    @Test
    void testPrihodNaSchetWB0605_1205() throws IOException {
        double prihodNaSchetWB = testsHelper.createUnitEcomonicaAndGetPrihodNaScetWB("e:\\Proga\\AnaliticsWB\\Отчеты0605_1205\\");
        double prihodNaSchetWBExpected = 12062.69;
        assertEquals(prihodNaSchetWB, prihodNaSchetWBExpected);
    }

    @Test
    void testPrihodNaSchetWB2904_0505() throws IOException {
        double prihodNaSchetWB = testsHelper.createUnitEcomonicaAndGetPrihodNaScetWB("e:\\Proga\\AnaliticsWB\\Отчеты2904_0505\\");
        double prihodNaSchetWBExpected = 5906.07;
        assertEquals(prihodNaSchetWB, prihodNaSchetWBExpected);
    }

    @Test
    void testPrihodNaSchetWB2204_2804() throws IOException {
        double prihodNaSchetWB = testsHelper.createUnitEcomonicaAndGetPrihodNaScetWB("e:\\Proga\\AnaliticsWB\\Отчеты2204_2804\\");
        double prihodNaSchetWBExpected = 5444.92;
        assertEquals(prihodNaSchetWB, prihodNaSchetWBExpected);
    }

    @Test
    void testPrihodNaSchetWB1504_2104() throws IOException {
        double prihodNaSchetWB = testsHelper.createUnitEcomonicaAndGetPrihodNaScetWB("e:\\Proga\\AnaliticsWB\\Отчеты1504_2104\\");
        double prihodNaSchetWBExpected = 2652.19;
        assertEquals(prihodNaSchetWB, prihodNaSchetWBExpected);
    }

    @Test
    void testPrihodNaSchetWB0804_1404() throws IOException {
        double prihodNaSchetWB = testsHelper.createUnitEcomonicaAndGetPrihodNaScetWB("e:\\Proga\\AnaliticsWB\\Отчеты0804_1404\\");
        double prihodNaSchetWBExpected = 2197.37;
        assertEquals(prihodNaSchetWB, prihodNaSchetWBExpected);
    }

    @Test
    void testPrihodNaSchetWB0104_0704() throws IOException {
        double prihodNaSchetWB = testsHelper.createUnitEcomonicaAndGetPrihodNaScetWB("e:\\Proga\\AnaliticsWB\\Отчеты0104_0704\\");
        double prihodNaSchetWBExpected = -615.45;
        assertEquals(prihodNaSchetWB, prihodNaSchetWBExpected);
    }
}
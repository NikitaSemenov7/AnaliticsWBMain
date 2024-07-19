import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

class TestCreation {
    private final TestsHelper testsHelper = new TestsHelper();

    @Test
    void testPrihodNaSchetWB() throws IOException, InterruptedException {
        double prihodNaSchetWB = testsHelper.createUnitEcomonicaAndGetPrihodNaScetWB("e:\\Proga\\AnaliticsWB\\Отчеты2024-09-30_2024-10-06\\");
        double prihodNaSchetWBExpected = 13610.39 + 115368.17;
        assertEquals(prihodNaSchetWBExpected, prihodNaSchetWB);
    }

}
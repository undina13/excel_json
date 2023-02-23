import controller.ExcelGenerator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.IOException;

public class ExcelGeneratorTest {
    public static final String SYS_PROPERTY_TMPDIR = "java.io.tmpdir";

    @Test
    public void testExample_1() throws IOException {
        ExcelGenerator excelGenerator = new ExcelGenerator();
        String jsonFile = excelGenerator.convertJsonToString("src/main/resources/example_1.json");
        excelGenerator.generateNameColumns("column1", "column2", "column3", "column4", "column5");
        excelGenerator.convertJsonToExcel(jsonFile, "test1");

        Workbook workbook;
        String wPath = System.getProperty(SYS_PROPERTY_TMPDIR);
        try (FileInputStream fileStream = new FileInputStream(wPath + "/" + "test1.xlsx")) {
            workbook = new XSSFWorkbook(fileStream);
        }
        Sheet sheet = workbook.getSheetAt(0);

        String[] title = new String[5];
        for (int i = 0; i < 5; i++) {
            title[i] = sheet.getRow(0).getCell(i).toString();
        }
        Assert.assertArrayEquals(title, new String[]{"column1", "column2", "column3", "column4", "column5"});

        String[] line1 = new String[5];
        for (int i = 0; i < 5; i++) {
            line1[i] = sheet.getRow(1).getCell(i).toString();
        }
        Assert.assertArrayEquals(line1, new String[]{"Винты", "580.0", "шт", "30ХГСА", "4.634"});

        String[] line2 = new String[5];
        for (int i = 0; i < 5; i++) {
            line2[i] = sheet.getRow(2).getCell(i).toString();
        }
        Assert.assertArrayEquals(line2, new String[]{"Саморезы", "329.0", "шт", "30ХГСА", "6.432"});

        String[] line3 = new String[5];
        for (int i = 0; i < 5; i++) {
            line3[i] = sheet.getRow(3).getCell(i).toString();
        }
        Assert.assertArrayEquals(line3, new String[]{"Герметик", "1.0", "кг", "Не применимо", "0.024"});

        String[] line4 = new String[5];
        for (int i = 0; i < 5; i++) {
            line4[i] = sheet.getRow(4).getCell(i).toString();
        }
        Assert.assertArrayEquals(line4, new String[]{"Гайки", "121.2", "шт", "30ХГСА", "1.024"});
    }


    @Test
    public void testExample_2() throws IOException {
        ExcelGenerator excelGenerator = new ExcelGenerator();
        String jsonFile = excelGenerator.convertJsonToString("src/main/resources/example_2.json");
        excelGenerator
                .generateNameColumns("Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8");
        excelGenerator.convertJsonToExcel(jsonFile, "test2");

        Workbook workbook;
        String wPath = System.getProperty(SYS_PROPERTY_TMPDIR);
        try (FileInputStream fileStream = new FileInputStream(wPath + "/" + "test2.xlsx")) {
            workbook = new XSSFWorkbook(fileStream);
        }
        Sheet sheet = workbook.getSheetAt(0);

        String[] title = new String[8];
        for (int i = 0; i < 8; i++) {
            title[i] = sheet.getRow(0).getCell(i).toString();
        }
        Assert.assertArrayEquals(title, new String[]{"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8"});

        String[] line1 = new String[8];
        for (int i = 0; i < 8; i++) {
            line1[i] = sheet.getRow(1).getCell(i).toString();
        }
        Assert.assertArrayEquals(line1, new String[]{"Иванов", "Иван", "Иванович", "м", "32.0", "187.0", "85.2", "true"});

        String[] line2 = new String[8];
        for (int i = 0; i < 8; i++) {
            line2[i] = sheet.getRow(2).getCell(i).toString();
        }
        Assert.assertArrayEquals(line2, new String[]{"Петров", "Петр", "Петрович", "м", "22.0", "181.0", "66.1", "false"});

        String[] line3 = new String[8];
        for (int i = 0; i < 8; i++) {
            line3[i] = sheet.getRow(3).getCell(i).toString();
        }
        Assert.assertArrayEquals(line3, new String[]{"Михайлов", "Михаил", "Михайлович", "м", "44.0", "185.0", "81.0", "true"});

        String[] line4 = new String[8];
        for (int i = 0; i < 8; i++) {
            line4[i] = sheet.getRow(4).getCell(i).toString();
        }
        Assert.assertArrayEquals(line4, new String[]{"Васильева", "Василиса", "Васильевна", "ж", "26.0", "165.0", "51.2", "false"});

        String[] line5 = new String[8];
        for (int i = 0; i < 8; i++) {
            line5[i] = sheet.getRow(5).getCell(i).toString();
        }
        Assert.assertArrayEquals(line5, new String[]{"Александрова", "Александра", "Александровна", "ж", "30.0", "172.0", "56.9", "true"});
    }
}

package controller;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;
import java.util.logging.Level;
import java.util.logging.Logger;

public class ExcelGenerator {
    public static final String SYS_PROPERTY_TMPDIR = "java.io.tmpdir";
    private static final Logger logger = Logger.getLogger(ExcelGenerator.class.getName());
    private Workbook book;
    private Sheet sheet;

    public ExcelGenerator() {
        this.book = new XSSFWorkbook();
        this.sheet = book.createSheet();
    }

    public void generateNameColumns(String... args) throws IOException {
        Row row = sheet.createRow(0);
        for (int i = 0; i < args.length; i++) {
            Cell name = row.createCell(i);
            name.setCellValue(args[i]);
        }
    }

    public String convertJsonToString(String file) throws IOException {
        FileReader fr = new FileReader(file);
        Scanner scan = new Scanner(fr);
        StringBuilder sb = new StringBuilder();
        while (scan.hasNextLine()) {
            sb.append(scan.nextLine());
        }
        fr.close();
        return sb.toString();
    }

    public void convertJsonToExcel(String jsonString, String excelFileName) throws IOException {
        ObjectMapper mapper = new ObjectMapper();
        List<LinkedHashMap<String, Object>> list = mapper
                .readValue(jsonString, new TypeReference<List<LinkedHashMap<String, Object>>>() {
                });
        Integer maxSize = null;
        int numRow = 1;
        for (LinkedHashMap<String, Object> o : list) {
            if (maxSize == null) {
                maxSize = o.size();
            }
            if (o.size() > maxSize) {
                logger.log(Level.WARNING, "В json в каждом объекте должно быть одинаковое количество полей");
            }
            Row row = sheet.createRow(numRow);

            int cellNum = 0;

            for (Map.Entry<String, Object> entry : o.entrySet()) {
                Cell cell = row.createCell(cellNum);
                Object s = entry.getValue();
                Integer intValue = null;
                try {
                    intValue = (Integer) s;
                    cell.setCellValue(intValue);
                } catch (ClassCastException e) {

                }
                if (intValue == null) {
                    try {
                        Double doubleValue = (Double) s;
                        cell.setCellValue(doubleValue);
                    } catch (ClassCastException e) {
                        cell.setCellValue(s.toString());
                    }
                }
                cellNum++;

            }
            numRow++;
        }
        String wPath = System.getProperty(SYS_PROPERTY_TMPDIR);
        book.write(new FileOutputStream(wPath + "/" + excelFileName + ".xlsx"));
    }
}

import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.extern.apachecommons.CommonsLog;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;

@Slf4j
public class Main {

    public static void main(String[] args) {
        JsonToExcelAndBack.test();
    }

    public static class JsonToExcelAndBack {
        public static void test() {
            try {
                // Читання JSON з файлу
                FileInputStream jsonFile = new FileInputStream("placeholders.json");
                ObjectMapper objectMapper = new ObjectMapper();
                List<Object> jsonData = objectMapper.readValue(jsonFile, List.class);

                // Створення нового Excel-документа
                Workbook workbook = new XSSFWorkbook();
                Sheet sheet = workbook.createSheet("Дані");

                // Запис JSON-даних в Excel
                export(jsonData, sheet);

                // Збереження Excel-файлу
                FileOutputStream excelFile = new FileOutputStream("дані.xlsx");
                workbook.write(excelFile);
                excelFile.close();

                // Тепер можемо прочитати дані з Excel із файлу "дані.xlsx" і конвертувати їх назад у JSON,
                // але цей код не надається тут.

               log.info("Конвертація завершена успішно.");
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

        private static void export(List<Object> jsonData, Sheet sheet) {
            int rowNum = 0;
            for (Object row : jsonData) {
                Row sheetRow = sheet.createRow(rowNum++);
                int cellNum = 0;
                for (Map.Entry<String, Object> entry : ((Map<String, Object>) row).entrySet()) {
                    Cell cell = sheetRow.createCell(cellNum++);
                    Object value = entry.getValue();
                    if (value instanceof Number) {
                        cell.setCellValue((Double) value);
                    }
                    if (value instanceof String) {
                        cell.setCellValue((String) value);
                    }
                    if (value instanceof List<?>) {
                        rowNum = getRowNum((List<Map<String, String>>) value, sheet, rowNum);
                        sheet.createRow(rowNum++);
                    }
                }
                sheet.createRow(rowNum++);
            }
        }

        private static int getRowNum(List<Map<String, String>> valueMap, Sheet sheet, int rowNum) {
            for (Map<String, String> map : valueMap) {
                sheet.createRow(rowNum++);
                int cellNum = 0;
                for (Map.Entry<String, String> entry : map.entrySet()) {
                    Row row = sheet.createRow(rowNum++);
                    Cell cell = row.createCell(cellNum++);
                    cell.setCellValue(entry.getKey());
                    Object value = entry.getValue();
                    Cell cellNext = row.createCell(cellNum);
                    if (value instanceof Number) {
                        cellNext.setCellValue((Double) value);
                    }
                    if (value instanceof String) {
                        cellNext.setCellValue((String) value);
                    }
                    if (value instanceof List<?>) {
                        rowNum = getRowNum((List<Map<String, String>>) value, sheet, rowNum);
                    }
                    cellNum = 0;
                }
            }
            return rowNum;
        }
    }
}

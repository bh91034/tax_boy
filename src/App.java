import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * POI LIB
 * Reference :
 *  - https://myhappyman.tistory.com/198
 *  - https://yangsosolife.tistory.com/7
 */
public class App {
    public static void main(String[] args) {
        String path = "C:/project/tax_helper/test_data/";
        String fileName = "test.xlsx";

        List<Map<Object, Object>> excelData = readExcel(path, fileName);

        for (int i = 0; i < excelData.size(); i++) {
            System.out.println(excelData.get(i));
        }
    }

    public static List<Map<Object, Object>> readExcel(String path, String fileName) {
        List<Map<Object, Object>> list = new ArrayList<>();
        if (path == null || fileName == null) {
            return list;
        }

        FileInputStream is = null;
        File excel = new File(path + fileName);
        try {
            is = new FileInputStream(excel);
            Workbook workbook = null;
            if (fileName.endsWith(".xls")) {
                workbook = new HSSFWorkbook(is);
            } else if (fileName.endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(is);
            }

            if (workbook != null) {
                int sheets = workbook.getNumberOfSheets();
                getSheet(workbook, sheets, list);
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (is != null) {
                try {
                    is.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        return list;
    }

    public static void getSheet(Workbook workbook, int sheets, List<Map<Object, Object>> list) {
        for (int z = 0; z < sheets; z++) {
            Sheet sheet = workbook.getSheetAt(z);
            int rows = sheet.getLastRowNum();
            getRow(sheet, rows, list);
        }
    }

    public static void getRow(Sheet sheet, int rows, List<Map<Object, Object>> list) {
        for (int i = 0; i <= rows; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                int cells = row.getPhysicalNumberOfCells();
                list.add(getCell(row, cells));
            }
        }
    }

    public static Map<Object, Object> getCell(Row row, int cells) {
        String[] columns = { "column1", "column2", "column3", "column4", "column5", "column6" };
        Map<Object, Object> map = new HashMap<>();
        for (int j = 0; j < cells; j++) {
            if (j >= columns.length) {
                break;
            }

            Cell cell = row.getCell(j);
            if (cell != null) {
                switch (cell.getCellType()) {
                case BLANK:
                    map.put(columns[j], "");
                    break;
                case STRING:
                    map.put(columns[j], cell.getStringCellValue());
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        map.put(columns[j], cell.getDateCellValue());
                    } else {
                        map.put(columns[j], cell.getNumericCellValue());
                    }
                    break;
                case ERROR:
                    map.put(columns[j], cell.getErrorCellValue());
                    break;
                default:
                    map.put(columns[j], "");
                    break;
                }
            }
        }

        return map;
    }
}
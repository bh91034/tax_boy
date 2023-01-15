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
        String path = "test_data/";

        File folder = new File(path);
        File[] listOfFiles = folder.listFiles();

        for (int i = 0; i < listOfFiles.length; i++) {
            if (listOfFiles[i].isFile()) {
                System.out.println("File : " + listOfFiles[i].getName());
                List<Map<Object, Object>> excelData = readExcel(path, listOfFiles[i].getName());
                for (int idx = 0; idx < excelData.size(); idx++) {
                    System.out.println(excelData.get(idx));
                }
            }
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
                getSheet(workbook, sheets, list, true);
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

    public static void getSheet(Workbook workbook, int sheets, List<Map<Object, Object>> list, boolean firstRowAsTitle) {
        for (int z = 0; z < sheets; z++) {
            Sheet sheet = workbook.getSheetAt(z);
            int rows = sheet.getLastRowNum();
            if (firstRowAsTitle && rows <= 0) {
                continue;
            } else {
                getRow(sheet, rows, list, firstRowAsTitle);
            }
        }
    }

    public static void getRow(Sheet sheet, int rows, List<Map<Object, Object>> list, boolean firstRowAsTitle) {
        String[] columns = null;
        for (int i = 0; i <= rows; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                int cells = row.getPhysicalNumberOfCells();
                Map<Object, Object> cellMap = getCell(row, cells, columns);
                if (firstRowAsTitle && i == 0) {
                    columns = parseColumnTitle(cellMap);
                } else {
                    list.add(cellMap);
                }
            }
        }
    }

    private static String[] parseColumnTitle(Map<Object, Object> cellMap) {
        String[] values = new String[cellMap.size()];
        int index = 0;
        for (Map.Entry<Object, Object> mapEntry : cellMap.entrySet()) {
            String key = (String)mapEntry.getKey();
            int idx = Integer.valueOf(key.replaceAll("[^0-9]", ""));
            values[idx-1] = (String)mapEntry.getValue();
            index++;
        }
        return values;
    }

    public static Map<Object, Object> getCell(Row row, int cells, String[] columns) {
        if (columns == null) {
            columns = new String[cells];
            for (int i = 0; i< cells; i++) {
                columns[i] = "column"+(i+1);
            }
        }

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
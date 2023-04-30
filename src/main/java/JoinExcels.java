import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class JoinExcels {
    public static void main(String[] args) throws IOException {
        String propsFilePath = System.getProperty("properties");
        Properties properties = new Properties();
        try {
            properties.load(new FileInputStream(propsFilePath));
        } catch (FileNotFoundException e) {
            System.out.println("Properties file '" + propsFilePath + "' not found. Using the properties file present in resources folder instead");
            propsFilePath = "./src/main/resources/excel.properties";
            properties.load(new FileInputStream(propsFilePath));
        }

        JoinExcels excels = new JoinExcels();
        String outFilePath = properties.getProperty("outFilePath").trim() + "\\Out_" + excels.getCurrentDate() + ".xlsx";
        String key1 = properties.getProperty("file1.file2.key").trim().toLowerCase();
        String key2 = properties.getProperty("file1.file3.key").trim().toLowerCase();

        String[] outputColumns = properties.getProperty("output_headers").toLowerCase().split(",");
        for (int i = 0; i < outputColumns.length; i++) {
            outputColumns[i] = outputColumns[i].trim();
        }

        final String[] FILES = {"file1", "file2", "file3"};
        List<XSSFWorkbook> books = new ArrayList<>() {{
            add(new XSSFWorkbook(new FileInputStream(properties.getProperty(FILES[0]))));
            add(new XSSFWorkbook(new FileInputStream(properties.getProperty(FILES[1]))));
            add(new XSSFWorkbook(new FileInputStream(properties.getProperty(FILES[2]))));
        }};
        System.out.println("Properties read from '" + propsFilePath + "'");
        Sheet sheet1 = books.get(0).getSheetAt(0);
        Sheet sheet2 = books.get(1).getSheetAt(0);
        Sheet sheet3 = books.get(2).getSheetAt(0);

        Map<String, Object> rows = new HashMap<>();

        XSSFWorkbook book = new XSSFWorkbook();
        Sheet sheet = book.createSheet("Data");
        FileOutputStream out = new FileOutputStream(outFilePath);

        book = excels.writeHeaders(book, sheet, 0, outputColumns, false);
        System.out.println("Reading data from excel files");

        for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
            rows.putAll(excels.getRowData(FILES[0], sheet1, i));
            String value1 = rows.get(FILES[0] + "." + key1).toString();
            String value2 = rows.get(FILES[0] + "." + key2).toString();

            rows.putAll(excels.getRowDataIfExist(FILES[1], sheet2, key1, value1));
            rows.putAll(excels.getRowDataIfExist(FILES[2], sheet3, key1, value2));

            book = excels.writeExcel(book, sheet, i, outputColumns, rows);
        }
        book.write(out);
        book.close();
        books.forEach(workbook -> {
            try {
                workbook.close();
            } catch (IOException e) {
                System.out.println("Can't close the workbook");
                throw new RuntimeException(e);
            }
        });
        out.close();
        System.out.println("File '" + outFilePath + "' created");
    }

    public Map<String, Object> getRowData(String fileName, Sheet sheet, int rowIndex) {
        int size = sheet.getRow(rowIndex).getLastCellNum();
        Map<String, Object> rowData = new HashMap<>();
        List<String> headers = new LinkedList<>();

        sheet.getRow(0).forEach(cell -> headers.add(cell.getStringCellValue()));

        for (int i = 0; i <= size - 1; i++) {
            Cell cell = sheet.getRow(rowIndex).getCell(i);
            Object value = this.getCellValue(cell);
            rowData.put(fileName + "." + headers.get(i).toLowerCase(), value);
        }
        return rowData;
    }

    public Map<String, Object> getRowDataIfExist(String fileName, Sheet sheet, String key, Object value) {
        for (int j = 1; j <= sheet.getLastRowNum(); j++) {
            Map<String, Object> data = this.getRowData(fileName, sheet, j);
            Object file2value = data.getOrDefault(fileName + "." + key, "NOT_FOUND");
            if (value.toString().equals(file2value.toString())) {
                return data;
            }
        }
        return Collections.emptyMap();
    }

    public Object getCellValue(Cell cell) {
        CellType cellType = cell.getCellType().equals(CellType.FORMULA) ? cell.getCachedFormulaResultType() : cell.getCellType();

        if (cellType.equals(CellType.STRING) || cellType.equals(CellType.BLANK)) {
            return cell.getStringCellValue();
        } else if (cellType.equals(CellType.NUMERIC)) {
            return (DateUtil.isCellDateFormatted(cell)) ? cell.getDateCellValue() : cell.getNumericCellValue();
        } else if (cellType.equals(CellType.BOOLEAN)) {
            return cell.getBooleanCellValue();
        } else {
            return cell.getErrorCellValue();
        }
    }

    public void setCellValue(Cell cell, Object value) {
        if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else {
            cell.setCellValue((String) value);
        }
    }

    public XSSFWorkbook writeExcel(XSSFWorkbook book, Sheet sheet, int rowNum, String[] keys, Map<String, Object> data) {
        Row row = sheet.createRow(rowNum);

        for (int i = 0; i < keys.length; i++) {
            Cell cell = row.createCell(i);
            this.setCellValue(cell, data.get(keys[i]));
        }
        return book;
    }

    public XSSFWorkbook writeHeaders(XSSFWorkbook book, Sheet sheet, int rowNum, String[] keys, boolean includeFileName) {
        Row row = sheet.createRow(rowNum);
        String header = "";

        for (int i = 0; i < keys.length; i++) {
            Cell cell = row.createCell(i);
            if (includeFileName) {
                header = keys[i];
            } else {
                header = keys[i].split("\\.")[1];
            }
            cell.setCellValue(header);
        }
        return book;
    }

    public String getCurrentDate() {
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyymmddHHmmss");
        LocalDateTime ltd = LocalDateTime.now();
        return dtf.format(ltd);
    }
}
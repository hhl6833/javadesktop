package com.hn6833.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import java.io.File;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelUtil {
    private static DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
    private static DataFormatter dataFormatter = new DataFormatter();

    private ExcelUtil() {
    }

    /**
     * 获取excel文件对应的workbook对象实体，支持.xls和.xlsx文件格式
     *
     * @param filePath excle文件路径
     * @return excle文档对应的workbook实体
     * @throws Exception
     */
    private static Workbook getWorkbook(String filePath) throws Exception {
        File file = new File(filePath);
        Workbook workbook = null;
        if (file.exists()) {
            try {
                workbook = WorkbookFactory.create(file);
            } catch (Exception e) {
                throw e;
            }
        } else {
            throw new Exception("[" + filePath + "]文件不存在");
        }
        return workbook;
    }

    /**
     * 获取cell的值
     *
     * @param cell cell表格实体
     * @return cell表格的值
     */
    private static String getCellValue(Cell cell) {
        String value = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                    value = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        value = dateFormat.format(cell.getDateCellValue());
                    } else {
                        value = String.valueOf(dataFormatter.formatCellValue(cell));
                    }
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    value = String.valueOf(cell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    value = String.valueOf(cell.getCellFormula());
                    break;
            }
        }
        return value;
    }

    /**
     * 获取数据key值对应的cell位置
     *
     * @param sheet excel文档单个sheet页对应实体
     * @param keys  key值set数组，例：["a","b","c"]
     * @return key值和对应的cell表格位置的Map集合对象
     */
    private static Map<String, CellReference> getKeysCellRef(Sheet sheet, Set<String> keys) {
        Map<String, CellReference> result = new HashMap<String, CellReference>();

        for (Row row : sheet) {
            for (Cell cell : row) {
                String cv = getCellValue(cell);
                if (keys.contains(cv)) {
                    result.put(cv, new CellReference(cell.getRowIndex(), cell.getColumnIndex()));
                    keys.remove(cv);
                }
                if (keys.size() == 0) {
                    break;
                }
            }
        }

        return result;
    }

    /***
     * 根据keys集合中key单元格位置，获取单个sheet页中所有的数据
     *
     * @param sheet excel文档单个sheet页对应实体
     * @param keys  key值与位置的映射
     * @return 所有数据的集合
     * @throws Exception
     */
    private static List<Map<String, String>> getSheetData(Sheet sheet, Map<String, CellReference> keys) throws Exception {
        //获取数据范围
        int minRowIndex = -1;
        for (Map.Entry<String, CellReference> entry : keys.entrySet()) {
            CellReference cr = entry.getValue();
            if (cr.getRow() > minRowIndex) {
                minRowIndex = cr.getRow();
            }
        }
        //读取数据
        List<Map<String, String>> result = new LinkedList<Map<String, String>>();
        if (minRowIndex != -1) {
            for (int i = minRowIndex; i < sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i + 1);
                if (row == null) {
                    continue;
                }
                boolean isRowBlank = true;
                Map<String, String> o = new HashMap<String, String>();
                for (Map.Entry<String, CellReference> entry : keys.entrySet()) {
                    Cell cell = row.getCell(entry.getValue().getCol());
                    String cellValue = getCellValue(cell);
                    if (cellValue != null && cellValue.length() > 0) {
                        isRowBlank = false;
                    }
                    o.put(entry.getKey(), cellValue);
                }
                if (!isRowBlank) {
                    result.add(o);
                }
            }
        } else {
            throw new Exception("获取数据范围出错");
        }

        return result;
    }

    /***
     * 读取指定路径下的，指定sheet的，与keys对应的数据
     *
     * @param filePath   excel文件路径
     * @param sheetIndex sheet页面索引
     * @param keys       key值set数组，例：["a","b","c"]
     * @return 当页中与keys对应的所有数据的集合
     */
    public static List<Map<String, String>> read(String filePath, int sheetIndex, Set<String> keys) {
        List<Map<String, String>> result = null;
        try {
            Workbook workbook = getWorkbook(filePath);
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            Map<String, CellReference> keysCellRef = getKeysCellRef(sheet, keys);
            result = getSheetData(sheet, keysCellRef);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }
}

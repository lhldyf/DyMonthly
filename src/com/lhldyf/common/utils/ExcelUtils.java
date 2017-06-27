package com.lhldyf.common.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Created by LH on 17/6/27.
 */
public class ExcelUtils {
    private static final String EXCEL_XLS = "xls";
    private static final String EXCEL_XLSX = "xlsx";
    private static final SimpleDateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");

    /**
     * 判断Excel的版本,获取Workbook
     *
     * @param in
     * @param filename
     * @return
     * @throws IOException
     */
    public static Workbook getWorkbok(InputStream in, File file) throws IOException {
        Workbook wb = null;
        if (file.getName().endsWith(EXCEL_XLS)) {     //Excel 2003
            wb = new HSSFWorkbook(in);
        } else if (file.getName().endsWith(EXCEL_XLSX)) {    // Excel 2007/2010
            wb = new XSSFWorkbook(in);
        }
        return wb;
    }

    /**
     * 判断文件是否是excel
     *
     * @throws Exception
     */
    public static void checkExcelVaild(File file) throws Exception {
        if (!file.exists()) {
            throw new Exception("文件不存在");
        }
        if (!(file.isFile() && (file.getName().endsWith(EXCEL_XLS) || file.getName().endsWith(EXCEL_XLSX)))) {
            throw new Exception("文件不是Excel");
        }
    }

    public static List<List<String>> readSheet(String fileName, int index) {

        List<List<String>> list = new ArrayList<>();
        try{
            // 同时支持Excel 2003、2007
            File excelFile = new File(fileName); // 创建文件对象
            FileInputStream is = new FileInputStream(excelFile); // 文件流
            checkExcelVaild(excelFile);
            Workbook workbook = getWorkbok(is, excelFile);
            //Workbook workbook = WorkbookFactory.create(is); // 这种方式 Excel2003/2007/2010都是可以处理的

            //int sheetCount = workbook.getNumberOfSheets(); // Sheet的数量
            /**
             * 设置当前excel中sheet的下标：0开始
             */
            Sheet sheet = workbook.getSheetAt(index);    // 遍历指定的sheet

            for (Row row : sheet) {
                List<String> oneRow = new ArrayList<>();
                for (Cell cell : row) {
                    oneRow.add(readCellValue(cell));
                }
                list.add(oneRow);
            }
        }catch(Exception e) {
            e.printStackTrace();
        }


        return list;
    }

    private static String readCellValue(Cell cell){
        String cellValue = "";
        if (cell.toString() == null) {
            return cellValue;
        }
        int cellType = cell.getCellType();

        switch (cellType) {
            case Cell.CELL_TYPE_STRING:        // 文本
                cellValue = cell.getRichStringCellValue().getString();
                break;
            case Cell.CELL_TYPE_NUMERIC:    // 数字、日期
                if (DateUtil.isCellDateFormatted(cell)) {
                    cellValue = fmt.format(cell.getDateCellValue());
                } else {
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    cellValue = String.valueOf(cell.getRichStringCellValue().getString());
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:    // 布尔型
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_BLANK: // 空白
                cellValue = cell.getStringCellValue();
                break;
            case Cell.CELL_TYPE_ERROR: // 错误
                cellValue = "错误#";
                break;
            case Cell.CELL_TYPE_FORMULA:    // 公式
                // 得到对应单元格的公式
                //cellValue = cell.getCellFormula();
                // 得到对应单元格的字符串
                cell.setCellType(Cell.CELL_TYPE_STRING);
                cellValue = String.valueOf(cell.getRichStringCellValue().getString());
                break;
            default:
                cellValue = "";
        }
        return cellValue;
    }

    public static void writeToExcel(Map<String,List<String>> map,String path){
        if(null == map || map.isEmpty()) {
            System.out.println("list为空，不执行写Excel操作");
        }

        XSSFWorkbook workbook = new XSSFWorkbook();
        //获取参数个数作为excel列数
        int columeCount = 6;
        //获取List size作为excel行数
        int rowCount = 20;

        XSSFSheet sheet = workbook.createSheet("sheet name");

        int i = 0;
        for (String key :map.keySet()) {
            List<String> oneRow = map.get(key);
            //System.out.println("准备写源文件名为："+key+"的数据");
            XSSFRow row = sheet.createRow(i);
            for (int j = 0; j < oneRow.size(); j++) {
                row.createCell(j);
                row.getCell(j).setCellValue(oneRow.get(j));
            }
            i++;
        }


        //写到磁盘上
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(new File(path));
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
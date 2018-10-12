package com.louie.util.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author maoliang
 * @date 2018/10/12
 */
public class ExcelUtil {
    private final static String xls = "xls";
    private final static String xlsx = "xlsx";

    /**
     * 保存workbook
     *
     * @param targetFile
     * @param workbook
     * @throws IOException
     */
    public static void saveWorkbook(File targetFile, Workbook workbook) throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream(targetFile);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();
    }

    /**
     * 读入excel文件，解析后返回,key为第一行内容
     *
     * @throws IOException
     */
    public static Map<String, List<String>> readExcel(Sheet sheet) throws IOException {
        //获得当前sheet的开始行
        int firstRowNum = sheet.getFirstRowNum();
        //获得当前sheet的结束行
        int lastRowNum = sheet.getLastRowNum();
        //获取标题行
        Row titleRow = sheet.getRow(firstRowNum);
        //获得当前行的开始列
        int firstCellNum = titleRow.getFirstCellNum();
        //获得当前行的列数
        int lastCellNum = titleRow.getLastCellNum();

        Map<String, List<String>> result = new HashMap<String, List<String>>(lastCellNum);
        //获取标题
        String[] titles = new String[lastCellNum];
        for (int cellNum = firstCellNum; cellNum < lastCellNum; cellNum++) {
            Cell cell = titleRow.getCell(cellNum);
            titles[cellNum] = getCellValue(cell);
            result.put(titles[cellNum], new ArrayList<String>());
        }
        //循环除了第一行的所有行
        for (int rowNum = firstRowNum + 1; rowNum <= lastRowNum; rowNum++) {
            //获得当前行
            Row row = sheet.getRow(rowNum);
            if (row == null) {
                continue;
            }
            //循环当前行
            for (int cellNum = firstCellNum; cellNum < lastCellNum; cellNum++) {
                Cell cell = row.getCell(cellNum);
                result.get(titles[cellNum]).add(getCellValue(cell));
            }
        }
        return result;
    }

    public static boolean checkFile(File file) throws IOException {
        //判断文件是否存在
        if (null == file) {
            return false;
        }
        //获得文件名
        String fileName = file.getName();
        //判断文件是否是excel文件
        if (!fileName.endsWith(xls) && !fileName.endsWith(xlsx)) {
            return false;
        }
        return true;
    }

    public static Workbook getWorkBook(File file) {
        //获得文件名
        String fileName = file.getName();
        //创建Workbook工作薄对象，表示整个excel
        Workbook workbook = null;
        try {
            //获取excel文件的io流
            InputStream is = new FileInputStream(file);
            //根据文件后缀名不同(xls和xlsx)获得不同的Workbook实现类对象
            if (fileName.endsWith(xls)) {
                //2003
                workbook = new HSSFWorkbook(is);
            } else if (fileName.endsWith(xlsx)) {
                //2007
                workbook = new XSSFWorkbook(is);
            }
            is.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return workbook;
    }

    public static String getCellValue(Cell cell) {
        String cellValue = "";
        if (cell == null) {
            return cellValue;
        }
        //把数字当成String来读，避免出现1读成1.0的情况
        if (cell.getCellType() == CellType.NUMERIC) {
            cell.setCellType(CellType.STRING);
        }
        //判断数据的类型
        switch (cell.getCellType()) {
            case NUMERIC:
                //数字
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case STRING:
                //字符串
                cellValue = String.valueOf(cell.getStringCellValue());
                break;
            case BOOLEAN:
                //Boolean
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case FORMULA:
                //公式
                //cellValue = String.valueOf(cell.getCellFormula());
                cellValue = String.valueOf(cell.getStringCellValue());
                break;
            case BLANK:
                //空值
                cellValue = "";
                break;
            case ERROR:
                //故障
                cellValue = "非法字符";
                break;
            default:
                cellValue = "未知类型";
                break;
        }
        return cellValue;
    }

    /**
     * 设置边框
     *
     * @param sheet
     */
    public static void setBorder(Sheet sheet) {
        CellStyle borderCellStyle = sheet.getWorkbook().createCellStyle();
        borderCellStyle.setBorderTop(BorderStyle.THIN);
        borderCellStyle.setBorderBottom(BorderStyle.THIN);
        borderCellStyle.setBorderLeft(BorderStyle.THIN);
        borderCellStyle.setBorderRight(BorderStyle.THIN);
        //获得当前sheet的开始行
        int firstRowNum = sheet.getFirstRowNum();
        //获得当前sheet的结束行
        int lastRowNum = sheet.getLastRowNum();
        //获取头行
        Row titleRow = sheet.getRow(firstRowNum);
        //获得当前行的开始列
        int firstCellNum = titleRow.getFirstCellNum();
        //获得当前行的列数
        int lastCellNum = titleRow.getLastCellNum();
        for (int rowIndex = firstRowNum; rowIndex <= lastRowNum; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            for (int cellIndex = firstCellNum; cellIndex < lastCellNum; cellIndex++) {
                Cell cell = row.getCell(cellIndex);
                if (cell == null) {
                    cell = row.createCell(cellIndex, titleRow.getCell(cellIndex).getCellType());
                }
                CellStyle cellStyle = cell.getCellStyle();
                if (cellStyle == null || "General".equals(cellStyle.getDataFormatString())) {
                    cellStyle = borderCellStyle;
                } else {
                    cellStyle.setBorderTop(BorderStyle.THIN);
                    cellStyle.setBorderBottom(BorderStyle.THIN);
                    cellStyle.setBorderLeft(BorderStyle.THIN);
                    cellStyle.setBorderRight(BorderStyle.THIN);
                }
                cell.setCellStyle(cellStyle);
            }
        }
    }

}

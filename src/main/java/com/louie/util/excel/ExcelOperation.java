package com.louie.util.excel;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;

/**
 * @author maoliang
 * @date 2018/10/11
 */
public class ExcelOperation {

    //private final static String TITLE_班级编号 = "班级编号";
    //private final static String TITLE_班级名称 = "班级名称";
    //private final static String TITLE_学籍号 = "学籍号";
    //private final static String TITLE_民族代码 = "民族代码";
    //private final static String TITLE_姓名 = "姓名";
    //private final static String TITLE_性别 = "性别";
    //private final static String TITLE_出生日期 = "出生日期";
    //private final static String TITLE_身高 = "身高";
    //private final static String TITLE_体重 = "体重";
    //private final static String TITLE_肺活量 = "肺活量";
    //private final static String TITLE_50米跑 = "50米跑";
    //private final static String TITLE_立定跳远 = "立定跳远";
    //private final static String TITLE_坐位体前屈 = "坐位体前屈";
    //private final static String TITLE_800米跑 = "800米跑";
    //private final static String TITLE_1000米跑 = "1000米跑";
    //private final static String TITLE_一分钟仰卧起坐 = "一分钟仰卧起坐";
    //private final static String TITLE_引体向上 = "引体向上";

    public static void main(String[] args) {
        try {
            File sourceDir = new File("/Users/maoliang/Downloads/temp/old2/");
            File targetDir = new File("/Users/maoliang/Downloads/temp/new/");
            if (!targetDir.exists()) {
                targetDir.mkdirs();
            }
            for (File file : sourceDir.listFiles()) {
                operation(file, targetDir);
            }
            System.out.println();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    public static void operation(File file, File targetDir) throws IOException {
        if (!ExcelUtil.checkFile(file)) {
            return;
        }
        //文件按规则重命名
        int fileIndex = file.getName().indexOf(".");
        String fileName = file.getName().substring(0, fileIndex);
        if (!fileName.contains("班")) {
            fileName = fileName + "班";
        }
        if (fileName.length() < 5) {
            fileName = fileName.substring(0, 2) + "0" + fileName.substring(2, 4);
        }
        File targetFile = new File(targetDir, fileName + file.getName().substring(fileIndex));
        Files.copy(file.toPath(), targetFile.toPath(), StandardCopyOption.REPLACE_EXISTING);

        Workbook workbook = ExcelUtil.getWorkBook(targetFile);
        Sheet sheet = workbook.getSheetAt(0);
        //获取标题行
        Row titleRow = sheet.getRow(sheet.getFirstRowNum());
        if (titleRow.getLastCellNum() > 13) {
            //删除左边多的几列
            for (int i = 0; i < 4; i++) {
                Cell cell = titleRow.getCell(i);
                cell.removeCellComment();
                sheet.setColumnHidden(i, false);
            }
            sheet.shiftColumns(4, titleRow.getLastCellNum(), -4);
        }

        //加边框
        ExcelUtil.setBorder(sheet);

        //设置列宽
        sheet.setColumnWidth(0, 8 * 256);
        sheet.setColumnWidth(1, 5 * 256);
        sheet.setColumnWidth(2, 11 * 256);
        sheet.setColumnWidth(3, 8 * 256);
        sheet.setColumnWidth(4, 8 * 256);
        sheet.setColumnWidth(5, 8 * 256);
        sheet.setColumnWidth(6, 8 * 256);
        sheet.setColumnWidth(7, 8 * 256);
        sheet.setColumnWidth(8, 10 * 256);
        sheet.setColumnWidth(9, 8 * 256);
        sheet.setColumnWidth(10, 9 * 256);
        sheet.setColumnWidth(11, 13 * 256);
        sheet.setColumnWidth(12, 8 * 256);

        //打印设置
        PrintSetup ps = sheet.getPrintSetup();
        ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
        // 打印方向，true：横向，false：纵向(默认)
        ps.setLandscape(true);
        //ps.setVResolution((short)600);

        //保存修改
        ExcelUtil.saveWorkbook(targetFile, workbook);
    }

}

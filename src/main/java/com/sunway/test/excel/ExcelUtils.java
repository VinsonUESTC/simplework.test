/*
 * Copyright (c) 2018. www.sunway.com Edit by Vinson
 */


package com.sunway.test.excel;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class ExcelUtils {

    public static void main(String[] args) throws IOException {
        CopyXlsToXlsx("/Users/vinson/Downloads/科技职业学院3#宿舍楼WLAN交资表.xls",null,false, 0);
    }

    //从xls拷贝到xlsx
    public static void CopyXlsToXlsx(String srcfilepath, String desfilename, boolean withformula, int sheetindex) throws IOException {
        //定义变量
        File excelfile = new File(srcfilepath);
        String savepath = null;
        HSSFCell xlscell = null;
        XSSFCell xlsxcell = null;
        XSSFCellStyle xlsxcellstyle = null;
        HSSFCellStyle xlscellstyle = null;
        Font xlsxcellfont = null;
        Font xlscellfont = null;
        //创建保存路径
        if(desfilename==null){
            savepath = excelfile.getAbsolutePath()+"x";
        }else {
            savepath = excelfile.getParent() + "/" + desfilename;
        }
        //创建工作簿
        HSSFWorkbook xlsworkbook = new HSSFWorkbook(new FileInputStream(excelfile));
        HSSFSheet xlssheet = xlsworkbook.getSheetAt(sheetindex);
        XSSFWorkbook xlsxworkbook = new XSSFWorkbook();
        XSSFSheet xlsxsheet = xlsxworkbook.createSheet(xlssheet.getSheetName());
        //遍历工作表
        for(int i = 0; i < xlssheet.getLastRowNum()+1; i++){
            HSSFRow xlsrow = xlssheet.getRow(i);
            XSSFRow xlsxrow = xlsxsheet.createRow(i);
            for (int j = 0; j < xlsrow.getLastCellNum(); j++){
                xlscell = xlsrow.getCell(j);
                if(xlscell!=null) {
                    xlsxcell = xlsxrow.createCell(j);
                    xlsxcellstyle = xlsxworkbook.createCellStyle();
                    xlscellstyle = xlscell.getCellStyle();
                    xlsxcellfont = xlsxworkbook.createFont();
                    xlscellfont = xlscellstyle.getFont(xlsworkbook);
                    FormatCellStyle(xlscellstyle, xlsxcellstyle);
                    FormatCellFont(xlscellfont, xlsxcellfont);
                    xlsxcellstyle.setFont(xlsxcellfont);
                    xlsxcell.setCellStyle(xlsxcellstyle);
                    if (withformula) {
                        CopyCellWithFormula(xlscell, xlsxcell);
                    } else {
                        CopyCellWithoutFormula(xlscell, xlsxcell, xlsworkbook);
                    }
                }
            }
        }
        for (int i = 0; i < xlsxsheet.getRow(xlsxsheet.getTopRow()).getLastCellNum(); i++){
            xlsxsheet.autoSizeColumn(i);
        }
        SaveXlsxWorkbook(xlsxworkbook,savepath);
    }

    //从xls拷贝到xlsx（扩展）
    public static void CopyXlsToXlsx(String srcfilepath, String desfilename, boolean withformula) throws IOException {
        //定义变量
        File excelfile = new File(srcfilepath);
        String savepath = excelfile.getParent()+"/"+desfilename;
        ArrayList<HSSFSheet> xlssheetlist = new ArrayList<>();
        HSSFRow xlsrow = null;
        XSSFRow xlsxrow = null;
        HSSFCell xlscell = null;
        XSSFCell xlsxcell = null;
        XSSFCellStyle xlsxcellstyle = null;
        HSSFCellStyle xlscellstyle = null;
        Font xlsxcellfont = null;
        Font xlscellfont = null;
        //创建工作簿
        HSSFWorkbook xlsworkbook = new HSSFWorkbook(new FileInputStream(excelfile));
        XSSFWorkbook xlsxworkbook = new XSSFWorkbook();
        //遍历工作簿
        xlssheetlist = ListSheets(xlsworkbook);
        for(HSSFSheet xlssheet : xlssheetlist){
            XSSFSheet xlsxsheet = xlsxworkbook.createSheet(xlssheet.getSheetName());
            //遍历工作表
            for(int i = 0; i < xlssheet.getLastRowNum()+1; i++){
                xlsrow = xlssheet.getRow(i);
                xlsxrow = xlsxsheet.createRow(i);
                for (int j = 0; j < xlsrow.getLastCellNum(); j++) {
                    xlscell = xlsrow.getCell(j);
                    if (xlscell != null) {
                        xlsxcell = xlsxrow.createCell(j);
                        xlsxcellstyle = xlsxworkbook.createCellStyle();
                        xlscellstyle = xlscell.getCellStyle();
                        xlsxcellfont = xlsxworkbook.createFont();
                        xlscellfont = xlscellstyle.getFont(xlsworkbook);
                        FormatCellStyle(xlscellstyle, xlsxcellstyle);
                        FormatCellFont(xlscellfont, xlsxcellfont);
                        xlsxcellstyle.setFont(xlsxcellfont);
                        xlsxcell.setCellStyle(xlsxcellstyle);
                        if (withformula) {
                            CopyCellWithFormula(xlscell, xlsxcell);
                        } else {
                            CopyCellWithoutFormula(xlscell, xlsxcell, xlsworkbook);
                        }
                    }
                }
            }
            for (int i = 0; i < xlsxsheet.getRow(xlsxsheet.getTopRow()).getLastCellNum(); i++){
                xlsxsheet.autoSizeColumn(i);
            }
        }
        SaveXlsxWorkbook(xlsxworkbook,savepath);
    }

    //列出工作簿中所有sheets
    public static ArrayList<HSSFSheet> ListSheets(HSSFWorkbook xlsworkbook){
        ArrayList<HSSFSheet> xlssheetlist = new ArrayList<>();
        for(int i = 0; i < xlsworkbook.getNumberOfSheets(); i++){
            xlssheetlist.add(xlsworkbook.getSheetAt(i));
        }
        return xlssheetlist;
    }

    //列出工作簿中所有sheets
    public static ArrayList<XSSFSheet> ListSheets(XSSFWorkbook xlsxworkbook){
        ArrayList<XSSFSheet> xlssheetlist = new ArrayList<>();
        for(int i = 0; i < xlsxworkbook.getNumberOfSheets(); i++){
            xlssheetlist.add(xlsxworkbook.getSheetAt(i));
        }
        return xlssheetlist;
    }

    //拷贝单元格值（含公式）
    public static void CopyCellWithFormula(HSSFCell xlscell, XSSFCell xlsxcell){
        switch (xlscell.getCellTypeEnum()){
            case _NONE:
                break;
            case BLANK:
                break;
            case ERROR:
                xlsxcell.setCellErrorValue(xlscell.getErrorCellValue());
                break;
            case STRING:
                xlsxcell.setCellValue(xlscell.getStringCellValue());
                break;
            case BOOLEAN:
                xlsxcell.setCellValue(xlscell.getBooleanCellValue());
                break;
            case FORMULA:
                xlsxcell.setCellFormula(xlscell.getCellFormula());
                break;
            case NUMERIC:
                xlsxcell.setCellValue(xlscell.getNumericCellValue());
                break;
        }
    }

    //拷贝单元格值（不含公式）
    public static void CopyCellWithoutFormula(HSSFCell xlscell, XSSFCell xlsxcell, HSSFWorkbook xlsworkbook){
        HSSFFormulaEvaluator xlsfme = HSSFFormulaEvaluator.create(xlsworkbook,null,null);
        CellValue xlscellvalue = xlsfme.evaluate(xlscell);
        if (xlscell.getCellTypeEnum().equals(CellType.FORMULA)){
            switch (xlsfme.evaluateFormulaCellEnum(xlscell)){
                case _NONE:
                    break;
                case BLANK:
                    break;
                case ERROR:
                    xlsxcell.setCellErrorValue(xlscellvalue.getErrorValue());
                    break;
                case STRING:
                    xlsxcell.setCellValue(xlscellvalue.getStringValue());
                    break;
                case BOOLEAN:
                    xlsxcell.setCellValue(xlscellvalue.getBooleanValue());
                    break;
                case NUMERIC:
                    xlsxcell.setCellValue(xlscellvalue.getNumberValue());
                    break;
            }
        }else {
            switch (xlscell.getCellTypeEnum()) {
                case _NONE:
                    break;
                case BLANK:
                    break;
                case ERROR:
                    xlsxcell.setCellErrorValue(xlscell.getErrorCellValue());
                    break;
                case STRING:
                    xlsxcell.setCellValue(xlscell.getStringCellValue());
                    break;
                case BOOLEAN:
                    xlsxcell.setCellValue(xlscell.getBooleanCellValue());
                    break;
                case NUMERIC:
                    xlsxcell.setCellValue(xlscell.getNumericCellValue());
                    break;
            }
        }
    }

    //保存工作簿
    public static void SaveXlsxWorkbook(XSSFWorkbook xlsxworkbook, String savepath) throws IOException {
        FileOutputStream fo = new FileOutputStream(savepath);
        xlsxworkbook.write(fo);
        fo.close();
    }

    //格式化单元格
    public static void FormatCellStyle(HSSFCellStyle hssfCellStyle, XSSFCellStyle xssfCellStyle){
        xssfCellStyle.setAlignment(hssfCellStyle.getAlignmentEnum());
        xssfCellStyle.setBorderBottom(hssfCellStyle.getBorderBottomEnum());
        xssfCellStyle.setTopBorderColor(hssfCellStyle.getTopBorderColor());
        xssfCellStyle.setBottomBorderColor(hssfCellStyle.getBottomBorderColor());
        xssfCellStyle.setLeftBorderColor(hssfCellStyle.getLeftBorderColor());
        xssfCellStyle.setRightBorderColor(hssfCellStyle.getRightBorderColor());
        xssfCellStyle.setBorderBottom(hssfCellStyle.getBorderBottomEnum());
        xssfCellStyle.setBorderRight(hssfCellStyle.getBorderRightEnum());
        xssfCellStyle.setBorderTop(hssfCellStyle.getBorderTopEnum());
        xssfCellStyle.setBorderLeft(hssfCellStyle.getBorderLeftEnum());
        xssfCellStyle.setDataFormat(hssfCellStyle.getDataFormat());
        xssfCellStyle.setFillBackgroundColor(hssfCellStyle.getFillBackgroundColor());
        xssfCellStyle.setFillForegroundColor(hssfCellStyle.getFillForegroundColor());
        xssfCellStyle.setFillPattern(hssfCellStyle.getFillPatternEnum());
        xssfCellStyle.setHidden(hssfCellStyle.getHidden());
        xssfCellStyle.setIndention(hssfCellStyle.getIndention());
        xssfCellStyle.setLocked(hssfCellStyle.getLocked());
        xssfCellStyle.setQuotePrefixed(hssfCellStyle.getQuotePrefixed());
        xssfCellStyle.setRotation(hssfCellStyle.getRotation());
        xssfCellStyle.setShrinkToFit(hssfCellStyle.getShrinkToFit());
        xssfCellStyle.setVerticalAlignment(hssfCellStyle.getVerticalAlignmentEnum());
        xssfCellStyle.setWrapText(hssfCellStyle.getWrapText());
    }

    //格式化字体
    public static void FormatCellFont(Font srcfont, Font desfont){
        desfont.setBold(srcfont.getBold());
        desfont.setCharSet(srcfont.getCharSet());
        desfont.setColor(srcfont.getColor());
        desfont.setFontHeight(srcfont.getFontHeight());
        desfont.setItalic(srcfont.getItalic());
        desfont.setFontName(srcfont.getFontName());
        desfont.setUnderline(srcfont.getUnderline());
        desfont.setStrikeout(srcfont.getStrikeout());
        desfont.setTypeOffset(srcfont.getTypeOffset());
    }
}

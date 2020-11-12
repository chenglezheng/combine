package com.lc.combine.service.impl;

import com.lc.combine.service.CombineExcelService;
import com.lc.combine.util.DateUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;


/**
 * @Author chenglezheng
 * @Date 2020/11/11 16:11
 */

@Service
public class CombineExcelServiceImpl implements CombineExcelService {

    /**
     * 合并所有当前文件下含有“向阳奶站-客户报数”的文件
     * @throws Exception
     */
    @Override
    public void combinne() throws Exception{
        /*String Path = new File("").getAbsolutePath();*/
        String Path = "D:\\testcombine";
        File file = new File(Path);
        File[] tempList = file.listFiles();
        List<String> fileName=new ArrayList<>();
        String fileNamePri="";
        String year="2020";
        String month="01";
        for (int i = 0; i < tempList.length; i++) {
            if (tempList[i].isFile() && tempList[i].toString().contains("客户报数")) {
                fileName.add(tempList[i].toString());
                if(i==10){
                    String[] strings=tempList[i].toString().split("-");
                    fileNamePri=strings[0];
                    String [] temp=fileNamePri.split("\\\\");
                    fileNamePri=temp[temp.length-1];
                    year=strings[2];
                    month=strings[3];
                }
            }
        }
        System.out.println("当前目录下有"+fileName.size()+"个文件需要合并!");
        HSSFWorkbook newExcel = new HSSFWorkbook();
        int i=0;
        int fileSuccessCount=0;
        for (Object fromExcelName : fileName.toArray()) {
            try {
                InputStream in = new FileInputStream(fromExcelName.toString());
                HSSFWorkbook fromExcel = new HSSFWorkbook(in);
                int length = fromExcel.getNumberOfSheets();
                if(length<=1){
                    HSSFSheet oldSheet = fromExcel.getSheetAt(0);
                    if(i==0){
                        String[] strings=fromExcelName.toString().toString().split("-");
                        HSSFSheet newSheet = newExcel.createSheet(strings[3]+"月"+strings[4].split("[.]")[0]+"日");
                        copySheet(newExcel, oldSheet, newSheet,fileNamePri);
                        System.out.println(fromExcelName.toString()+"已合并!");
                        i++;
                    }else if(i==(fileName.size()-1)){
                        String[] strings=fromExcelName.toString().toString().split("-");
                        HSSFSheet newSheet = newExcel.createSheet(strings[3]+"月"+strings[4].split("[.]")[0]+"日");
                        copySheet(newExcel, oldSheet, newSheet,fileNamePri);
                        System.out.println(fromExcelName.toString()+"已合并!");
                        i++;
                    }else {
                        HSSFSheet newSheet = newExcel.createSheet(i+++"");
                        copySheet(newExcel, oldSheet, newSheet,fileNamePri);
                        System.out.println(fromExcelName.toString()+"已合并!");
                    }
                    fileSuccessCount++;
                }else {
                    System.out.println(fromExcelName.toString()+"文件中存在多个Sheet页，不符合合并规则!");
                }
            }catch (Exception e){
                HSSFSheet newSheet = newExcel.createSheet(i+++"");
                e.printStackTrace();
                System.out.println(fromExcelName.toString()+"有未知错误，合并失败！");
            }
        }
        //定义新生成的xlx表格文件
        String allFileName = Path+ "\\"+year+"."+month+fileNamePri+"报数及调整.xls";
        FileOutputStream fileOut = new FileOutputStream(allFileName);
        newExcel.write(fileOut);
        fileOut.flush();
        fileOut.close();
        System.out.println("合并成功文件数："+fileSuccessCount+"\t 失败文件数："+(fileName.size()-fileSuccessCount));
        System.out.println("合并文件名称为:"+year+"."+month+fileNamePri+"报数及调整.xls"+"\t 目录为:"+Path);
    }

    public static void setNewCellStyle(HSSFCellStyle newStyle,HSSFWorkbook wb) {
        //设置样式
        newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        newStyle.setAlignment(HorizontalAlignment.CENTER);
        newStyle.setFillForegroundColor((short)9);
        newStyle.setBorderBottom(BorderStyle.THIN);
        newStyle.setBorderLeft(BorderStyle.THIN);
        newStyle.setBorderTop(BorderStyle.THIN);
        newStyle.setBorderRight(BorderStyle.THIN);

    }

    /**
     * 合并单元格
     * @param fromSheet
     * @param toSheet
     */
    public static void mergeSheetAllRegion(HSSFSheet fromSheet, HSSFSheet toSheet) {
        int num = fromSheet.getNumMergedRegions();
        CellRangeAddress cellR = null;
        for (int i = 0; i < num; i++) {
            cellR = fromSheet.getMergedRegion(i);
            toSheet.addMergedRegion(cellR);
        }
    }

    /**
     * 复制单元格
     * @param wb
     * @param fromCell
     * @param toCell
     */
    public static void copyCell(HSSFWorkbook wb, HSSFCell fromCell, HSSFCell toCell,HSSFCellStyle ceelStyle) {
        toCell.setCellStyle(ceelStyle);
        if (fromCell.getCellComment() != null) {
            toCell.setCellComment(fromCell.getCellComment());
        }
        // 不同数据类型处理
        int fromCellType = fromCell.getCellType();
        toCell.setCellType(fromCellType);
        if (fromCellType == HSSFCell.CELL_TYPE_NUMERIC) {
            if (HSSFDateUtil.isCellDateFormatted(fromCell)) {
                toCell.setCellValue(fromCell.getDateCellValue());
            } else {
                toCell.setCellValue(fromCell.getNumericCellValue());
            }
        } else if (fromCellType == HSSFCell.CELL_TYPE_STRING) {
            toCell.setCellValue(fromCell.getRichStringCellValue());
        } else if (fromCellType == HSSFCell.CELL_TYPE_BLANK) {
            // nothing21
        } else if (fromCellType == HSSFCell.CELL_TYPE_BOOLEAN) {
            toCell.setCellValue(fromCell.getBooleanCellValue());
        } else if (fromCellType == HSSFCell.CELL_TYPE_ERROR) {
            toCell.setCellErrorValue(fromCell.getErrorCellValue());
        } else if (fromCellType == HSSFCell.CELL_TYPE_FORMULA) {
            toCell.setCellFormula(fromCell.getCellFormula());
        }

    }

    /**
     * 行复制功能
     * @param wb
     * @param oldRow
     * @param toRow
     */
    public static void copyRow(HSSFWorkbook wb, HSSFRow oldRow, HSSFRow toRow,HSSFCellStyle newStyle) {
        toRow.setHeight(oldRow.getHeight());
        int i=0;
        for (Iterator cellIt = oldRow.cellIterator(); cellIt.hasNext();) {
            HSSFCell tmpCell = (HSSFCell) cellIt.next();
            if(i==0){
                HSSFCell newCell = toRow.createCell(i);
                //处理第一列
                updateDateFormat(tmpCell);
                copyCell(wb, tmpCell, newCell,newStyle);
                i++;
            }else if(i==4){
                i++;
            }else {
                if(i>4){
                    int j=i;
                    HSSFCell newCell = toRow.createCell(j-1);
                    copyCell(wb, tmpCell, newCell,newStyle);
                    i++;
                }else {
                    HSSFCell newCell = toRow.createCell(i);
                    copyCell(wb, tmpCell, newCell,newStyle);
                    i++;
                }
            }
        }
    }

    /**
     * 处理日期转换(仅限处理第一行)
     * @param fromCellType
     */
    private static void updateDateFormat(HSSFCell fromCell){
        try{
            int fromCellType = fromCell.getCellType();
            if (fromCellType == HSSFCell.CELL_TYPE_NUMERIC) {
                Double numericCellValue = fromCell.getNumericCellValue();
                fromCell.setCellValue(DateUtils.handleDate(numericCellValue.longValue()));
            } else if (fromCellType == HSSFCell.CELL_TYPE_STRING) {
                fromCell.setCellValue(fromCell.getRichStringCellValue());
            } else if (fromCellType == HSSFCell.CELL_TYPE_BLANK) {
                // nothing21
            } else if (fromCellType == HSSFCell.CELL_TYPE_BOOLEAN) {
                fromCell.setCellValue(fromCell.getBooleanCellValue());
            } else if (fromCellType == HSSFCell.CELL_TYPE_ERROR) {
                fromCell.setCellErrorValue(fromCell.getErrorCellValue());
            } else if (fromCellType == HSSFCell.CELL_TYPE_FORMULA) {
                fromCell.setCellFormula(fromCell.getCellFormula());
            }
        }catch(Exception e){
            System.out.println("第一列日期处理错误！");
            e.printStackTrace();
        }
    }



    /**
     * Sheet复制
     * @param wb
     * @param fromSheet
     * @param toSheet
     */
    public static void copySheet(HSSFWorkbook wb, HSSFSheet fromSheet, HSSFSheet toSheet,String stationName) {
        mergeSheetAllRegion(fromSheet, toSheet);
        HSSFCellStyle newStyle = wb.createCellStyle();
        setNewCellStyle(newStyle,wb);
        int length = 200;
        for (int i = 0; i <= length; i++) {
            toSheet.setColumnWidth(i, 3800);
        }
        int i=0;
        for (Iterator rowIt = fromSheet.rowIterator(); rowIt.hasNext();) {
            if(i==0){
                HSSFRow newRow = toSheet.createRow(0);
                HSSFCell newCell = newRow.createCell(i);
                newCell.setCellStyle(newStyle);
                newCell.setCellValue(stationName+"正常报数单");
                i++;
            }else {
                HSSFRow oldRow = (HSSFRow) rowIt.next();
                HSSFRow newRow = toSheet.createRow(oldRow.getRowNum());
                copyRow(wb, oldRow, newRow,newStyle);
            }
        }
        toSheet.addMergedRegion(new CellRangeAddress(0,0,0,8));
    }

    public class HSSFDateUtil extends DateUtil {

    }


}

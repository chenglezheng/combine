package com.lc.combine.service.impl;

import com.lc.combine.service.CombineExcelService;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
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
        String Path = new File("").getAbsolutePath();
        /*String Path = "D:\\testcombine";*/
        File file = new File(Path);
        File[] tempList = file.listFiles();
        List<String> fileName=new ArrayList<>();
        String year="";
        String month="";
        String fileNamePri="";
        for (int i = 0; i < tempList.length; i++) {
            if (tempList[i].isFile() && tempList[i].toString().contains("商户报数")) {
                String[] strings=tempList[i].toString().split("\\\\");
                String tempFileNamePri=strings[strings.length-1];
                fileName.add(tempFileNamePri);
                String[] temp=tempFileNamePri.split("-");
                fileNamePri=temp[0].substring(0,temp[0].length()-2);
                year=temp[2]+".";
                month=temp[3];
            }
        }
        System.out.println("当前目录下有"+fileName.size()+"个文件需要合并!");
        HSSFWorkbook newExcel = new HSSFWorkbook();
        int i=1;
        int fileSuccessCount=0;
        for (Object fromExcelName : fileName.toArray()) {
            String fromExcelNamePath=Path+"\\"+fromExcelName.toString();
            try {
                InputStream in = new FileInputStream(fromExcelNamePath);
                HSSFWorkbook fromExcel = new HSSFWorkbook(in);
                int length = fromExcel.getNumberOfSheets();
                if(length<=1){
                    HSSFSheet oldSheet = fromExcel.getSheetAt(0);
                    HSSFSheet newSheet = newExcel.createSheet(i+++"");
                    copySheet(newExcel, oldSheet, newSheet);
                    System.out.println(fromExcelNamePath+"已合并!");
                    fileSuccessCount++;
                }else {
                    System.out.println(fromExcelNamePath+"文件中存在多个Sheet页，不符合合并规则!");
                }
            }catch (Exception e){
                HSSFSheet newSheet = newExcel.createSheet(i+++"");
                e.printStackTrace();
                System.out.println(fromExcelNamePath+"有未知错误，合并失败！");
            }
        }
        //定义新生成的xlx表格文件
        String allFileName = Path+ "\\"+year+month+fileNamePri+"每日线下签收单.xls";
        FileOutputStream fileOut = new FileOutputStream(allFileName);
        newExcel.write(fileOut);
        fileOut.flush();
        fileOut.close();
        System.out.println("合并成功文件数："+fileSuccessCount+"\t 失败文件数："+(fileName.size()-fileSuccessCount));
        System.out.println("合并文件名称为:"+year+month+fileNamePri+"每日线下签收单.xls"+"\t 目录为:"+Path);
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
    public static void copySpecialRow(HSSFWorkbook wb, HSSFRow oldRow, HSSFRow toRow,HSSFCellStyle newStyle) {
        toRow.setHeight(oldRow.getHeight());
        for (int i = 0; i <oldRow.getPhysicalNumberOfCells();i++) {
            HSSFCell newCell = toRow.createCell(i);
            newCell.setCellStyle(newStyle);
        }
        int j=2;
        for (Iterator cellIt = oldRow.cellIterator(); cellIt.hasNext();) {
            HSSFCell tmpCell = (HSSFCell) cellIt.next();
            HSSFCell newCell = toRow.createCell(j);
            copyCell(wb, tmpCell, newCell,newStyle);
            j++;
        }
    }

    /**
     * 行复制功能
     * @param wb
     * @param oldRow
     * @param toRow
     */
    public static void copySpecialRow1(HSSFWorkbook wb, HSSFRow oldRow, HSSFRow toRow,HSSFCellStyle newStyle) {
        toRow.setHeight(oldRow.getHeight());
        for (int i = 0; i <oldRow.getPhysicalNumberOfCells() ; i++) {
            HSSFCell newCell = toRow.createCell(i);
            newCell.setCellStyle(newStyle);
        }
        int j=0;
        for (Iterator cellIt = oldRow.cellIterator(); cellIt.hasNext();) {
            HSSFCell tmpCell = (HSSFCell) cellIt.next();
            HSSFCell newCell = toRow.createCell(j);
            copyCell(wb, tmpCell, newCell,newStyle);
            j++;
        }
    }


    /**
     * 行复制功能
     * @param wb
     * @param oldRow
     * @param toRow
     */
    public static void copySpecial1Row(HSSFWorkbook wb, HSSFRow oldRow, HSSFRow toRow,HSSFCellStyle newStyle) {
        toRow.setHeight(oldRow.getHeight());
        for (int i = 0; i <oldRow.getPhysicalNumberOfCells() ; i++) {
            HSSFCell newCell = toRow.createCell(i);
            newCell.setCellStyle(newStyle);
        }
        int j=1;
        for (Iterator cellIt = oldRow.cellIterator(); cellIt.hasNext();) {
            if(j==1){
                HSSFCell tmpCell = (HSSFCell) cellIt.next();
            }else{
                HSSFCell tmpCell = (HSSFCell) cellIt.next();
                HSSFCell newCell = toRow.createCell(j);
                copyCell(wb, tmpCell, newCell,newStyle);
            }
            j++;
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
        for (int i = 0; i <oldRow.getPhysicalNumberOfCells() ; i++) {
            HSSFCell newCell = toRow.createCell(i);
            newCell.setCellStyle(newStyle);
        }
        int j=0;
        for (Iterator cellIt = oldRow.cellIterator(); cellIt.hasNext();) {
            if(j==0){
                HSSFCell tmpCell = (HSSFCell) cellIt.next();
                j++;
            }else{
                HSSFCell tmpCell = (HSSFCell) cellIt.next();
                HSSFCell newCell = toRow.createCell(j-1);
                copyCell(wb, tmpCell, newCell,newStyle);
                j++;
            }
        }
    }


    /**
     * Sheet复制
     * @param wb
     * @param fromSheet
     * @param toSheet
     */
    public static void copySheet(HSSFWorkbook wb, HSSFSheet fromSheet, HSSFSheet toSheet) {
        mergeSheetAllRegion(fromSheet, toSheet);
        HSSFCellStyle newStyle = wb.createCellStyle();
        setNewCellStyle(newStyle,wb);
        int length = 200;
        for (int i = 0; i <= length; i++) {
            toSheet.setColumnWidth(i, 3800);
        }
        int i=0;
        HSSFRow tempRow = null;
        for (Iterator rowIt = fromSheet.rowIterator(); rowIt.hasNext();) {
            if(i==0){
                HSSFRow oldRow = (HSSFRow) rowIt.next();
                HSSFRow newRow = toSheet.createRow(i);
                /*for (int k = 0; k <23 ; k++) {
                    HSSFCell newCell = newRow.createCell(k);
                    newCell.setCellStyle(newStyle);
                }*/
                i++;
            }if(i==1 || i==2){
                HSSFRow oldRow = (HSSFRow) rowIt.next();
                HSSFRow newRow = toSheet.createRow(i);
                copySpecialRow(wb, oldRow, newRow,newStyle);
                i++;
            }if(i==3){
                HSSFRow oldRow = (HSSFRow) rowIt.next();
                tempRow=oldRow;
                HSSFRow newRow = toSheet.createRow(i);
                if(oldRow.getPhysicalNumberOfCells()==36){
                    copySpecialRow1(wb, oldRow, newRow,newStyle);
                }else {
                    copyRow(wb, oldRow, newRow,newStyle);
                }
                i++;
            }else{
                HSSFRow oldRow = (HSSFRow) rowIt.next();
                tempRow=oldRow;
                HSSFRow newRow = toSheet.createRow(i);
                copyRow(wb, oldRow, newRow,newStyle);
                i++;
            }
        }
        //复写最后一行
        HSSFRow newRow = toSheet.createRow(i-1);
        copySpecial1Row(wb, tempRow, newRow,newStyle);
        newRow = toSheet.createRow(i);
        for (int k = 0; k <tempRow.getPhysicalNumberOfCells() ; k++) {
            HSSFCell newCell = newRow.createCell(k);
            newCell.setCellStyle(newStyle);
            if (k==0){
                newCell.setCellValue("签字");
            }
        }
        toSheet.addMergedRegion(new CellRangeAddress(1,2,0,1));
        toSheet.addMergedRegion(new CellRangeAddress(i,i,0,1));
    }

    public class HSSFDateUtil extends DateUtil {

    }


}

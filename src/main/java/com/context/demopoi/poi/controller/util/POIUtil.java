package com.context.demopoi.poi.controller.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.stream.Stream;

@Component
@Slf4j
public class POIUtil {

    /**
     * 获取Excel workBook
     * @param inputStream
     * @return
     */
    public HSSFWorkbook getWorkBook(InputStream inputStream){
        if(inputStream==null){
            log.info("getWorkBook,传入参数inputStream=null");
            return null;
        }
        try {
            return new HSSFWorkbook(new POIFSFileSystem(inputStream));
        } catch (IOException e) {
            e.printStackTrace();
            log.error("POIUtil.getWorkBook()异常:{}",e.getMessage());
        }
        return null;
    }

    public HSSFSheet getSheet(HSSFWorkbook workbook,int index){
        if(workbook==null){
            log.info("getSheet,传入参数workbook=null");
            return null;
        }
        return workbook.getSheetAt(index);
    }

    public List<String> getCellValueList(HSSFRow row){
        if(row==null){
            log.info("getCellValueList,row=null");
            return null;
        }
        List<String> cellValueList = new ArrayList<>();
        int cellNum = row.getPhysicalNumberOfCells();
        for (int i1 = 0; i1 < cellNum; i1++) {
            cellValueList.add(getCellValue(row.getCell(i1)));
        }
        return cellValueList;
    }

    public List<HSSFRow> getRowList(HSSFSheet hssfSheet){
        if(hssfSheet==null){
            log.info("getRowList,hssfSheet=null");
            return null;
        }
        List<HSSFRow> rowList = new ArrayList<>();
        int rowNum = hssfSheet.getPhysicalNumberOfRows();
        for (int i = 0; i < rowNum; i++) {
            rowList.add(hssfSheet.getRow(i));
        }
        return rowList;
    }

    /***
     *
     * @param hssfCell
     */
    private String getCellValue(HSSFCell hssfCell){
        if(hssfCell==null){
            return null;
        }
        String cellValue = null;
        switch (hssfCell.getCellTypeEnum()){
            case STRING:
                cellValue = hssfCell.getStringCellValue();
                break;
            case NUMERIC:
                cellValue = String.valueOf(hssfCell.getNumericCellValue());
                break;
            case BOOLEAN:
                cellValue = String.valueOf(hssfCell.getBooleanCellValue());
                break;
            default:
                cellValue = "";
        }
        return cellValue;
    }


}

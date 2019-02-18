package com.context.demopoi.poi.controller;

import com.context.demopoi.poi.controller.util.POIUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

@RestController
@Slf4j
public class POIDemoController {

    @Autowired
    private POIUtil poiUtil;

    @PostMapping(path = "/importExcel")
    public String importExcel(MultipartFile multipartFile){
        String fileName = multipartFile.getOriginalFilename();
        System.out.println("fileName="+fileName);
        try {
            HSSFWorkbook workbook = poiUtil.getWorkBook(multipartFile.getInputStream());
            HSSFSheet sheet = poiUtil.getSheet(workbook,0);
            List<HSSFRow> rowList = poiUtil.getRowList(sheet);
            rowList.stream().forEach(row -> poiUtil.getCellValueList(row).stream().forEach(System.out::println));
        } catch (IOException e) {
            e.printStackTrace();
        }
        return "ok";
    }

    @GetMapping(path = "/importTest")
    public String imortTest(){
        return "test";
    }

}

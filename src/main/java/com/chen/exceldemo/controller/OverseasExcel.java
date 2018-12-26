package com.chen.exceldemo.controller;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.Sheet;
import com.chen.exceldemo.domain.OverseasExcelEntity;
import com.google.common.collect.Lists;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

@RestController
@RequestMapping("overseasExcel")
@Slf4j
public class OverseasExcel {

    @PostMapping("/overseasImport")
    public String overseasImport(@RequestParam(value = "file", required = false) MultipartFile file) {
        InputStream inputStream = null;
        try {
            inputStream = file.getInputStream();
            List<Object> read = EasyExcelFactory.read(inputStream, new Sheet(1, 1, OverseasExcelEntity.class));
            ArrayList<OverseasExcelEntity> overseasExcels = Lists.newArrayList();
            read.forEach(r->overseasExcels.add((OverseasExcelEntity)r));

            if (overseasExcels.size() > 0) {
                for (int i = 0; i < overseasExcels.size(); i++) {
                    log.info("i: {}",i);
                    if (overseasExcels.get(i).getCountryArea() == null){
                        String countryArea = overseasExcels.get(i - 1).getCountryArea();
                        overseasExcels.get(i).setCountryArea(countryArea);
                    }
                    log.info("data: {}",overseasExcels.get(i));
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        return "";
    }


    @PostMapping("/overseasExport")
    public String overseasExport() {
        ArrayList<OverseasExcelEntity> overseasExcel = Lists.newArrayList();
        for (int i = 0; i < 10; i++) {
            OverseasExcelEntity overseasExcelEntity = new OverseasExcelEntity();
            overseasExcelEntity.setCountryArea("印度");
            overseasExcelEntity.setCity("印尼"+i);
            overseasExcelEntity.setScene("场景"+i);
            overseasExcel.add(overseasExcelEntity);
        }
        OutputStream out = null;
        try {
            out = new FileOutputStream("/Users/liliang/Documents/9685.xlsx");
            ExcelWriter writer = EasyExcelFactory.getWriter(out);
            Sheet sheet = new Sheet(1, 1);

            sheet.setSheetName("第一个Sheet");
            sheet.setAutoWidth(Boolean.TRUE);

            writer.write(overseasExcel,sheet);
            writer.merge(1,8,1,1);

            writer.finish();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }


        return "";
    }
}


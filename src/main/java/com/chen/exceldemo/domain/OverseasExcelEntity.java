package com.chen.exceldemo.domain;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;
import lombok.Data;

@Data
public class OverseasExcelEntity extends BaseRowModel {
    @ExcelProperty(value = "表头1",index = 0)
    private String countryArea;
    @ExcelProperty(value = "表头2",index = 1)
    private String city;
    @ExcelProperty(value = "表头3",index = 2)
    private String Scene;
}

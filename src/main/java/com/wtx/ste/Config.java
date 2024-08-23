package com.wtx.ste;

import lombok.Data;

@Data
public class Config {
    private String swaggerJsonPath;
    private String excelPath;
    private ColumnConfig columnConfig;
}

package com.wtx.ste;

import io.swagger.v3.oas.models.OpenAPI;
import io.swagger.v3.oas.models.Operation;
import io.swagger.v3.oas.models.PathItem;
import io.swagger.v3.oas.models.Paths;
import io.swagger.v3.parser.OpenAPIV3Parser;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.yaml.snakeyaml.Yaml;

import java.io.*;
import java.util.Map;

public class App {

    private final static Workbook WORK_BOOK = new XSSFWorkbook();
    private static Config CONFIG = new Config();
    private final static String CONFIG_PATH = "/config.yml";

    private static void initConfig() {
        Yaml yaml = new Yaml();
        InputStream inputSteam = App.class.getResourceAsStream(CONFIG_PATH);
        CONFIG = yaml.loadAs(inputSteam, Config.class);
    }

    public static void main(String[] args) {

        //初始化配置
        initConfig();

        // 解析Swagger JSON
        OpenAPI openAPI = new OpenAPIV3Parser().read(CONFIG.getSwaggerJsonPath());

        // 创建Excel工作簿
        Sheet sheet = WORK_BOOK.createSheet("接口清单");


        // 创建表头
        Row headerRow = sheet.createRow(0);

        createCommonStyleCell(headerRow, CONFIG.getColumnConfig().getPathColumnNum()).setCellValue("Path");
        createCommonStyleCell(headerRow, CONFIG.getColumnConfig().getMethodColumnNum()).setCellValue("Method");
        createCommonStyleCell(headerRow, CONFIG.getColumnConfig().getSummaryColumnNum()).setCellValue("Summary");
        // 处理paths
        Paths paths = openAPI.getPaths();
        for (Map.Entry<String, PathItem> entry : paths.entrySet()) {
            PathItem path = entry.getValue();
            for (Map.Entry<PathItem.HttpMethod, io.swagger.v3.oas.models.Operation> operationEntry : path.readOperationsMap().entrySet()) {
                Row row = sheet.createRow(sheet.getLastRowNum() + 1);

                createCommonStyleCell(row, CONFIG.getColumnConfig().getPathColumnNum()).setCellValue(entry.getKey());
                createCommonStyleCell(row, CONFIG.getColumnConfig().getMethodColumnNum()).setCellValue(operationEntry.getKey().toString());
                Operation operation = operationEntry.getValue();
                if (operation.getSummary() != null) {
                    createCommonStyleCell(row, CONFIG.getColumnConfig().getSummaryColumnNum()).setCellValue(operation.getSummary());
                }
            }
        }
        sheet.autoSizeColumn(CONFIG.getColumnConfig().getPathColumnNum());
        sheet.autoSizeColumn(CONFIG.getColumnConfig().getMethodColumnNum());
        sheet.autoSizeColumn(CONFIG.getColumnConfig().getSummaryColumnNum());
        // 写入文件
        try (FileOutputStream outputStream = new FileOutputStream(CONFIG.getExcelPath())) {
            WORK_BOOK.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                WORK_BOOK.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private static Cell createCommonStyleCell(Row row, int column) {
        Cell cell = row.createCell(column);
        CellStyle cellStyle = WORK_BOOK.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellStyle(cellStyle);
        return cell;
    }
}

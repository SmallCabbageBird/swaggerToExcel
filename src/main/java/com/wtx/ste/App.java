package com.wtx.ste;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import io.swagger.v3.oas.models.OpenAPI;
import io.swagger.v3.oas.models.Operation;
import io.swagger.v3.oas.models.PathItem;
import io.swagger.v3.oas.models.Paths;
import io.swagger.v3.oas.models.media.ArraySchema;
import io.swagger.v3.oas.models.media.Content;
import io.swagger.v3.oas.models.media.MediaType;
import io.swagger.v3.oas.models.media.Schema;
import io.swagger.v3.oas.models.parameters.RequestBody;
import io.swagger.v3.parser.OpenAPIV3Parser;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.yaml.snakeyaml.Yaml;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

public class App {

    private final static ObjectMapper objectMapper = new ObjectMapper();
    private final static Workbook WORK_BOOK = new XSSFWorkbook();
    private static Config CONFIG = new Config();
    private final static String CONFIG_PATH = "/config.yml";

    private static void initConfig() {
        Yaml yaml = new Yaml();
        InputStream inputSteam = App.class.getResourceAsStream(CONFIG_PATH);
        CONFIG = yaml.loadAs(inputSteam, Config.class);
    }

    public static void main(String[] args) throws JsonProcessingException {

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
        createCommonStyleCell(headerRow, CONFIG.getColumnConfig().getReqParamsColumnNum()).setCellValue("RequestParams");
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
                String reqParams = "";
                switch (operationEntry.getKey()) {
                    case POST:
                        reqParams = parsePostReq(operationEntry, openAPI);
                        break;
                }
                createCommonStyleCell(row, CONFIG.getColumnConfig().getReqParamsColumnNum()).setCellValue(reqParams);
            }
        }
        sheet.autoSizeColumn(CONFIG.getColumnConfig().getPathColumnNum());
        sheet.autoSizeColumn(CONFIG.getColumnConfig().getMethodColumnNum());
        sheet.autoSizeColumn(CONFIG.getColumnConfig().getSummaryColumnNum());
        sheet.autoSizeColumn(CONFIG.getColumnConfig().getReqParamsColumnNum());
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

    private static String parsePostReq(Map.Entry<PathItem.HttpMethod, io.swagger.v3.oas.models.Operation> operationEntry, OpenAPI openAPI) throws JsonProcessingException {
        LinkedHashMap<String, Object> paramsMap = new LinkedHashMap<>();
        RequestBody requestBody = operationEntry.getValue().getRequestBody();
        if (requestBody != null) {
            Content content = requestBody.getContent();
            MediaType mediaType = content.get("application/json");
            Schema schema = mediaType.getSchema();
            String ref = schema.get$ref();
            if (ref.startsWith("#/components/schemas/")) {
                parseSchema(schema, openAPI, paramsMap);
                return objectMapper.writeValueAsString(paramsMap);
            }
        }

        return "";
    }

    private static void parseSchema(Schema schema, OpenAPI openAPI, LinkedHashMap<String, Object> paramsMap) {
        String ref = schema.get$ref();
        if (ref != null && ref.startsWith("#/components/schemas/")) {
            parseSchema(openAPI.getComponents().getSchemas().get(ref.split("#/components/schemas/")[1]), openAPI, paramsMap);
        } else {
            Map<String, Schema> properties = schema.getProperties();
            if (properties != null && !properties.isEmpty()) {
                for (Map.Entry<String, Schema> schemaEntry : properties.entrySet()) {
                    String type = schemaEntry.getValue().getType();
                    if ("string".equals(type)) {
                        paramsMap.put(schemaEntry.getKey(), "");
                    } else if ("integer".equals(type)) {
                        paramsMap.put(schemaEntry.getKey(), 0);
                    } else if ("boolean".equals(type)) {
                        paramsMap.put(schemaEntry.getKey(), true);
                    } else if ("long".equals(type)) {
                        paramsMap.put(schemaEntry.getKey(), 0L);
                    } else if ("short".equals(type)) {
                        paramsMap.put(schemaEntry.getKey(), 0);
                    } else if ("double".equals(type)) {
                        paramsMap.put(schemaEntry.getKey(), 0.0);
                    } else if ("float".equals(type)) {
                        paramsMap.put(schemaEntry.getKey(), 0.0F);
                    } else if ("char".equals(type)) {
                        paramsMap.put(schemaEntry.getKey(), null);
                    } else if ("array".equals(type)) {
                        LinkedHashMap<String, Object> subParamsMap = new LinkedHashMap<>();
                        List<LinkedHashMap<String, Object>> newList = new ArrayList();
                        newList.add(subParamsMap);
                        paramsMap.put(schemaEntry.getKey(), newList);
                        ArraySchema arraySchema = (ArraySchema) schemaEntry.getValue();
                        parseSchema(arraySchema.getItems(), openAPI, subParamsMap);
                    } else if ("object".equals(type)) {
                        paramsMap.put(schemaEntry.getKey(), null);
                    } else if (null == type) {
                        LinkedHashMap<String, Object> subParamsMap = new LinkedHashMap<>();
                        paramsMap.put(schemaEntry.getKey(), subParamsMap);
                        parseSchema(schemaEntry.getValue(), openAPI, subParamsMap);

                    } else {
                        System.out.println("未识别的类型:" + type);
                    }
                }
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

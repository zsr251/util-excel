package com.javasoso.util.excel;


import java.io.File;
import java.io.IOException;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * excel工具
 * Created by jasonzhu
 */
public class ExcelUtil {

    private static Logger logger = LoggerFactory.getLogger(ExcelUtil.class);

    /**
     * 验证EXCEL文件
     */
    public static boolean validateExcel(String filePath) {
        return isExcel2003(filePath) || isExcel2007(filePath);
    }

    /**
     * 是否是2003Excel文件
     */
    public static boolean isExcel2003(String filePath) {
        return filePath != null && filePath.matches("^.+\\.(?i)(xls)$");
    }

    /**
     * 是否是2007Excel文件
     */
    public static boolean isExcel2007(String filePath) {
        return filePath != null && filePath.matches("^.+\\.(?i)(xlsx)$");
    }

    /**
     * 获得指定文件 指定Sheet返回值
     *
     * @param file Excel文件
     * @param sheetNum 第几个sheet
     * @param cls 类型
     * @param startRow 第几行开始 0 第一行
     * @param endRow 第几行结束 0/null 默认最后一行  -1 是倒数第一行
     */
    public static <T> List<T> getModelList(File file, int sheetNum, Class<T> cls, Integer startRow,
        Integer endRow) {
        try {
            Workbook workbook = WorkbookFactory.create(file);
            sheetNum = sheetNum > workbook.getNumberOfSheets() ? 0 : sheetNum;

            return ExcelUtil.getModelList(workbook.getSheetAt(sheetNum), cls, startRow, endRow);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        return new ArrayList<T>();
    }

    /**
     * 根据对象列表，获取WorkBook
     *
     * @param modelList 对象列表
     * @param cls 对象类型
     * @param sheetName sheet名字
     * @param startRow 第几行开始 0 第一行
     */
    public static <T> Workbook createWorkBook(List<T> modelList, Class<T> cls, String sheetName,
        Integer startRow) {
        if (sheetName == null || sheetName.length() < 1) {
            sheetName = "sheet";
        }
        startRow = startRow == null || startRow < 0 ? 0 : startRow;
        // 创建文档
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet(sheetName);
        // 需要导出的属性
        Map<Field, ExcelOut> fieldMap = getClassField(cls, ExcelOut.class);
        // 返回空的workbook
        if (fieldMap == null || fieldMap.size() < 1 || modelList == null || modelList.size() < 1) {
            return workbook;
        }
        // 创建标题行
        Row row1 = sheet.createRow(startRow);
        // 定义Cell格式
        CreationHelper creationHelper = workbook.getCreationHelper();
        Map<Field,CellStyle> cellStyleMap = new HashMap<>();
        for (Map.Entry<Field, ExcelOut> fieldExcelCellEntry : fieldMap.entrySet()) {
            Field field = fieldExcelCellEntry.getKey();
            //设置可访问私有属性
            field.setAccessible(true);
            // 设置标题行
            row1.createCell(fieldExcelCellEntry.getValue().value())
                .setCellValue(fieldExcelCellEntry.getValue().name());
            if (field.getType().equals(Date.class)) {
                CellStyle cellStyle = workbook.createCellStyle();
                // 设置日期格式
                cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(fieldExcelCellEntry.getValue().dateFormat()));
                cellStyleMap.put(field,cellStyle);
            }
        }


        // 创建正文
        for (int i = 0; i < modelList.size(); i++) {
            Row rown = sheet.createRow(i + startRow + 1);
            T model = modelList.get(i);
            // 设置列值
            for (Map.Entry<Field, ExcelOut> fieldExcelCellEntry : fieldMap.entrySet()) {
                Field field = fieldExcelCellEntry.getKey();
                //设置可访问私有属性
                field.setAccessible(true);
                // 设置内容行
                Cell cell = rown.createCell(fieldExcelCellEntry.getValue().value());
                try {
                    if (field.getType().equals(Date.class)) {
                        cell.setCellValue((Date) field.get(model));
                        cell.setCellStyle(cellStyleMap.get(field));
                    } else {
                        cell.setCellValue(String.valueOf(field.get(model)));
                    }
                } catch (Exception e) {
                    cell.setCellValue("参数格式化失败");
                }
            }
        }
        return workbook;
    }

    /**
     * 获得指定Sheet 返回值
     *
     * @param sheet 指定sheet
     * @param cls 指定类型
     * @param startRow 开始行 默认0
     * @param endRow 结束行 默认总行数
     */
    private static <T> List<T> getModelList(Sheet sheet, Class<T> cls, Integer startRow,
        Integer endRow) {
        List<T> resultList = new ArrayList<T>();
        //需要赋值的属性
        Map<Field, ExcelIn> fieldMap = getClassField(cls, ExcelIn.class);
        if (fieldMap == null || fieldMap.size() < 1 || sheet == null) {
            return resultList;
        }
        try {
            startRow = startRow == null || startRow < 0 ? 0 : startRow;
            endRow = endRow == null || endRow <= 0 ? sheet.getPhysicalNumberOfRows()
                : endRow <= sheet.getPhysicalNumberOfRows() ? endRow
                    : sheet.getPhysicalNumberOfRows();
            // 如果小于0 则是倒数第几行
            endRow = endRow < 0 ? sheet.getPhysicalNumberOfRows() + endRow + 1 : endRow;
            T t;
            for (int i = startRow; i < endRow; i++) {
                try {
                    //当前行
                    Row row = sheet.getRow(i);
                    t = cls.newInstance();
                    for (Map.Entry<Field, ExcelIn> fieldExcelCellEntry : fieldMap.entrySet()) {
                        Field field = fieldExcelCellEntry.getKey();
                        //设置可访问私有属性
                        field.setAccessible(true);
                        Object cellValue = getCellValue(
                            row.getCell(fieldExcelCellEntry.getValue().value()), field.getType());
                        fieldExcelCellEntry.getKey().set(t, cellValue);
                    }
                    resultList.add(t);
                } catch (Exception e) {
                    logger.error("Excel指定sheet第【{}】行，解析失败 原因:{}", i + 1, e.getMessage());
                }
            }
        } catch (Exception e) {
            logger.error("Excel指定sheet读取异常", e);
        }
        return resultList;
    }

    /**
     * 获得指定类 指定注解类型的field
     *
     * @param cls 目标类
     * @param annotation 注解类
     */
    private static <T, F extends Annotation> Map<Field, F> getClassField(Class<T> cls,
        Class<F> annotation) {
        Map<Field, F> map = new HashMap<>();
        if (cls == null || annotation == null) {
            return map;
        }
        Field[] fields = cls.getDeclaredFields();
        for (Field field : fields) {
            F f = field.getAnnotation(annotation);
            if (f != null) {
                map.put(field, f);
            }
        }
        return map;
    }

    /**
     * 获得Cell多类型值
     */
    private static Object getCellValue(Cell cell, Class paramType) {
        if (cell == null || paramType == null) {
            return null;
        }
        try {
            if (paramType.equals(String.class)) {
                return cell.getStringCellValue();
            }
            if (paramType.equals(Boolean.class) || "boolean".equals(paramType.getName())) {
                return cell.getBooleanCellValue();
            }
            if (paramType.equals(Integer.class) || "int".equals(paramType.getName())) {
                return Integer.parseInt(cell.getStringCellValue());
            }
            if (paramType.equals(Long.class) || "long".equals(paramType.getName())) {
                return Long.parseLong(cell.getStringCellValue());
            }
            if (paramType.equals(Double.class) || "double".equals(paramType.getName())) {
                return cell.getNumericCellValue();
            }
            if (paramType.equals(Date.class)) {
                return cell.getDateCellValue();
            }
            if (paramType.equals(BigDecimal.class)) {
                return new BigDecimal(cell.getNumericCellValue());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }
}

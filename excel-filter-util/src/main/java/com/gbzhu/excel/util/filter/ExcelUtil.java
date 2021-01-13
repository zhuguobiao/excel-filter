package com.gbzhu.excel.util.filter;

import org.apache.log4j.Logger;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

public class ExcelUtil {

    private Logger logger;

    private String inFilePath;

    private String outFilePath;

    private String filterColumn;

    public ExcelUtil(String inFilePath, String outFilePath, String filterColumn, Logger logger) {
        this.logger = logger;
        this.inFilePath = inFilePath;
        this.outFilePath = outFilePath;
        this.filterColumn = filterColumn;
    }

    public void filterExcel() {
        InputStream inp = null;
        Workbook workbook = null;
        int filterColumnIndex = -1;
        try {
            inp = new FileInputStream(inFilePath);
            workbook = WorkbookFactory.create(inp);

            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(0);
            // 如果未发现有行,则返回
            if (row == null)
                return;
            // 第一行规定必须是标题
            for (Cell cell : row) {
                if (cell.getCellTypeEnum() == CellType.STRING && cell.getStringCellValue().equals(filterColumn)) {
                    filterColumnIndex = cell.getColumnIndex();
                    break;
                }
            }
            // 如果待筛选的字段一个也没找到就返回
            if (filterColumnIndex == -1) {
                logger.error("指定分组的列不存在。");
            }

            Map<String, XSSFWorkbook> workbookGroup = new HashMap<>();

            for (Row r : sheet) {
                if (r.getRowNum() == 0) continue;
                Cell cell = r.getCell(filterColumnIndex);
                if (cell.getCellTypeEnum() == CellType.STRING) {
                    String value = cell.getStringCellValue();
                    if (workbookGroup.containsKey(value)) {
                        XSSFWorkbook workbookTmp = workbookGroup.get(value);
                        Sheet sheetTmp = workbookTmp.getSheetAt(0);
                        int lastIndex = sheetTmp.getLastRowNum();
                        Row lastRow = sheetTmp.createRow(lastIndex + 1);
                        cloneRow(workbook, workbookTmp, r, lastRow);
                    } else {
                        XSSFWorkbook workbookTmp = new XSSFWorkbook();
                        workbookGroup.put(value, workbookTmp);
                        logger.info("创建表" + value + ".xlsx ...");
                        Sheet sheet_w = workbookTmp.createSheet("Sheet1");
                        //添加标题行
                        Row row_w = sheet_w.createRow(0);
                        String headerInfo = "";
                        for (Cell srcCell : row) {
                            row_w.createCell(srcCell.getColumnIndex()).setCellValue(srcCell.getStringCellValue());
                            headerInfo = headerInfo + srcCell.getStringCellValue() + "\t";
                        }
                        logger.info("在表" + value + ".xlsx中添加表头：" + headerInfo);
                        Row targetRow1 = sheet_w.createRow(1);
                        cloneRow(workbook, workbookTmp, r, targetRow1);
                    }
                }

            }

            for (String fileName : workbookGroup.keySet()) {
                FileOutputStream fos = null;
                try {
                    File file = new File(outFilePath + "\\" + fileName + ".xlsx");
                    fos = new FileOutputStream(file);
                    workbookGroup.get(fileName).write(fos);
                } catch (Exception e) {
                    e.printStackTrace();
                    logger.error("写入文件异常", e);
                } finally {
                    if (fos != null) {
                        try {
                            fos.close();
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                }
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (EncryptedDocumentException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (inp != null) {
                try {
                    inp.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

    }

    public void cloneRow(Workbook srcWorkbook, XSSFWorkbook targetWorkbook, Row srcRow, Row targetRow) {
        FormulaEvaluator evaluator = srcWorkbook.getCreationHelper().createFormulaEvaluator();
        for (Cell srcCell : srcRow) {
            Cell targetCell = targetRow.createCell(srcCell.getColumnIndex(), srcCell.getCellTypeEnum());
            CellStyle cs = targetWorkbook.createCellStyle();
            cs.cloneStyleFrom(srcCell.getCellStyle());
            targetCell.setCellStyle(cs);
            switch (srcCell.getCellTypeEnum()) {
                case _NONE:
                case BLANK:
                    break;
                case BOOLEAN:
                    targetCell.setCellValue(srcCell.getBooleanCellValue());
                    break;
                case ERROR:
                    targetCell.setCellErrorValue(srcCell.getErrorCellValue());
                    break;
                case FORMULA:
                    //唯一的问题就是公式由于是根据单元格计算,新表可能相应的单元格存放不同的数据,造成计算结果不同
                    //c_w.setCellFormula(c.getCellFormula());
                    CellValue cellValue = evaluator.evaluate(srcCell);
                    targetCell.setCellType(cellValue.getCellType());
                    switch (cellValue.getCellTypeEnum()) {
                        case BOOLEAN:
                            targetCell.setCellValue(cellValue.getBooleanValue());
                            break;
                        case ERROR:
                            targetCell.setCellErrorValue(cellValue.getErrorValue());
                            break;
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(srcCell)) {
                                targetCell.setCellValue(srcCell.getDateCellValue());
                            } else {
                                targetCell.setCellValue(srcCell.getNumericCellValue());
                            }
                            break;
                        case STRING:
                            targetCell.setCellValue(srcCell.getRichStringCellValue());
                            break;
                        default:
                            break;
                    }
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(srcCell)) {
                        targetCell.setCellValue(srcCell.getDateCellValue());
                    } else {
                        targetCell.setCellValue(srcCell.getNumericCellValue());
                    }
                    break;
                case STRING:
                    targetCell.setCellValue(srcCell.getRichStringCellValue());
                    break;
            }
        }
    }
}

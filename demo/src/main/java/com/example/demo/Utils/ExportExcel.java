package com.example.demo.Utils;

import org.apache.log4j.Logger;
import java.io.*;
import java.lang.reflect.Field;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import javax.servlet.http.HttpServletResponse;

/**
 * @author IsMayan
 * @ClassName: ExportExcel
 * @Description: Excel导出工具类
 * @date 2019-10-12 11:10:48
 */
@Component
public class ExportExcel {
    private static final Logger logger = Logger.getLogger(ExportExcel.class);


    /**
     * @param name
     * @param mapList(headers 字段名,dataList 数据集合,fileName sheet的名字)
     * @param response
     * @return <T>
     * @throws
     * @Title: exportMultisheetExcel
     * @Description:
     */
    public <T> void exportMultisheetExcel(String name, List<Map> mapList, HttpServletResponse response) {
        BufferedOutputStream bos = null;
        try {
            String fileName = name + ".xlsx";
            bos = getBufferedOutputStream(fileName, response);
            doExport(mapList, bos);
        } catch (Exception e) {
            logger.error(e);
        } finally {
            try {
                if (bos != null) {
                    bos.close();
                }
            } catch (IOException e) {
                logger.error(e);
            }
        }
    }


    private BufferedOutputStream getBufferedOutputStream(String fileName, HttpServletResponse response) throws Exception {
        response.setContentType("application/x-msdownload");
        response.setHeader("Content-Disposition",
                "attachment;filename=\"" + getFileName() + ".xls" + "\"");
        return new BufferedOutputStream(response.getOutputStream());
    }

    private static <T> void doExport(List<Map> mapList, OutputStream outputStream) {
        int maxBuff = 9999999;
        // 创建excel工作文本，100表示默认允许保存在内存中的行数
        SXSSFWorkbook wb = new SXSSFWorkbook(maxBuff);
        try {
            for (int i = 0; i < mapList.size(); i++) {
                Map map = mapList.get(i);
                String[] headers = (String[]) map.get("headers");
                Collection<T> dataList = (Collection<T>) map.get("dataList");
                String fileName = (String) map.get("fileName");
                createSheet(wb, null, headers, dataList, fileName, maxBuff);
            }

            if (outputStream != null) {
                wb.write(outputStream);
            }
        } catch (Exception e) {
            logger.error(e);
        }

    }

    private static <T> void createSheet(SXSSFWorkbook wb, String[] exportFields, String[] headers, Collection<T> dataList, String fileName, int maxBuff)
            throws NoSuchFieldException, IllegalAccessException, IOException {

        Sheet sh = wb.createSheet(fileName);

        CellStyle style = wb.createCellStyle();
        CellStyle style2 = wb.createCellStyle();
        //创建表头
        Font font = wb.createFont();
        font.setFontName("宋体");
        font.setFontHeightInPoints((short) 10);//设置字体大小
//        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD); //宽度
        style.setFont(font);//选择需要用到的字体格式

        style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);// 设置背景色(字段名的)
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中

        style.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框

        style2.setFont(font);//选择需要用到的字体格式

        style2.setFillForegroundColor(HSSFColor.WHITE.index);// 设置背景色
        style2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style2.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER); //垂直居中
//        style2.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 水平向下居中
        style2.setAlignment(HSSFCellStyle.ALIGN_LEFT); //水平布局：居左
        style2.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
        style2.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
        style2.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
        style2.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框

        Row headerRow = sh.createRow(0); //表头

        int headerSize = headers.length;
        for (int cellnum = 0; cellnum < headerSize; cellnum++) {
            Cell cell = headerRow.createCell(cellnum);
            cell.setCellStyle(style);
//            sh.setColumnWidth(cellnum, 4000);
            cell.setCellValue(headers[cellnum]);
        }

        int rownum = 0;
        Iterator<T> iterator = dataList.iterator();
        while (iterator.hasNext()) {
            T data = iterator.next();
            Row row = sh.createRow(rownum + 1);

            Field[] fields = getExportFields(data.getClass(), exportFields);
            for (int cellnum = 0; cellnum < headerSize; cellnum++) {
                Cell cell = row.createCell(cellnum);
                cell.setCellStyle(style2);
                Field field = fields[cellnum];

                setData(field, data, field.getName(), cell);
            }
            rownum = sh.getLastRowNum();
            // 大数据量时将之前的数据保存到硬盘
            if (rownum % maxBuff == 0) {
                ((SXSSFSheet) sh).flushRows(maxBuff); // 超过maxBuff行后将之前的数据刷新到硬盘

            }
        }
    }

    private static <T> void doExport(String[] headers, String[] exportFields, Collection<T> dataList,
                                     String fileName, OutputStream outputStream) {
        int maxBuff = 99999999;
        // 创建excel工作文本，100表示默认允许保存在内存中的行数
        SXSSFWorkbook wb = new SXSSFWorkbook(maxBuff);
        try {
            createSheet(wb, exportFields, headers, dataList, fileName, maxBuff);
            if (outputStream != null) {
                wb.write(outputStream);
            }
        } catch (Exception e) {
            logger.error(e);
        }
    }

    /**
     * 获取单条数据的属性
     *
     * @param object
     * @param property
     * @param <T>
     * @return
     * @throws NoSuchFieldException
     * @throws IllegalAccessException
     */
    private static <T> Field getDataField(T object, String property)
            throws NoSuchFieldException, IllegalAccessException {
        Field dataField;
        if (property.contains(".")) {
            String p = property.substring(0, property.indexOf("."));
            dataField = object.getClass().getDeclaredField(p);
            return dataField;
        } else {
            dataField = object.getClass().getDeclaredField(property);
        }
        return dataField;
    }

    private static Field[] getExportFields(Class<?> targetClass, String[] exportFieldNames) {
        Field[] fields = null;
        if (exportFieldNames == null || exportFieldNames.length < 1) {
            fields = targetClass.getDeclaredFields();
        } else {
            fields = new Field[exportFieldNames.length];
            for (int i = 0; i < exportFieldNames.length; i++) {
                try {
                    fields[i] = targetClass.getDeclaredField(exportFieldNames[i]);
                } catch (Exception e) {
                    try {
                        fields[i] = targetClass.getSuperclass().getDeclaredField(exportFieldNames[i]);
                    } catch (Exception e1) {
                        throw new IllegalArgumentException("无法获取导出字段", e);
                    }

                }
            }
        }
        return fields;
    }

    /**
     * 根据属性设置对应的属性值
     *
     * @param dataField 属性
     * @param object    数据对象
     * @param property  表头的属性映射
     * @param cell      单元格
     * @param <T>
     * @return
     * @throws IllegalAccessException
     * @throws NoSuchFieldException
     */
    private static <T> void setData(Field dataField, T object, String property, Cell cell)
            throws IllegalAccessException, NoSuchFieldException {
        dataField.setAccessible(true); //允许访问private属性
        Object val = dataField.get(object); //获取属性值
        Sheet sh = cell.getSheet(); //获取excel工作区
        CellStyle style = cell.getCellStyle(); //获取单元格样式
        int cellnum = cell.getColumnIndex();
        if (val != null) {
            if (dataField.getType().toString().endsWith("String")) {
                cell.setCellValue((String) val);
            } else if (dataField.getType().toString().endsWith("Integer") || dataField.getType().toString().endsWith("int")) {
                cell.setCellValue((Integer) val);
            } else if (dataField.getType().toString().endsWith("Long") || dataField.getType().toString().endsWith("long")) {
                cell.setCellValue(val.toString());
            } else if (dataField.getType().toString().endsWith("Double") || dataField.getType().toString().endsWith("double")) {
                cell.setCellValue((Double) val);
            } else if (dataField.getType().toString().endsWith("Date")) {
                DateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                cell.setCellValue(format.format((Date) val));
            } else if (dataField.getType().toString().endsWith("List")) {
                List list1 = (List) val;
                int size = list1.size();
                for (int i = 0; i < size; i++) {
                    //加1是因为要去掉点号
                    int start = property.indexOf(dataField.getName()) + dataField.getName().length() + 1;
                    String tempProperty = property.substring(start, property.length());
                    Field field = getDataField(list1.get(i), tempProperty);
                    Cell tempCell = cell;
                    if (i > 0) {
                        int rowNum = cell.getRowIndex() + i;
                        Row row = sh.getRow(rowNum);
                        if (row == null) {//另起一行
                            row = sh.createRow(rowNum);
                            //合并之前的空白单元格（在这里需要在header中按照顺序把list类型的字段放到最后，方便显示和合并单元格）
                            for (int j = 0; j < cell.getColumnIndex(); j++) {
                                sh.addMergedRegion(new CellRangeAddress(cell.getRowIndex(), cell.getRowIndex() + size - 1, j, j));
                                Cell c = row.createCell(j);
                                c.setCellStyle(style);
                            }
                        }
                        tempCell = row.createCell(cellnum);
                        tempCell.setCellStyle(style);
                    }
                    //递归传参到单元格并获取偏移量（这里获取到的偏移量都是第二层后list的偏移量）
                    setData(field, list1.get(i), tempProperty, tempCell);
                }
            } else {
                if (property.contains(".")) {
                    String p = property.substring(property.indexOf(".") + 1, property.length());
                    Field field = getDataField(val, p);
                    setData(field, val, p, cell);
                } else {
                    cell.setCellValue(val.toString());
                }
            }
        }
    }

    /**
     * 根据文件名判断用XXXXWorkbook
     *
     * @param filePath
     * @return
     */

    private static Workbook createWorkBook(String filePath) {
        Workbook wb = null;
        if (filePath == null) {
            return null;
        }
        String extString = filePath.substring(filePath.lastIndexOf("."));
        InputStream is = null;
        try {
            is = new FileInputStream(filePath);
            if (".xls".equals(extString)) {
                return wb = new HSSFWorkbook(is);
            } else if (".xlsx".equals(extString)) {
                return wb = new XSSFWorkbook(is);
            } else if (".xlsm".equals(extString)) {
                return new XSSFWorkbook(is);
            } else {
                return wb;
            }
        } catch (FileNotFoundException e) {
            logger.error(e);
        } catch (IOException e) {
            logger.error(e);
        }
        return wb;
    }

    /**
     * 用日期编写名字
     *
     * @return
     */

    protected String getFileName() {
        SimpleDateFormat datetime = new SimpleDateFormat("yyyyMMddhhmmssSSS");
        Date time = new Date();
        String name = datetime.format(time);
        return name;
    }
}

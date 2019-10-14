package com.example.demo.Utils;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

import static org.apache.commons.lang.StringUtils.isBlank;

/**
 * @author IsMayan
 * @ClassName: ImportExcel
 * @Description: Excel导入工具类
 * @date 2019-10-12 11:10:48
 */
@Component
public class ImportExcel {
    private static final Logger logger = Logger.getLogger(ImportExcel.class);

    private static final String path = "E:/UpLoad/";

    public  List<Object> importExcel(Object vo, MultipartFile file) {
        String originalFilename = file.getOriginalFilename();
        String prefix = originalFilename.substring(originalFilename.lastIndexOf(".") + 1);
        String fileName = "上传表" + "." + prefix;
        File newFile = null;
        try {
            File fileone = new File(path , fileName);
            //查看路径是否存在，不存在就创建
            if (!fileone.getParentFile().exists()) {
                fileone.getParentFile().mkdirs();
            }
            newFile = new File(path  + fileName);
            file.transferTo(newFile);

        } catch (Exception e) {
            logger.error(e);
        }

        InputStream input = null;
        try {
            input = new FileInputStream(newFile);
        } catch (FileNotFoundException e) {
            logger.error(e);
        }
        return  importDataFromExcel(vo,input,fileName);
    }


    /**
     * @param @param  vo javaBean
     * @param @param  is 输入流
     * @param @param  excelFileName
     * @param @return
     * @return List<Object>
     * @throws
     * @Title: importDataFromExcel
     * @Description: 将sheet中的数据保存到list中，
     * 1、调用此方法时，vo的属性个数必须和excel文件每行数据的列数相同且一一对应，vo的所有属性都为String
     * 2、在action调用此方法时，需声明
     * private File excelFile;上传的文件
     * private String excelFileName;原始文件的文件名
     * 3、页面的file控件name需对应File的文件名
     */
    public  List<Object> importDataFromExcel(Object vo, InputStream is, String excelFileName) {

        List<Object> list = new ArrayList<>();

        try {
            //创建工作簿
            Workbook workbook = this.createWorkbook(is, excelFileName);
            //创建工作表sheet
            Sheet sheet = this.getSheet(workbook, 0);
            //获取sheet中数据的行数
            int rows = sheet.getPhysicalNumberOfRows();
            //获取表头单元格个数
            int cells = sheet.getRow(0).getPhysicalNumberOfCells();
            //利用反射，给JavaBean的属性进行赋值
            Field[] fields = vo.getClass().getDeclaredFields();
            for (int i = 1; i < rows; i++) {//第一行为标题栏，从第二行开始取数据
                Row row = sheet.getRow(i);
                int index = 0;
                while (index < cells) {
                    String value = (String) ImportExcel.getCellFormatValue(row.getCell(index));
//                    Cell cell = row.getCell(index);
//                    if (null == cell) {
//                        cell = row.createCell(index);
//                    }
//                    cell.setCellType(Cell.CELL_TYPE_STRING);
//                    String value = null == cell.getStringCellValue() ? "" : cell.getStringCellValue();
                    Field field = fields[index];
                    String fieldName = field.getName();
                    String methodName = "set" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
                    Method setMethod = vo.getClass().getMethod(methodName, new Class[]{String.class});
                    setMethod.invoke(vo, new Object[]{value});
                    index++;

                }
                if (isHasValues(vo)) {//判断对象属性是否有值
                    list.add(vo);
                    vo = vo.getClass().getConstructor(new Class[]{}).newInstance(new Object[]{});//重新创建一个vo对象
                }

            }

        } catch (Exception e) {
            logger.error(e);
        } finally {
            try {
                is.close();//关闭流
            } catch (Exception e2) {
                logger.error(e2);
            }
        }
        return list;

    }

    /**
     * @param @param  is
     * @param @param  excelFileName
     * @param @return
     * @param @throws IOException
     * @return Workbook
     * @throws
     * @Title: createWorkbook
     * @Description: 判断excel文件后缀名，生成不同的workbook
     */
    public Workbook createWorkbook(InputStream is, String excelFileName) throws IOException {
        if (excelFileName.endsWith("xls")) {
            return new HSSFWorkbook(is);
        } else if (excelFileName.endsWith("xlsx")) {
            return new XSSFWorkbook(is);
        } else if(excelFileName.endsWith("xlsm")) {
            return new XSSFWorkbook(is);
        }
        return null;
    }

    /**
     * @param @param  object
     * @param @return
     * @return boolean
     * @throws
     * @Title: isHasValues
     * @Description: 判断一个对象所有属性是否有值，如果一个属性有值(非空)，则返回true
     */
    public boolean isHasValues(Object object) {
        Field[] fields = object.getClass().getDeclaredFields();
        boolean flag = false;
        for (int i = 0; i < fields.length; i++) {
            String fieldName = fields[i].getName();
            String methodName = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
            Method getMethod;
            try {
                getMethod = object.getClass().getMethod(methodName);
                Object obj = getMethod.invoke(object);
                if (null != obj&& ! "".equals(obj)) {
                    flag = true;
                    break;
                }
            } catch (Exception e) {
                logger.error(e);
            }

        }
        return flag;

    }

    /**
     * @param @param  workbook
     * @param @param  sheetIndex
     * @param @return
     * @return Sheet
     * @throws
     * @Title: getSheet
     * @Description: 根据sheet索引号获取对应的sheet
     */
    public Sheet getSheet(Workbook workbook, int sheetIndex) {
        return workbook.getSheetAt(0);
    }

    /**
     * 将字段转为相应的格式
     *
     * @param cell
     * @return
     */
    private static Object getCellFormatValue(Cell cell) {
        Object cellValue = null;
        if (cell != null) {
            //判断cell类型
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC: {
                   String ce = String.valueOf(cell.getNumericCellValue());
                   if(ce.substring(ce.lastIndexOf(".")+1).length() <= 1){
                       cellValue = String.valueOf( (int)(cell.getNumericCellValue()) );
                   }else {
                       cellValue = String.valueOf(cell.getNumericCellValue());
                   }
                    break;
                }
                case Cell.CELL_TYPE_FORMULA: {
                    if (DateUtil.isCellDateFormatted(cell)) {
                        cellValue = cell.getDateCellValue();////转换为日期格式YYYY-mm-dd
                    } else {
                        cellValue = String.valueOf(cell.getNumericCellValue()); //数字
                    }
                    break;
                }
                case Cell.CELL_TYPE_STRING: {
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                }
                default:
                    cellValue = "";
            }
        } else {
            cellValue = "";
        }
        return cellValue;
    }

}

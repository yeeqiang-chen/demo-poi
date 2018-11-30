package com.yiqiang.learn.poi.util;

import com.yiqiang.learn.poi.annotation.ExcelResources;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections.MapUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

/**
 * Title:
 * Description:
 *      该类实现了将一组对象转换为Excel表格，并且可以从Excel表格中读取到一组List对象中
 *      该类利用了BeanUtils框架中的反射完成
 *      使用该类的前提，在相应的实体对象上通过ExcelReources来完成相应的注解
 * Create Time: 2018/10/22 1:01
 *
 * @author: YEEChan
 * @version: 1.0
 */
public class ExcelUtil {

    private static final org.slf4j.Logger LOGGER = LoggerFactory.getLogger(ExcelUtil.class);
    private ExcelUtil() {
    }

    private static class SingletonHolder {
        private static ExcelUtil instance = new ExcelUtil();
    }

    public static ExcelUtil getInstance() {
        return SingletonHolder.instance;
    }

    /**
     * 处理对象转换为Excel
     *
     * @param template
     * @param objs
     * @param clz
     * @param isClasspath
     * @return
     */
    private ExcelTemplate handlerObj2Excel(String template, List objs, Class clz, boolean isClasspath) {
        ExcelTemplate et = ExcelTemplate.getInstance();
        try {
            if (isClasspath) {
                et.readTemplateByClasspath(template);
            } else {
                et.readTemplateByPath(template);
            }
            List<ExcelHeader> headers = getHeaderList(clz);
            Collections.sort(headers);
            //输出值
            for (Object obj : objs) {
                et.createNewRow();
                for (ExcelHeader eh : headers) {
                    et.createCell(BeanUtils.getProperty(obj, getMethodName(eh)));
                }
            }
        } catch (Exception e) {
            LOGGER.error("ExcelUtil.handlerObj2ExcelWithTemplate.ex", e);
        }
        return et;
    }

    /**
     * 根据标题获取相应的方法名称
     *
     * @param eh
     * @return
     */
    private String getMethodName(ExcelHeader eh) {
        String mn = eh.getMethodName().substring(3);
        mn = mn.substring(0, 1).toLowerCase() + mn.substring(1);
        return mn;
    }

    /**
     * 将对象转换为Excel并且导出，该方法是基于模板的导出，导出到流
     *
     * @param datas       模板中的替换的常量数据
     * @param template    模板路径
     * @param os          输出流
     * @param objs        对象列表
     * @param clz         对象的类型
     * @param isClasspath 模板是否在classPath路径下
     * @param ifInsertSer 是否插入序号
     */
    public void exportObj2ExcelByTemplate(Map<String, String> datas, String template, OutputStream os,
                                          List objs, Class clz, boolean isClasspath, boolean ifInsertSer) {
        ExcelTemplate et = handlerObj2Excel(template, objs, clz, isClasspath);
        if (MapUtils.isNotEmpty(datas)) {
            et.replaceFinalData(datas);
        }
        if (ifInsertSer) {
            et.insertSer();
        }
        et.wirteToStream(os);
    }

    /**
     * 将对象转换为Excel并且导出，该方法是基于模板的导出，导出到一个具体的路径中
     *
     * @param datas       模板中的替换的常量数据
     * @param template    模板路径
     * @param outPath     输出路径
     * @param objs        对象列表
     * @param clz         对象的类型
     * @param isClasspath 模板是否在classPath路径下
     * @param ifInsertSer 是否插入序号
     */
    public void exportObj2ExcelByTemplate(Map<String, String> datas, String template, String outPath,
                                          List objs, Class clz, boolean isClasspath, boolean ifInsertSer) {
        ExcelTemplate et = handlerObj2Excel(template, objs, clz, isClasspath);
        if (MapUtils.isNotEmpty(datas)) {
            et.replaceFinalData(datas);
        }
        if (ifInsertSer) {
            et.insertSer();
        }
        et.writeToFile(outPath);
    }

    /**
     * 将对象转换为Excel并且导出，该方法是基于模板的导出，导出到流,基于Properties作为常量数据
     *
     * @param prop        基于Properties的常量数据模型
     * @param template    模板路径
     * @param os          输出流
     * @param objs        对象列表
     * @param clz         对象的类型
     * @param isClasspath 模板是否在classPath路径下
     */
    public void exportObj2ExcelByTemplate(Properties prop, String template, OutputStream os, List objs, Class clz, boolean isClasspath) {
        ExcelTemplate et = handlerObj2Excel(template, objs, clz, isClasspath);
        et.replaceFinalData(prop);
        et.wirteToStream(os);
    }

    /**
     * 将对象转换为Excel并且导出，该方法是基于模板的导出，导出到一个具体的路径中,基于Properties作为常量数据
     *
     * @param prop        基于Properties的常量数据模型
     * @param template    模板路径
     * @param outPath     输出路径
     * @param objs        对象列表
     * @param clz         对象的类型
     * @param isClasspath 模板是否在classPath路径下
     */
    public void exportObj2ExcelByTemplate(Properties prop, String template, String outPath, List objs, Class clz, boolean isClasspath) {
        ExcelTemplate et = handlerObj2Excel(template, objs, clz, isClasspath);
        et.replaceFinalData(prop);
        et.writeToFile(outPath);
    }

    private Workbook handleObj2Excel(List objs, Class clz, boolean isXssf) {
        Workbook wb = null;
        try {
            if (isXssf) {
                wb = new XSSFWorkbook();
            } else {
                wb = new HSSFWorkbook();
            }
            CellStyle cellStyle = wb.createCellStyle();
            Sheet sheet = wb.createSheet();
            Row r = sheet.createRow(0);
            List<ExcelHeader> headers = getHeaderList(clz);
            Collections.sort(headers);

            // 设置样式
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            cellStyle.setTopBorderColor(IndexedColors.BLACK.index);
            cellStyle.setBottomBorderColor(IndexedColors.BLACK.index);
            cellStyle.setLeftBorderColor(IndexedColors.BLACK.index);
            cellStyle.setRightBorderColor(IndexedColors.BLACK.index);
            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setWrapText(true);
            Font font = wb.createFont();
            font.setBold(true);
            cellStyle.setFont(font);
            //写标题
            for (int i = 0; i < headers.size(); i++) {
                Cell cell = r.createCell(i);
                cell.setCellStyle(cellStyle);
                cell.setCellValue(headers.get(i).getTitle());
            }
            //写数据
            Object obj = null;
            for (int i = 0; i < objs.size(); i++) {
                r = sheet.createRow(i + 1);
                obj = objs.get(i);
                for (int j = 0; j < headers.size(); j++) {
                    Cell cell = r.createCell(j);
                    cell.setCellStyle(cellStyle);
                    cell.setCellValue(BeanUtils.getProperty(obj, getMethodName(headers.get(j))));
                }
            }
        } catch (Exception e) {
            LOGGER.error("ExcelUtil.handleObj2Excel.ex", e);
        }
        return wb;
    }

    /**
     * 导出对象到Excel，不是基于模板的，直接新建一个Excel完成导出，基于路径的导出
     *
     * @param outPath 导出路径
     * @param objs    对象列表
     * @param clz     对象类型
     * @param isXssf  是否是2007版本
     */
    public void exportObj2Excel(String outPath, List objs, Class clz, boolean isXssf) {
        Workbook wb = handleObj2Excel(objs, clz, isXssf);
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(outPath);
            wb.write(fos);
        } catch (Exception e) {
            LOGGER.error("ExcelUtil.exportObj2Excel.ex", e);
        } finally {
            try {
                if (fos != null) fos.close();
            } catch (IOException e) {
                LOGGER.error("ExcelUtil.exportObj2Excel.ex", e);
            }
        }
    }

    /**
     * 导出对象到Excel，不是基于模板的，直接新建一个Excel完成导出，基于流
     *
     * @param os     输出流
     * @param objs   对象列表
     * @param clz    对象类型
     * @param isXssf 是否是2007版本
     */
    public void exportObj2Excel(OutputStream os, List objs, Class clz, boolean isXssf) {
        try {
            Workbook wb = handleObj2Excel(objs, clz, isXssf);
            wb.write(os);
        } catch (IOException e) {
            LOGGER.error("ExcelUtil.exportObj2Excel.ex", e);
        }
    }

    /**
     * 从类路径读取相应的Excel文件到对象列表
     *
     * @param path     类路径下的path
     * @param clz      对象类型
     * @param readLine 开始行，注意是标题所在行
     * @param tailLine 底部有多少行，在读入对象时，会减去这些行
     * @return
     */
    public List<Object> readExcel2ObjsByClasspath(String path, Class clz, int readLine, int tailLine) {
        Workbook wb = null;
        try {
            wb = WorkbookFactory.create(ExcelUtil.class.getResourceAsStream(path));
            return handlerExcel2Objs(wb, clz, readLine, tailLine);
        } catch (Exception e) {
            LOGGER.error("ExcelUtil.readExcel2ObjsByClasspath.ex", e);
        }
        return null;
    }

    /**
     * 从文件路径读取相应的Excel文件到对象列表
     *
     * @param path     文件路径下的path
     * @param clz      对象类型
     * @param readLine 开始行，注意是标题所在行
     * @param tailLine 底部有多少行，在读入对象时，会减去这些行
     * @return
     */
    public List<Object> readExcel2ObjsByPath(String path, Class clz, int readLine, int tailLine) {
        Workbook wb = null;
        try {
            wb = WorkbookFactory.create(new File(path));
            return handlerExcel2Objs(wb, clz, readLine, tailLine);
        } catch (Exception e) {
            LOGGER.error("ExcelUtil.readExcel2ObjsByPath.ex", e);
        }
        return null;
    }

    /**
     * 从类路径读取相应的Excel文件到对象列表，标题行为0，没有尾行
     *
     * @param path 路径
     * @param clz  类型
     * @return 对象列表
     */
    public List<Object> readExcel2ObjsByClasspath(String path, Class clz) {
        return this.readExcel2ObjsByClasspath(path, clz, 0, 0);
    }

    /**
     * 从文件路径读取相应的Excel文件到对象列表，标题行为0，没有尾行
     *
     * @param path 路径
     * @param clz  类型
     * @return 对象列表
     */
    public List<Object> readExcel2ObjsByPath(String path, Class clz) {
        return this.readExcel2ObjsByPath(path, clz, 0, 0);
    }

    private String getCellValue(Cell c) {
        String o = null;
        if (c.getCellType() == CellType.BLANK) {
            o = "";
        } else if (c.getCellType() == CellType.BOOLEAN) {
            o = String.valueOf(c.getBooleanCellValue());
        } else if (c.getCellType() == CellType.FORMULA) {
            o = String.valueOf(c.getCellFormula());
        } else if (c.getCellType() == CellType.NUMERIC) {
            o = String.valueOf(c.getNumericCellValue());
        } else if (c.getCellType() == CellType.STRING) {
            o = c.getStringCellValue();
        } else {
            o = null;
        }
        return o;
    }

    /**
     *
     * @param wb
     * @param clz
     * @param readLine
     *          从第几行开始读
     * @param tailLine
     *          尾部有几行(放其它标记的#)
     * @return
     */
    private List<Object> handlerExcel2Objs(Workbook wb, Class clz, int readLine, int tailLine) {
        Sheet sheet = wb.getSheetAt(0);
        List<Object> objs = null;
        try {
            Row row = sheet.getRow(readLine);
            objs = new ArrayList<Object>();
            Map<Integer, String> maps = getHeaderMap(sheet, row, clz);
            if (maps == null || maps.size() <= 0) throw new RuntimeException("要读取的Excel的格式不正确，检查是否设定了合适的行");
            for (int i = readLine + 1; i <= sheet.getLastRowNum() - tailLine; i++) {
                row = sheet.getRow(i);
                Object obj = clz.newInstance();
                for (Cell c : row) {
                    int ci = c.getColumnIndex();
                    String mn = maps.get(ci).substring(3);
                    mn = mn.substring(0, 1).toLowerCase() + mn.substring(1);
                    BeanUtils.copyProperty(obj, mn, this.getCellValue(c));
                }
                objs.add(obj);
            }
        } catch (Exception e) {
            LOGGER.error("ExcelUtil.handlerExcel2Objs.ex", e);
        }
        return objs;
    }

    private List<ExcelHeader> getHeaderList(Class clz) {
        List<ExcelHeader> headers = new ArrayList<ExcelHeader>();
        Method[] ms = clz.getDeclaredMethods();
        for (Method m : ms) {
            String mn = m.getName();
            if (mn.startsWith("get")) {
                if (m.isAnnotationPresent(ExcelResources.class)) {
                    ExcelResources er = m.getAnnotation(ExcelResources.class);
                    headers.add(new ExcelHeader(er.order(), er.title(), mn));
                }
            }
        }
        return headers;
    }

    private Map<Integer, String> getHeaderMap(Sheet sheet, Row titleRow, Class clz) {
        List<ExcelHeader> headers = getHeaderList(clz);
        Map<Integer, String> maps = new HashMap<Integer, String>();
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        // 表头有合并,默认从第一行,第一列开始解析表头
        if (CollectionUtils.isNotEmpty(mergedRegions)) {
            mergedRegions.stream().forEach(x -> {
                int firstRow = x.getFirstRow();
                int firstColumn = x.getFirstColumn();
                int lastColumn = x.getLastColumn();
                // 表头横跨多个列,不计为属性字段
                if (firstColumn != lastColumn) {
                    for (int i = firstColumn; i <= lastColumn; i++) {
                        // row = firstRow + 1,虽然合并了,但是还是放在第一个row上,+1排除掉上面合并的title
                        String title = sheet.getRow(firstRow + 1).getCell(i).getStringCellValue();
                        for (ExcelHeader eh : headers) {
                            if (eh.getTitle().equals(title.trim())) {
                                maps.put(i, eh.getMethodName().replace("get", "set"));
                                break;
                            }
                        }

                    }
                } else {
                    String title = sheet.getRow(firstRow).getCell(firstColumn).getStringCellValue();
                    for (ExcelHeader eh : headers) {
                        if (eh.getTitle().equals(title.trim())) {
                            maps.put(firstColumn, eh.getMethodName().replace("get", "set"));
                            break;
                        }
                    }
                }
            });
            // 无合并
        } else {
            for (Cell c : titleRow) {
                String title = c.getStringCellValue();
                for (ExcelHeader eh : headers) {
                    if (eh.getTitle().equals(title.trim())) {
                        maps.put(c.getColumnIndex(), eh.getMethodName().replace("get", "set"));
                        break;
                    }
                }
            }
        }
        return maps;
    }
}

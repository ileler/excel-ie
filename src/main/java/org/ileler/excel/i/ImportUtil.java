package org.ileler.excel.i;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.ileler.excel.util.ExcelUtil;
import org.ileler.excel.util.LangCode;

import java.awt.Color;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;


/**
 * Created by ileler@qq.com on 2016/5/12.
 */
public final class ImportUtil {

    private static final int startNum = 1;

    public static <T> void exportErrorData(InputStream is, int sheetIndex, OutputStream out,
                                           EntityBean entityBean, Collection<T> dataset)
            throws IllegalAccessException, IllegalArgumentException,
            InvocationTargetException, NoSuchMethodException,
            SecurityException, IOException {
        FieldBean[] fieldBeans = entityBean == null ? null : entityBean.getFieldBeans();
        if (is == null || out == null || entityBean == null || fieldBeans.length < 1) {
            return;
        }
        Workbook workBook = null;
        CellStyle estyle = null;
        Font font = null;
        try {
            workBook = new HSSFWorkbook(is);
            estyle = workBook.createCellStyle();
            estyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
            estyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            estyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            font = workBook.createFont();
            font.setColor(HSSFColor.RED.index);
            estyle.setFont(font);
        } catch (Exception e) {
            workBook = new XSSFWorkbook(is);
            estyle = workBook.createCellStyle();
            estyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
            estyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
            estyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
            font = workBook.createFont();
            font.setColor(new XSSFColor(Color.RED).getIndexed());
            estyle.setFont(font);
        }

        Sheet sheet = workBook.getSheetAt(sheetIndex);
        workBook.setActiveSheet(sheetIndex);
        int i = 0, j = fieldBeans.length;
        Row titleRow = sheet.getRow(0);

        for (; i < j; i++) {
            FieldBean fieldBean = fieldBeans[i];
            if (!fieldBean.getTitle().equalsIgnoreCase(titleRow.getCell(i).getStringCellValue())) {
                break;
            }
        }
        if (i < j) {
            return;
        } else {
            Cell etitleCell = titleRow.createCell(j);
            etitleCell.setCellValue(entityBean.getEtitle());
            etitleCell.setCellStyle(estyle);
            etitleCell.setAsActiveCell();
        }

        Row templateRow = sheet.getRow(1);
        CellStyle rowStyle = templateRow.getRowStyle();
        java.util.List<CellStyle> cellStyles = new ArrayList<CellStyle>(j);
        for (i = 0; i < j; i++) {
            cellStyles.add(templateRow.getCell(i).getCellStyle());
        }

        for (int a = startNum, b = sheet.getLastRowNum(); a <= b; a++) {
            sheet.removeRow(sheet.getRow(a));
        }

        if (dataset != null) {
            Iterator<T> it = dataset.iterator();
            int index = startNum;
            while (it.hasNext()) {
                Row row = sheet.createRow(index);
                if (rowStyle != null) {
                    row.setRowStyle(rowStyle);
                }
                T t = it.next();
                int cellIndex = 0;
                for (i = 0; i < j; i++) {
                    FieldBean fieldBean = fieldBeans[i];
                    Cell cell = row.createCell(cellIndex);
                    cell.setCellStyle(cellStyles.get(i));
                    if (t instanceof Map) {
                        Map<String, Object> map = (Map<String, Object>)t;
                        cell.setCellValue(map.get(fieldBean.getEfName()).toString());
                    } else {
                        cell.setCellValue("");
                    }
                    cellIndex++;
                }
                Cell emsgCell = row.createCell(i);
                emsgCell.setCellStyle(estyle);
                if (t instanceof Map) {
                    Map<String, Object> map = (Map<String, Object>)t;
                    emsgCell.setCellValue(map.get(entityBean.getEfname()).toString());
                } else {
                    emsgCell.setCellValue("");
                }
                index++;
            }
        }
        workBook.write(out);
    }

    public static <T> void exportErrorData(InputStream is, int sheetIndex, InputStream templateN, LangCode langCode, OutputStream out,
                                           Collection<T> dataset) {
        ByteArrayOutputStream baos = null;
        try {
            baos = getBAOS(is);
            Map<Integer, String[]> map = ExcelUtil.readXLS(new ByteArrayInputStream(baos.toByteArray()), sheetIndex);
            if (map == null)  map = ExcelUtil.readXLSX(new ByteArrayInputStream(baos.toByteArray()), sheetIndex);
            if (map == null || map.size() < 1)  return;
            exportErrorData(new ByteArrayInputStream(baos.toByteArray()), sheetIndex, out, BeanUtil.getEntityBean(templateN, map.get(0), langCode), dataset);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (baos != null) {
                try {
                    baos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static <T> void exportErrorData(InputStream is, OutputStream out,
                                           InputStream templateN, LangCode langCode, Collection<T> dataset) {
        exportErrorData(is, 0, templateN, langCode, out, dataset);
    }

    private static boolean isEmpty(String[] cols) {
        if (cols == null || cols.length < 1)    return true;
        for (String col : cols) {
            if (!StringUtils.isBlank(col))  return false;
        }
        return true;
    }

    private static LeadinResult resolve(Map<Integer, String[]> map, EntityBean entityBean, LangCode langCode, Map<String, Object> session) {
        if (map == null || entityBean == null)     return null;
        Map<Integer, Map<String, Object>> rightList = new HashMap<>(0);
        Map<Integer, Map<String, Object>> errorList = new HashMap<>(0);
        //循环行
        for (int i = startNum, j = map.size(); i < j; i++) {
            String[] cols = map.get(i);
            if (isEmpty(cols))  continue;
            Map<String, Object> result = new HashMap<String, Object>();
            if (entityBean.checkRow(cols, result, langCode, session)) {
                if (result.size() < 1)   continue;
                rightList.put(i, result);
            } else {
                if (result.size() < 1 || errorList == null)   continue;
                errorList.put(i, result);
            }
        }
        return new LeadinResult((rightList != null && rightList.size() > 0 ? rightList : null), (errorList != null && errorList.size() > 0 ? errorList : null));
    }

    private static LeadinResult resolve(InputStream is, int sheet, Map<Integer, String[]> map, EntityBean entityBean, LangCode langCode, OutputStream errorOut, Map<String, Object> session) {
        if (is == null || map == null || entityBean == null)     return null;
        LeadinResult result = resolve(map, entityBean, langCode, session);
        if (errorOut != null && result != null && result.hasError()) {
            try {
                exportErrorData(is, sheet, errorOut, entityBean, result.getErrorList().values());
            } catch (IllegalAccessException | IllegalArgumentException
                    | InvocationTargetException | NoSuchMethodException
                    | SecurityException | IOException e) {
                e.printStackTrace();
            }
        }
        return result;
    }

    public static InputStream getTemplateName(InputStream is, int sheet, InputStream[] templateNS, LangCode langCode) {
        try {
            Map<Integer, String[]> map = ExcelUtil.readXLS(is, sheet);
            if (map == null) map = ExcelUtil.readXLSX(is, sheet);
            if (map == null || map.size() < 1)  return null;
            return BeanUtil.getTemplateName(templateNS, map.get(0), langCode);
        } catch (Exception e) {}
        return null;
    }

    public static InputStream getTemplateName(InputStream is, InputStream[] templateNS, LangCode langCode) {
        return getTemplateName(is, 0, templateNS, langCode);
    }

    public static LeadinResult resolve(InputStream is, int sheet, InputStream templateN, LangCode langCode, Map<String, Object> session) {
        if (is == null || templateN == null)     return null;
        ByteArrayOutputStream baos = null;
        try {
            baos = getBAOS(is);
            Map<Integer, String[]> map = ExcelUtil.readXLS(new ByteArrayInputStream(baos.toByteArray()), sheet);
            if (map == null)  map = ExcelUtil.readXLSX(new ByteArrayInputStream(baos.toByteArray()), sheet);
            if (map == null || map.size() < 1)  return null;
            return resolve(map, BeanUtil.getEntityBean(templateN, map.get(0), langCode), langCode, session);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (baos != null) {
                try {
                    baos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return null;
    }

    public static LeadinResult resolve(InputStream is, InputStream templateN, LangCode langCode, Map<String, Object> session) {
        return resolve(is, 0, templateN, langCode, session);
    }

    public static LeadinResult resolve(InputStream is, int sheet, InputStream templateN, LangCode langCode, OutputStream errorOut, Map<String, Object> session) {
        if (is == null || templateN == null)     return null;
        ByteArrayOutputStream baos = null;
        try {
            baos = getBAOS(is);
            Map<Integer, String[]> map = ExcelUtil.readXLS(new ByteArrayInputStream(baos.toByteArray()), sheet);
            if (map == null)  map = ExcelUtil.readXLSX(new ByteArrayInputStream(baos.toByteArray()), sheet);
            if (map == null || map.size() < 1)  return null;
            return resolve(new ByteArrayInputStream(baos.toByteArray()), sheet, map, BeanUtil.getEntityBean(templateN, map.get(0), langCode), langCode, errorOut, session);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (baos != null) {
                try {
                    baos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return null;
    }

    public static LeadinResult resolve(InputStream is, InputStream templateN, LangCode langCode, OutputStream errorOut, Map<String, Object> session) {
        return resolve(is, 0, templateN, langCode, errorOut, session);
    }

    public static LeadinResult resolveXLS(InputStream is, int sheet, InputStream templateN, LangCode langCode, Map<String, Object> session) {
        if (is == null)     return null;
        Map<Integer, String[]> map = ExcelUtil.readXLS(is, sheet);
        if (map == null || map.size() < 1)  return null;
        return resolve(map, BeanUtil.getEntityBean(templateN, map.get(0), langCode), langCode, session);
    }

    public static LeadinResult resolveXLS(InputStream is, InputStream templateN, LangCode langCode, Map<String, Object> session) {
        return resolveXLS(is, 0, templateN, langCode, session);
    }

    public static LeadinResult resolveXLSX(InputStream is, int sheet, InputStream templateN, LangCode langCode, Map<String, Object> session) {
        if (is == null)     return null;
        Map<Integer, String[]> map = ExcelUtil.readXLSX(is, sheet);
        if (map == null || map.size() < 1)  return null;
        return resolve(map, BeanUtil.getEntityBean(templateN, map.get(0), langCode), langCode, session);
    }

    public static LeadinResult resolveXLSX(InputStream is, InputStream templateN, LangCode langCode, Map<String, Object> session) {
        return resolveXLSX(is, 0, templateN, langCode, session);
    }

    public static LeadinResult resolveXLS(InputStream is, int sheet, InputStream templateN, LangCode langCode, OutputStream errorOut, Map<String, Object> session) {
        if (is == null)     return null;
        ByteArrayOutputStream baos = null;
        try {
            baos = getBAOS(is);
            Map<Integer, String[]> map = ExcelUtil.readXLS(new ByteArrayInputStream(baos.toByteArray()), sheet);
            if (map == null || map.size() < 1)  return null;
            return resolve(new ByteArrayInputStream(baos.toByteArray()), sheet, map, BeanUtil.getEntityBean(templateN, map.get(0), langCode), langCode, errorOut, session);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (baos != null) {
                try {
                    baos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return null;
    }

    public static LeadinResult resolveXLS(InputStream is, InputStream templateN, LangCode langCode, OutputStream errorOut, Map<String, Object> session) {
        return resolveXLS(is, 0, templateN, langCode, errorOut, session);
    }

    public static LeadinResult resolveXLSX(InputStream is, int sheet, InputStream templateN, LangCode langCode, OutputStream errorOut, Map<String, Object> session) {
        if (is == null)     return null;
        ByteArrayOutputStream baos = null;
        try {
            baos = getBAOS(is);
            Map<Integer, String[]> map = ExcelUtil.readXLSX(new ByteArrayInputStream(baos.toByteArray()), sheet);
            if (map == null || map.size() < 1)  return null;
            return resolve(new ByteArrayInputStream(baos.toByteArray()), sheet, map, BeanUtil.getEntityBean(templateN, map.get(0), langCode), langCode, errorOut, session);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (baos != null) {
                try {
                    baos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return null;
    }

    public static LeadinResult resolveXLSX(InputStream is, InputStream templateN, LangCode langCode, OutputStream errorOut, Map<String, Object> session) {
        return resolveXLSX(is, 0, templateN, langCode, errorOut, session);
    }

    public static LeadinResult resolveXLS(InputStream is, int sheet, InputStream[] templateNS, LangCode langCode, Map<String, Object> session) {
        if (is == null || templateNS == null || templateNS.length < 1)  return null;
        Map<Integer, String[]> map = ExcelUtil.readXLS(is, sheet);
        if (map == null || map.size() < 1)  return null;
        return resolve(map, BeanUtil.getEntityBean(templateNS, map.get(0), langCode), langCode, session);
    }

    public static LeadinResult resolveXLS(InputStream is, InputStream[] templateNS, LangCode langCode, Map<String, Object> session) {
        return resolveXLS(is, 0, templateNS, langCode, session);
    }

    public static LeadinResult resolveXLSX(InputStream is, int sheet, InputStream[] templateNS, LangCode langCode, Map<String, Object> session) {
        if (is == null || templateNS == null || templateNS.length < 1)  return null;
        Map<Integer, String[]> map = ExcelUtil.readXLSX(is, sheet);
        if (map == null || map.size() < 1)  return null;
        return resolve(map, BeanUtil.getEntityBean(templateNS, map.get(0), langCode), langCode, session);
    }

    public static LeadinResult resolveXLSX(InputStream is, InputStream[] templateNS, LangCode langCode, Map<String, Object> session) {
        return resolveXLSX(is, 0, templateNS, langCode, session);
    }

    public static LeadinResult resolveXLS(InputStream is, int sheet, InputStream[] templateNS, LangCode langCode, OutputStream errorOut, Map<String, Object> session) {
        if (is == null || templateNS == null || templateNS.length < 1)  return null;
        ByteArrayOutputStream baos = null;
        try {
            baos = getBAOS(is);
            Map<Integer, String[]> map = ExcelUtil.readXLS(new ByteArrayInputStream(baos.toByteArray()), sheet);
            if (map == null || map.size() < 1)  return null;
            return resolve(new ByteArrayInputStream(baos.toByteArray()), sheet, map, BeanUtil.getEntityBean(templateNS, map.get(0), langCode), langCode, errorOut, session);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (baos != null) {
                try {
                    baos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return null;
    }

    public static LeadinResult resolveXLS(InputStream is, InputStream[] templateNS, LangCode langCode, OutputStream errorOut, Map<String, Object> session) {
        return resolveXLS(is, 0, templateNS, langCode, errorOut, session);
    }

    public static LeadinResult resolveXLSX(InputStream is, int sheet, InputStream[] templateNS, LangCode langCode, OutputStream errorOut, Map<String, Object> session) {
        if (is == null || templateNS == null || templateNS.length < 1)  return null;
        ByteArrayOutputStream baos = null;
        try {
            baos = getBAOS(is);
            Map<Integer, String[]> map = ExcelUtil.readXLSX(new ByteArrayInputStream(baos.toByteArray()), sheet);
            if (map == null || map.size() < 1)  return null;
            return resolve(new ByteArrayInputStream(baos.toByteArray()), sheet, map, BeanUtil.getEntityBean(templateNS, map.get(0), langCode), langCode, errorOut, session);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (baos != null) {
                try {
                    baos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return null;
    }

    public static LeadinResult resolveXLSX(InputStream is, InputStream[] templateNS, LangCode langCode, OutputStream errorOut, Map<String, Object> session) {
        return resolveXLSX(is, 0, templateNS, langCode, errorOut, session);
    }

    private static ByteArrayOutputStream getBAOS(InputStream is) {
        if (is == null)     return null;
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try {
            byte[] buffer = new byte[1024];
            int len;
            while ((len = is.read(buffer)) > -1) {
                baos.write(buffer, 0, len);
            }
            baos.flush();
        } catch (Exception e) {}
        return baos;
    }

}

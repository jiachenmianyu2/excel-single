package cn.com.flyfish;


import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.Collection;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * 2014.06.09 excel 操作类
 *
 * @author chenmy
 * @version 1.0
 */
public class ExcelMom<T> {
    /**
     * export excel by poi
     *
     * @param title
     * @param headers
     * @param dataset
     * @param out
     * @param pattern
     */
    @SuppressWarnings("unchecked")
    public void exportExcel(String title, String[] headers,
                            Collection<T> dataset, OutputStream out, String pattern) {
        HSSFWorkbook workbook = null;

        try {
            workbook = new HSSFWorkbook();

            HSSFSheet sheet = workbook.createSheet(title);

            sheet.setDefaultColumnWidth(15);

            HSSFCellStyle style = workbook.createCellStyle();

            HSSFFont font = workbook.createFont();

            style.setFont(font);

            HSSFRow row = sheet.createRow(0);

            for (int i = 0; i < headers.length; i++) {
                HSSFCell cell = row.createCell(i);
                cell.setCellStyle(style);
                HSSFRichTextString text = new HSSFRichTextString(headers[i]);
                cell.setCellValue(text);
            }
            //迭代数据
            Iterator<T> itor = dataset.iterator();
            int index = 0;
            while (itor.hasNext()) {
                index++;
                row = sheet.createRow(index);
                T entityBean = (T) itor.next();
                Field[] fields = entityBean.getClass().getDeclaredFields();
                //反射获得所有列
                for (int i = 0; i < fields.length; i++) {
                    HSSFCell cell = row.createCell(i);
                    Field field = fields[i];
                    String fieldName = field.getName();
                    String getMethodName = "get".concat(fieldName.substring(0, 1).toUpperCase()).concat(fieldName.substring(1));
                    Class _class = entityBean.getClass();
                    Method getMethod = _class.getMethod(getMethodName, new Class[]{});
                    Object value = getMethod.invoke(entityBean, new Object[]{});
                    value = value == null ? "" : value;
                    cell.setCellValue(value.toString());
                }
            }
            workbook.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if(out != null) out.close();
            } catch (Exception e2) {
                e2.printStackTrace();
            }
        }
    }
}

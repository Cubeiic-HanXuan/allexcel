package com.cubeiic.excel.reflection;

import com.cubeiic.excel.util.*;
//import com.cubeiic.excel.util.Const;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.BeanUtils;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author hanxuan
 * @date 2018/7/10 14:59
 * ClassName:ExcelUtil Function: excel快速读取、写入公共类
 * 只需要两步即可完成以前复杂的Excel读取 用法步骤： 1.定义需要读取的表头字段和表头对应的属性字段 String keyValue
 * ="手机名称:phoneName,颜色:color,售价:price"; 2.读取数据 List<PhoneModel> list =
 * ExcelUtil.readXls("D://test.xlsx",ExcelUtil.getMap(keyValue),"Phone");
 * @version V1.0
 * @since JDK 1.7
 * @see
 */
public class ReflectionExcel implements Serializable {

    private static final long serialVersionUID = 1L;


    /**
     * getMap:(将传进来的表头和表头对应的属性存进Map集合，表头字段为key,属性为value)
     * @author hanxuan
     * 把传进指定格式的字符串解析到Map中
     *  形如: String keyValue = "手机名称:phoneName,颜色:color,售价:price";
     * @return
     * @since JDK 1.7
     */
    public static Map<String, String> getMap(String keyValue) {
        Map<String, String> map = new HashMap<String, String>();
        if (keyValue != null) {
            String[] str = keyValue.split(",");
            for (String element : str) {
                String[] str2 = element.split(":");
                map.put(str2[0], str2[1]);
            }
        }
        return map;
    }

    /**
     * @author hanxuan
     * 把传进指定格式的字符串解析到List中
     * @return List
     * @Date 2018年5月9日 21:42:24
     * @since JDK 1.7
     */
    public static List<String> getList(String keyValue) {
        List<String> list = new ArrayList<String>();
        if (keyValue != null) {
            String[] str = keyValue.split(",");

            for (String element : str) {
                String[] str2 = element.split(":");
                list.add(str2[0]);
            }
        }
        return list;
    }

    /**
     * readXlsPart:(根据传进来的map集合读取Excel) 传进来4个参数 <String,String>类型，第二个要反射的类的具体路径)
     *
     * @author hanxuan
     * @param filePath
     *            Excel文件路径
     * @param map
     *            表头和属性的Map集合,其中Map中Key为Excel列的名称，Value为反射类的属性
     * @param classPath
     *            需要映射的model的路径
     * @return
     * @throws Exception
     * @since JDK 1.7
     */
    public static <T> List<T> readXlsPart(String filePath, Map map, String classPath, int... rowNumIndex)
            throws Exception {

        // 返回键的集合
        Set keySet = map.keySet();

        /** 反射用 **/
        Class<?> demo = null;
        Object obj = null;
        /** 反射用 **/

        List<Object> list = new ArrayList<Object>();
        demo = Class.forName(classPath);
        String fileType = filePath.substring(filePath.lastIndexOf(".") + 1, filePath.length());
        InputStream is = new FileInputStream(filePath);
        Workbook wb = null;

        if (fileType.equals("xls")) {
            wb = new HSSFWorkbook(is);
        } else if (fileType.equals("xlsx")) {
            wb = new XSSFWorkbook(is);
        } else {
            throw new Exception("您输入的excel格式不正确");
        }
        // 获取每个Sheet表
        for (int sheetNum = 0; sheetNum < 1; sheetNum++) {

            // 记录第x行为表头
            int rowNumX = -1;
            // 存放每一个field字段对应所在的列的序号
            Map<String, Integer> cellmap = new HashMap<String, Integer>();
            // 存放所有的表头字段信息
            List<String> headlist = new ArrayList();

            Sheet hssfSheet = wb.getSheetAt(sheetNum);

            // 设置默认最大行为2w行
            if (hssfSheet != null && hssfSheet.getLastRowNum() > 60000) {
                throw new Exception("Excel 数据超过60000行,请检查是否有空行,或分批导入");
            }

            // 循环行Row
            for (int rowNum = 0; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
                // 如果传值指定从第几行开始读，就从指定行寻找，否则自动寻找
                if (rowNumIndex != null && rowNumIndex.length > 0 && rowNumX == -1) {
                    Row hssfRow = hssfSheet.getRow(rowNumIndex[0]);
                    if (hssfRow == null) {
                        throw new RuntimeException("指定的行为空，请检查");
                    }
                    rowNum = rowNumIndex[0] - 1;
                }
                Row hssfRow = hssfSheet.getRow(rowNum);
                if (hssfRow == null) {
                    continue;
                }
                boolean flag = false;
                for (int i = 0; i < hssfRow.getLastCellNum(); i++) {
                    if (hssfRow.getCell(i) != null && !("").equals(hssfRow.getCell(i).toString().trim())) {
                        flag = true;
                    }
                }
                if (!flag) {
                    continue;
                }

                if (rowNumX == -1) {
                    // 循环列Cell
                    for (int cellNum = 0; cellNum <= hssfRow.getLastCellNum(); cellNum++) {

                        Cell hssfCell = hssfRow.getCell(cellNum);
                        if (hssfCell == null) {
                            continue;
                        }

                        String tempCellValue = hssfSheet.getRow(rowNum).getCell(cellNum).getStringCellValue();

                        tempCellValue = StringUtils.remove(tempCellValue, (char) 160);
                        tempCellValue = tempCellValue.trim();

                        headlist.add(tempCellValue);

                        Iterator it = keySet.iterator();

                        while (it.hasNext()) {
                            Object key = it.next();
                            if (StringUtils.isNotBlank(tempCellValue)
                                    && StringUtils.equals(tempCellValue, key.toString())) {
                                rowNumX = rowNum;
                                cellmap.put(map.get(key).toString(), cellNum);
                            }
                        }
                        if (rowNumX == -1) {
                            throw new Exception("没有找到对应的字段或者对应字段行上面含有不为空白的行字段");
                        }
                    }

                } else {
                    obj = demo.newInstance();
                    Iterator it = keySet.iterator();
                    while (it.hasNext()) {
                        Object key = it.next();
                        Integer cellNum_x = cellmap.get(map.get(key).toString());
                        if (cellNum_x == null || hssfRow.getCell(cellNum_x) == null) {
                            continue;
                        }
                        // 得到属性
                        String attr = map.get(key).toString();

                        Class<?> attrType = BeanUtils.findPropertyType(attr, new Class[] { obj.getClass() });

                        Cell cell = hssfRow.getCell(cellNum_x);
                        getValue(cell, obj, attr, attrType, rowNum, cellNum_x, key);

                    }
                    list.add(obj);
                }

            }
        }
        is.close();
        // wb.close();
        return (List<T>) list;
    }

    /**
     *
     *
     * @author hanxuan
     * @param filePath
     *            Excel文件路径
     * @param map
     *            表头和属性的Map集合,其中Map中Key为Excel列的名称，Value为反射类的属性
     * @param classPath
     *            需要映射的model的路径
     * @return
     * @throws Exception
     * @since JDK 1.7
     */
    public static <T> List<T> readXls(String filePath, Map map, String classPath, int... rowNumIndex) throws Exception {

        // 返回键的集合
        Set keySet = map.keySet();

        /** 反射用 **/
        Class<?> demo = null;
        Object obj = null;
        /** 反射用 **/

        List<Object> list = new ArrayList<Object>();
        demo = Class.forName(classPath);
        String fileType = filePath.substring(filePath.lastIndexOf(".") + 1, filePath.length());
        InputStream is = new FileInputStream(filePath);
        Workbook wb = null;

        if (fileType.equals("xls")) {
            wb = new HSSFWorkbook(is);
        } else if (fileType.equals("xlsx")) {
            wb = new XSSFWorkbook(is);
        } else {
            throw new Exception("您输入的excel格式不正确");
        }
        //默认循环所有sheet，如果rowNumIndex[]
        // 获取每个Sheet表
        for (int sheetNum = 0; sheetNum < wb.getNumberOfSheets(); sheetNum++) {

            // 记录第x行为表头
            int rowNum_x = -1;
            // 存放每一个field字段对应所在的列的序号
            Map<String, Integer> cellmap = new HashMap<String, Integer>();
            // 存放所有的表头字段信息
            List<String> headlist = new ArrayList();

            Sheet hssfSheet = wb.getSheetAt(sheetNum);

            // 循环行Row
            for (int rowNum = 0; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
                // 如果传值指定从第几行开始读，就从指定行寻找，否则自动寻找
                if (rowNumIndex != null && rowNumIndex.length > 0 && rowNum_x == -1) {
                    Row hssfRow = hssfSheet.getRow(rowNumIndex[0]);
                    if (hssfRow == null) {
                        throw new RuntimeException("指定的行为空，请检查");
                    }
                    rowNum = rowNumIndex[0] - 1;
                }
                Row hssfRow = hssfSheet.getRow(rowNum);
                if (hssfRow == null) {
                    continue;
                }
                boolean flag = false;
                for (int i = 0; i < hssfRow.getLastCellNum(); i++) {
                    if (hssfRow.getCell(i) != null && !("").equals(hssfRow.getCell(i).toString().trim())) {
                        flag = true;
                    }
                }
                if (!flag) {
                    continue;
                }

                if (rowNum_x == -1) {
                    // 循环列Cell
                    for (int cellNum = 0; cellNum <= hssfRow.getLastCellNum(); cellNum++) {

                        Cell hssfCell = hssfRow.getCell(cellNum);
                        if (hssfCell == null) {
                            continue;
                        }

                        String tempCellValue = hssfSheet.getRow(rowNum).getCell(cellNum).getStringCellValue();

                        tempCellValue = StringUtils.remove(tempCellValue, (char) 160);
                        tempCellValue = tempCellValue.trim();

                        headlist.add(tempCellValue);
                        Iterator it = keySet.iterator();
                        while (it.hasNext()) {
                            Object key = it.next();
                            if (StringUtils.isNotBlank(tempCellValue)
                                    && StringUtils.equals(tempCellValue, key.toString())) {
                                rowNum_x = rowNum;
                                cellmap.put(map.get(key).toString(), cellNum);
                            }
                        }
                        if (rowNum_x == -1) {
                            throw new Exception("没有找到对应的字段或者对应字段行上面含有不为空白的行字段");
                        }
                    }

                    // 读取到列后，检查表头是否完全一致--start
                    for (int i = 0; i < headlist.size(); i++) {
                        boolean boo = false;
                        Iterator itor = keySet.iterator();
                        while (itor.hasNext()) {
                            String tempname = itor.next().toString();
                            if (tempname.equals(headlist.get(i))) {
                                boo = true;
                            }
                        }
                        if (boo == false) {
                            throw new Exception("表头字段和定义的属性字段不匹配，请检查");
                        }
                    }

                    Iterator itor = keySet.iterator();
                    while (itor.hasNext()) {
                        boolean boo = false;
                        String tempname = itor.next().toString();
                        for (int i = 0; i < headlist.size(); i++) {
                            if (tempname.equals(headlist.get(i))) {
                                boo = true;
                            }
                        }
                        if (boo == false) {
                            throw new Exception("表头字段和定义的属性字段不匹配，请检查");
                        }
                    }
                    // 读取到列后，检查表头是否完全一致--end

                } else {
                    obj = demo.newInstance();
                    Iterator it = keySet.iterator();
                    while (it.hasNext()) {
                        Object key = it.next();
                        Integer cellNum_x = cellmap.get(map.get(key).toString());
                        if (cellNum_x == null || hssfRow.getCell(cellNum_x) == null) {
                            continue;
                        }
                        // 得到属性
                        String attr = map.get(key).toString();

                        Class<?> attrType = BeanUtils.findPropertyType(attr, new Class[] { obj.getClass() });

                        Cell cell = hssfRow.getCell(cellNum_x);
                        getValue(cell, obj, attr, attrType, rowNum, cellNum_x, key);

                    }
                    list.add(obj);
                }

            }
        }
        is.close();
        return (List<T>) list;
    }

    /**
     * readXls: 根据传进来的map集合读取Excel
     * @param file 文件流
     * @param map 表头和属性的Map集合,其中Map中Key为Excel列的名称，Value为反射类的属性
     * @param classPath 需要映射的model的路径
     * @param rowNumIndex 指定从 excel 的第几行开始扫描 不填写默认自动扫描读取文件
     * @param <T> 传入需要转换的具体实体对象 对象为 Class
     * @return
     * @author hanxuan
     * @throws Exception
     */
    public static <T> List<T> readXls(MultipartFile file , Map map, String classPath, int... rowNumIndex) throws Exception {

        // 返回键的集合
        Set keySet = map.keySet();

        /** 反射用 **/
        Class<?> reflection = null;
        Object obj = null;
        /** 反射用 **/

        List<Object> list = new ArrayList<Object>();
        reflection = Class.forName(classPath);

        // 文件上传路径
        String filePath = PathUtil.getClasspath() + Const.FILEPATHFILE;
        // 执行上传
        String fileName = FileUpload.fileUp(file, filePath, "excelFile");

        InputStream is = file.getInputStream();

        Workbook wb;

        if (fileName.endsWith("xls")) {
            wb = new HSSFWorkbook(is);
        } else if (fileName.endsWith("xlsx")) {
            wb = new XSSFWorkbook(is);
        } else {
            throw new Exception("您输入的excel格式不正确");
        }

        //默认循环所有sheet，如果rowNumIndex[]
        // 获取每个Sheet表
        for (int sheetNum = 0; sheetNum < wb.getNumberOfSheets(); sheetNum++) {

            // 记录第x行为表头
            int rowNum_x = -1;
            // 存放每一个field字段对应所在的列的序号
            Map<String, Integer> cellmap = new HashMap<String, Integer>();
            // 存放所有的表头字段信息
            List<String> headlist = new ArrayList();

            Sheet hssfSheet = wb.getSheetAt(sheetNum);

            // 循环行Row
            for (int rowNum = 0; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
                // 如果传值指定从第几行开始读，就从指定行寻找，否则自动寻找
                if (rowNumIndex != null && rowNumIndex.length > 0 && rowNum_x == -1) {
                    Row hssfRow = hssfSheet.getRow(rowNumIndex[0]);
                    if (hssfRow == null) {
                        throw new RuntimeException("指定的行为空，请检查");
                    }
                    rowNum = rowNumIndex[0] - 1;
                }
                Row hssfRow = hssfSheet.getRow(rowNum);
                if (hssfRow == null) {
                    continue;
                }
                boolean flag = false;
                for (int i = 0; i < hssfRow.getLastCellNum(); i++) {
                    if (hssfRow.getCell(i) != null && !("").equals(hssfRow.getCell(i).toString().trim())) {
                        flag = true;
                    }
                }
                if (!flag) {
                    continue;
                }

                if (rowNum_x == -1) {
                    // 循环列Cell
                    for (int cellNum = 0; cellNum <= hssfRow.getLastCellNum(); cellNum++) {

                        Cell hssfCell = hssfRow.getCell(cellNum);
                        if (hssfCell == null) {
                            continue;
                        }

                        String tempCellValue = hssfSheet.getRow(rowNum).getCell(cellNum).getStringCellValue();

                        tempCellValue = StringUtils.remove(tempCellValue, (char) 160);
                        tempCellValue = tempCellValue.trim();

                        headlist.add(tempCellValue);
                        Iterator it = keySet.iterator();
                        while (it.hasNext()) {
                            Object key = it.next();
                            if (StringUtils.isNotBlank(tempCellValue)
                                    && StringUtils.equals(tempCellValue, key.toString())) {
                                rowNum_x = rowNum;
                                cellmap.put(map.get(key).toString(), cellNum);
                            }
                        }
                        if (rowNum_x == -1) {
                            throw new Exception("没有找到对应的字段或者对应字段行上面含有不为空白的行字段");
                        }
                    }

                    // 读取到列后，检查表头是否完全一致--start
                    for (int i = 0; i < headlist.size(); i++) {
                        boolean boo = false;
                        Iterator itor = keySet.iterator();
                        while (itor.hasNext()) {
                            String tempname = itor.next().toString();
                            if (tempname.equals(headlist.get(i))) {
                                boo = true;
                            }
                        }
                        if (boo == false) {
                            throw new Exception("表头字段和定义的属性字段不匹配，请检查");
                        }
                    }

                    Iterator itor = keySet.iterator();
                    while (itor.hasNext()) {
                        boolean boo = false;
                        String tempname = itor.next().toString();
                        for (int i = 0; i < headlist.size(); i++) {
                            if (tempname.equals(headlist.get(i))) {
                                boo = true;
                            }
                        }
                        if (boo == false) {
                            throw new Exception("表头字段和定义的属性字段不匹配，请检查");
                        }
                    }
                    // 读取到列后，检查表头是否完全一致--end

                } else {
                    obj = reflection.newInstance();
                    Iterator it = keySet.iterator();
                    while (it.hasNext()) {
                        Object key = it.next();
                        Integer cellNum_x = cellmap.get(map.get(key).toString());
                        if (cellNum_x == null || hssfRow.getCell(cellNum_x) == null) {
                            continue;
                        }
                        // 得到属性
                        String attr = map.get(key).toString();

                        Class<?> attrType = BeanUtils.findPropertyType(attr, new Class[] { obj.getClass() });

                        Cell cell = hssfRow.getCell(cellNum_x);
                        getValue(cell, obj, attr, attrType, rowNum, cellNum_x, key);

                    }
                    list.add(obj);
                }

            }
        }
         is.close();
        return (List<T>) list;
    }


    /**
     * readXlsPart:(根据传进来的map集合读取Excel) 传进来4个参数 <String,String>类型，第二个要反射的类的具体路径)
     *
     * @author hanxuan
     * @param param filePath （Excel文件路径）
     *              map （表头和属性的Map集合,其中Map中Key为Excel列的名称，Value为反射类的属性）
     *              classPath （需要映射的model的路径）
     * @return
     * @throws Exception
     * @since JDK 1.7
     */
    public static <T> List<T> readXlsPart(ExcelParam param)
            throws Exception {

        // 返回键的集合
        Set keySet = param.getMap().keySet();

        /** 反射用 **/
        Class<?> demo = null;
        Object obj = null;
        /** 反射用 **/

        List<Object> list = new ArrayList<Object>();
        demo = Class.forName(param.getClassPath());
        String fileType = param.getFilePath().substring(param.getFilePath().lastIndexOf(".") + 1, param.getFilePath().length());
        InputStream is = new FileInputStream(param.getFilePath());
        Workbook wb = null;

        if (ExcelTypeEnum.EXCEL_THREE.getText().equals(fileType)) {
            wb = new HSSFWorkbook(is);
        } else if (ExcelTypeEnum.EXCEL_SEVEN.getText().equals(fileType)) {
            wb = new XSSFWorkbook(is);
        } else {
            throw new Exception("您输入的excel格式不正确");
        }
        int startSheetNum = 0;
        int endSheetNum = 1;
        if(null != param.getSheetIndex()){
            startSheetNum = param.getSheetIndex()-1;
            endSheetNum = param.getSheetIndex();
        }
        // 获取每个Sheet表
        for (int sheetNum = startSheetNum; sheetNum < endSheetNum; sheetNum++) {

            // 记录第x行为表头
            int rowNumX = -1;
            // 存放每一个field字段对应所在的列的序号
            Map<String, Integer> cellmap = new HashMap<String, Integer>();
            // 存放所有的表头字段信息
            List<String> headlist = new ArrayList();

            Sheet hssfSheet = wb.getSheetAt(sheetNum);

            // 设置默认最大行为2w行
            if (hssfSheet != null && hssfSheet.getLastRowNum() > 60000) {
                throw new Exception("Excel 数据超过60000行,请检查是否有空行,或分批导入");
            }

            // 循环行Row
            for (int rowNum = 0; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {

                // 如果传值指定从第几行开始读，就从指定行寻找，否则自动寻找
                if (param.getRowNumIndex() != null && rowNumX == -1) {
                    Row hssfRow = hssfSheet.getRow(param.getRowNumIndex());
                    if (hssfRow == null) {
                        throw new RuntimeException("指定的行为空，请检查");
                    }
                    rowNum = param.getRowNumIndex() - 1;
                }
                Row hssfRow = hssfSheet.getRow(rowNum);
                if (hssfRow == null) {
                    continue;
                }
                boolean flag = false;
                for (int i = 0; i < hssfRow.getLastCellNum(); i++) {
                    if (hssfRow.getCell(i) != null && !("").equals(hssfRow.getCell(i).toString().trim())) {
                        flag = true;
                    }
                }
                if (!flag) {
                    continue;
                }

                if (rowNumX == -1) {
                    // 循环列Cell
                    for (int cellNum = 0; cellNum <= hssfRow.getLastCellNum(); cellNum++) {

                        Cell hssfCell = hssfRow.getCell(cellNum);
                        if (hssfCell == null) {
                            continue;
                        }

                        String tempCellValue = hssfSheet.getRow(rowNum).getCell(cellNum).getStringCellValue();

                        tempCellValue = StringUtils.remove(tempCellValue, (char) 160);
                        tempCellValue = tempCellValue.trim();

                        headlist.add(tempCellValue);

                        Iterator it = keySet.iterator();

                        while (it.hasNext()) {
                            Object key = it.next();
                            if (StringUtils.isNotBlank(tempCellValue)
                                    && StringUtils.equals(tempCellValue, key.toString())) {
                                rowNumX = rowNum;
                                cellmap.put(param.getMap().get(key).toString(), cellNum);
                            }
                        }
                        if (rowNumX == -1) {
                            throw new Exception("没有找到对应的字段或者对应字段行上面含有不为空白的行字段");
                        }
                    }

                } else {
                    obj = demo.newInstance();
                    Iterator it = keySet.iterator();
                    while (it.hasNext()) {
                        Object key = it.next();
                        Integer cellNum_x = cellmap.get(param.getMap().get(key).toString());
                        if (cellNum_x == null || hssfRow.getCell(cellNum_x) == null) {
                            continue;
                        }
                        // 得到属性
                        String attr = param.getMap().get(key).toString();

                        Class<?> attrType = BeanUtils.findPropertyType(attr, new Class[] { obj.getClass() });

                        Cell cell = hssfRow.getCell(cellNum_x);
                        getValue(cell, obj, attr, attrType, rowNum, cellNum_x, key);

                    }
                    list.add(obj);
                }

            }
        }
        is.close();
        // wb.close();
        return (List<T>) list;
    }

    /**
     * setter:(反射的set方法给属性赋值)
     *
     * @author hanxuan
     * @param obj
     *            具体的类
     * @param att
     *            类的属性
     * @param value
     *            赋予属性的值
     * @param type
     *            属性是哪种类型 比如:String double boolean等类型
     * @throws Exception
     * @since JDK 1.7
     */
    public static void setter(Object obj, String att, Object value, Class<?> type, int row, int col, Object key)
            throws Exception {
        try {
            Method method = obj.getClass().getMethod("set" + StringUtil.toUpperCaseFirstOne(att), type);
            method.invoke(obj, value);
        } catch (Exception e) {
            throw new Exception("第" + (row + 1) + " 行  " + (col + 1) + "列   属性：" + key + " 赋值异常  "+e);
        }

    }

    /**
     * getAttrVal:(反射的get方法得到属性值)
     *
     * @author hanxuan
     * @param obj
     *            具体的类
     * @param att
     *            类的属性
     *        value
     *            赋予属性的值
     * @param type
     *            属性是哪种类型 比如:String double boolean等类型
     * @throws Exception
     * @since JDK 1.7
     */
    public static Object getAttrVal(Object obj, String att, Class<?> type) throws Exception {
        try {
            Method method = obj.getClass().getMethod("get" + StringUtil.toUpperCaseFirstOne(att));
            Object value = new Object();
            value = method.invoke(obj);
            return value;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }

    }

    /**
     * getValue:(得到Excel列的值)
     * @author hanxuan
     * @param cell
     * @param obj
     * @param attr
     * @param attrType
     * @param row
     * @param col
     * @param key
     * @throws Exception
     * @since JDK 1.7
     */
    public static void getValue(Cell cell, Object obj, String attr, Class attrType, int row, int col, Object key)
            throws Exception {
        Object val = null;
        if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
            val = cell.getBooleanCellValue();

        } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                try {
                    if (attrType == String.class) {
                        val = sdf.format(DateUtil.getJavaDate(cell.getNumericCellValue()));
                    } else {
                        val = dateConvertFormat(sdf.format(DateUtil.getJavaDate(cell.getNumericCellValue())));
                    }
                } catch (ParseException e) {
                    throw new Exception("第" + (row + 1) + " 行  " + (col + 1) + "列   属性：" + key + " 日期格式转换错误  ");
                }
            } else {
                if (attrType == String.class) {
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    val = cell.getStringCellValue();
                } else if (attrType == BigDecimal.class) {
                    val = new BigDecimal(cell.getNumericCellValue());
                } else if (attrType == long.class) {
                    val = (long) cell.getNumericCellValue();
                } else if (attrType == Double.class) {
                    val = cell.getNumericCellValue();
                } else if (attrType == Float.class) {
                    val = (float) cell.getNumericCellValue();
                } else if (attrType == int.class || attrType == Integer.class) {
                    val = (int) cell.getNumericCellValue();
                } else if (attrType == Short.class) {
                    val = (short) cell.getNumericCellValue();
                } else {
                    val = cell.getNumericCellValue();
                }
            }

        } else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
            val = cell.getStringCellValue();
        }

        setter(obj, attr, val, attrType, row, col, key);
    }

    /**
     * @author hanxuan
     * @param outFilePath 导出文件地址
     * @param keyValue 传入组装对象 如："年龄:age,生日:birthday"
     * @param list 对象数据集合
     * @param classPath 映射路径
     * @throws Exception
     */
    public static void exportExcel(String outFilePath, String keyValue, List<?> list, String classPath)
            throws Exception {

        Map<String, String> map = getMap(keyValue);
        List<String> keyList = getList(keyValue);

        Class<?> demo = null;
        demo = Class.forName(classPath);
        Object obj = demo.newInstance();
        // 创建HSSFWorkbook对象(excel的文档对象)
        HSSFWorkbook wb = new HSSFWorkbook();
        // 建立新的sheet对象（excel的表单）
        HSSFSheet sheet = wb.createSheet("sheet1");
        // 声明样式
        HSSFCellStyle style = wb.createCellStyle();
        // 居中显示
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // 在sheet里创建第一行为表头，参数为行索引(excel的行)，可以是0～65535之间的任何一个
        HSSFRow rowHeader = sheet.createRow(0);
        // 创建单元格并设置单元格内容

        // 存储属性信息
        Map<String, String> attMap = new HashMap();
        int index = 0;

        for (String key : keyList) {
            rowHeader.createCell(index).setCellValue(key);
            attMap.put(Integer.toString(index), map.get(key).toString());
            index++;
        }

        // 在sheet里创建表头下的数据
        for (int i = 0; i < list.size(); i++) {
            HSSFRow row = sheet.createRow(i+1);
            for (int j = 0; j < map.size(); j++) {

                Class<?> attrType = BeanUtils.findPropertyType(attMap.get(Integer.toString(j)),
                        new Class[] { obj.getClass() });

                Object value = getAttrVal(list.get(i), attMap.get(Integer.toString(j)), attrType);
                if(null==value){
                    value = "";
                }
                row.createCell(j).setCellValue(value.toString());
                style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            }
        }

        // 输出Excel文件
        try {
            FileOutputStream out = new FileOutputStream(outFilePath);
            wb.write(out);
            out.close();
        } catch (FileNotFoundException e) {
            throw new FileNotFoundException("导出失败！"+e);
        } catch (IOException e) {
            throw new IOException("导出失败！"+e);
        }

    }

    /**
     * exportExcel:(导出Excel)
     * @param response
     * @param keyValue 导出字段对应的keyValue
     * @param list 需要导出的数据列表
     * @param classPath 需要反射的实体对象 class
     * @param dateType 传入导出时间格式化类型 如：yyyy-MM-dd,yyyy/MM/dd
     * @param fileName sheet名称 可以传多个名称
     * @throws Exception
     */
    public static void exportExcelOutputStream(HttpServletResponse response, String keyValue, List<?> list, String classPath,String dateType,String... fileName)
            throws Exception {

        Map<String, String> map = getMap(keyValue);
        List<String> keyList = getList(keyValue);

        Class<?> demo = null;
        demo = Class.forName(classPath);
        Object obj = demo.newInstance();
        // 创建HSSFWorkbook对象(excel的文档对象)
        HSSFWorkbook wb = new HSSFWorkbook();
        // 建立新的sheet对象（excel的表单）
        HSSFSheet sheet = wb.createSheet("sheet1");
        // 声明样式
        HSSFCellStyle style = wb.createCellStyle();
        // 居中显示
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // 在sheet里创建第一行为表头，参数为行索引(excel的行)，可以是0～65535之间的任何一个
        HSSFRow rowHeader = sheet.createRow(0);
        // 创建单元格并设置单元格内容

        // 存储属性信息
        Map<String, String> attMap = new HashMap();
        int index = 0;

        for (String key : keyList) {
            rowHeader.createCell(index).setCellValue(key);
            attMap.put(Integer.toString(index), map.get(key).toString());
            index++;
        }

        // 在sheet里创建表头下的数据
        for (int i = 0; i < list.size(); i++) {
            HSSFRow row = sheet.createRow(i+1);
            for (int j = 0; j < map.size(); j++) {
                Class<?> attrType = BeanUtils.findPropertyType(attMap.get(Integer.toString(j)),
                        new Class[] { obj.getClass() });
                Object value = getAttrVal(list.get(i), attMap.get(Integer.toString(j)), attrType);
                if(null==value){
                    value = "";
                }
                /*时间格式化*/
                switch(dateType){
                    case "yyyy-MM-dd":
                        if (value instanceof Date){
                            SimpleDateFormat df = new SimpleDateFormat(dateType);
                            value = df.format(value);
                        }
                        break;
                    case "yyyy/MM/dd":
                        if (value instanceof Date){
                            SimpleDateFormat df = new SimpleDateFormat(dateType);
                            value = df.format(value);
                        }
                        break;
                    default:
                        if (value instanceof Date){
                            SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                            value = df.format(value);
                        }
                        break;
                }

                row.createCell(j).setCellValue(value.toString());
                style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            }
        }

        // 输出Excel文件
        try {
            SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            String newFileName = fileName[0];
            if(StringUtils.isEmpty(fileName[0])){
                newFileName = df.format(new Date());
            }
            OutputStream outstream = response.getOutputStream();
            response.reset();
            response.setHeader("Content-disposition",
                    "attachment; filename=" + new String(newFileName.getBytes(), "iso-8859-1") + ".xls");
            response.setContentType("application/x-download");
            wb.write(outstream);
            outstream.close();

        } catch (FileNotFoundException e) {
            throw new FileNotFoundException("导出失败！"+e);
        } catch (IOException e) {
            throw new IOException("导出失败！"+e);
        }

    }

    /**
     * String类型日期转为Date类型
     *
     * @param dateStr
     * @return
     * @throws ParseException
     * @throws Exception
     */
    public static Date dateConvertFormat(String dateStr) throws ParseException {
        Date date = new Date();
        SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        date = format.parse(dateStr);
        return date;
    }

    /**
     * hanxuan
     * 功能：判断字符串是否为日期格式
     * @param strDate 如："Wed Oct 12 2016 00:00:00 GMT+0800 (中国标准时间)"
     * @return
     */
    public static boolean isDate(String strDate) {

        boolean r = false;
        Date date = new Date();
        try {

            if (strDate.contains("CST")){
                SimpleDateFormat sdf = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
                date = (Date) sdf.parse(strDate);
            }else if (strDate.contains("GMT")){
                String dateString = strDate;
                dateString = dateString.replace("GMT", "").replaceAll("\\(.*\\)", "");
                //将字符串转化为date类型，格式2016-10-12
                SimpleDateFormat format =  new SimpleDateFormat("EEE MMM dd yyyy hh:mm:ss z",Locale.ENGLISH);
                date  = format.parse(dateString);
            }

            String formatStr = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(date);

            String regex = "^((\\d{2}(([02468][048])|([13579][26]))[\\-\\/\\s]?((((0?[13578])|(1[02]))[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])|(3[01])))|(((0?[469])|(11))[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])|(30)))|(0?2[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])))))|(\\d{2}(([02468][1235679])|([13579][01345789]))[\\-\\/\\s]?((((0?[13578])|(1[02]))[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])|(3[01])))|(((0?[469])|(11))[\\-\\/\\s]?((0?[1-9])|([1-2][0-9])|(30)))|(0?2[\\-\\/\\s]?((0?[1-9])|(1[0-9])|(2[0-8]))))))(\\s(((0?[0-9])|([1-2][0-3]))\\:([0-5]?[0-9])((\\s)|(\\:([0-5]?[0-9])))))?$";
            Pattern pattern = Pattern.compile(regex);
            Matcher m = pattern.matcher(formatStr);

            r = m.matches();

        } catch (ParseException e) {
            //e.printStackTrace();
            r = false;
        }
        return r;
    }

    public static void main(String[] args) {


        System.out.print("是否为时间："+isDate("Wed Oct 12 2016 00:00:00 GMT+0800 (中国标准时间)"));

    }


}
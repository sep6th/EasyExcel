package com.github.sep6th.util;

import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.github.sep6th.core.XmlCURD;
import com.github.sep6th.exception.ExcelException;

/** 
  * @ClassName : ExcelImportUtil
  * @Description : TODO(导入数据) 
  * @Author : Liuzy
  * @Date : 2017年11月30日 下午3:27:50 
  * @Version : V1.0  
  */

public class ExcelImportUtil {
    
    // 存储xml模板
	private static List<String> xmlColumnHeaderList = new ArrayList<String>();
	
	// 存储有问题的数据行号，反馈给用户 和 需要存储的合法数据
	private static Map<String,List<Object>> infoMap = new HashMap<String,List<Object>>();
    // 存储数据中必填项为空的行号
	private static List<Object> xyOfNullList = new ArrayList<Object>();
    // 存储数据中与数据库唯一标识重复的数据
	private static List<Object> xyOfRepList = new ArrayList<Object>();
    
    /**
     * 方法描述：  读取Excel
     * @param: inStream  通过 file.getInputStream() 获得, 其中file是 org.springframework.web.multipart.MultipartFile 实例.
     * @param: xmlSheetIndex  xml配置中模板sheet的序号, 从0开始.
     * @param: dataList  传入空的list, 去取合法数据.
     * @param: uniqueIdSet  从数据库查出唯一标识set集合 , 传入
     * @Author: Liuzy 
     * @Version: V1.0
     */
	@SuppressWarnings("deprecation")
    public static Map<String,List<Object>> readExcel(InputStream inStream, Integer xmlSheetIndex, List<List<Object>> dataList, Set<String> uniqueIdSet) {
	    SimpleDateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");
	    // 同时支持 Excel 2003、2007、2010
        Workbook workbook;
        // 存储导入 Excel 模板
        List<String> xlsColumnList = new ArrayList<String>();
        try {
            // 必填项
            boolean[] xmlColumnNotNullArr;
            // 存储xml 列名对应的ABC...
            char[] xslEngCodeArr;
            // 记录唯一标识列是第几列
            Integer xmlColumnUnique = null;
            
            workbook = WorkbookFactory.create(inStream);
            // Sheet 的数量
            int sheetCount = workbook.getNumberOfSheets(); 
            // 遍历每个 Sheet
            for (int i = 0; i < sheetCount; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                // 获取总行数
                int rowCount = sheet.getPhysicalNumberOfRows();
                //获取 excel 标题行对象
                Row rowOfTitle = sheet.getRow(2);
                
                // 获取总列数
                int cellCount = rowOfTitle.getPhysicalNumberOfCells();
                // 获取导入 excel 标题
                xlsColumnList.clear();
                for(int j = 0; j < cellCount; j++) {
                    xlsColumnList.add(rowOfTitle.getCell(j).getStringCellValue());
                }
                
                // xml 列名个数（即 总列数）
                Integer columnCount = XmlCURD.getColumns(xmlSheetIndex).getChildrenCount("column");
                if(columnCount == null || "".equals(columnCount)) {
                    throw new ExcelException("获取不到xml中[sheet]序号为" + xmlSheetIndex + ",[column]的个数！");
                }
                // 必填项数组  实例化
                xmlColumnNotNullArr = new boolean[columnCount];
                xslEngCodeArr = new char[columnCount];
                // 获取 xml column 信息
                for(int j = 0; j < columnCount; j++) {
                    if(xmlColumnHeaderList.isEmpty()) {
                        
                        // 获取标题
                        String columnName = XmlCURD.getXmlStringValue("sheets.sheet(" + xmlSheetIndex + ").columns.column(" + j + ").header");
                        if(columnName == null || "".equals(columnName)){
                            throw new ExcelException("获取不到xml中[sheet]序号为" + xmlSheetIndex + ",[column]序号为" + j + "[header]的值！");
                        }
                        xmlColumnHeaderList.add(columnName);
                        
                        // 获取必填列
                        String notNull = XmlCURD.getXmlStringValue("sheets.sheet(" + xmlSheetIndex + ").columns.column(" + j + ").notNull");
                        if("true".equals(notNull)) {
                            xmlColumnNotNullArr[j] = true;
                        } else {
                            xmlColumnNotNullArr[j] = false;
                        }
                        
                        // 获取唯一标识列是第几列
                        String unique = XmlCURD.getXmlStringValue("sheets.sheet(" + xmlSheetIndex + ").columns.column(" + j + ").unique");
                        if("true".equals(unique)) {
                            xmlColumnUnique = j;
                        }
                        
                        // 获取xml 列名对应的ABC...
                        xslEngCodeArr[j] = (char)(j + (int)'A');
                        
                    }
                }
                if(!compare(xlsColumnList, xmlColumnHeaderList)) {
                    throw new ExcelException("模板不匹配！请检查模板是否正确。");
                }
                
                // 获取数据总行数
                int rowDataCount = rowCount - 3;
                if(rowDataCount <= 0) {
                    continue;
                }
                // 从数据行开始，遍历每一行
                for (int r = 3; r < rowDataCount; r++) {
                    Row row = sheet.getRow(r);
                    List<Object> cellValue = new ArrayList<Object>();
                    // 遍历每一列,列数以标题行为准
                    for(int c = 0; c < cellCount; c++) {
                        Cell cell = row.getCell(c);
                        
                        // 默认单元格的类型为空
                        @SuppressWarnings({ "unused" })
                        int cellType = Cell.CELL_TYPE_BLANK;
                        if(cell!=null) {
                            cellType = cell.getCellType();
                        
                            switch (cell.getCellType()) {
                                 // 文本
                                case Cell.CELL_TYPE_STRING:
                                    cellValue.add(cell.getStringCellValue());
                                    break;
                                 // 数字、日期
                                case Cell.CELL_TYPE_NUMERIC:
                                    if (DateUtil.isCellDateFormatted(cell)) {
                                        try{
                                            // 日期型
                                            cellValue.add(fmt.format(cell.getDateCellValue()));
                                        }catch(Exception e){
                                            cellValue.add(null);
                                        }
                                    } else {
                                        Double data = cell.getNumericCellValue();
                                        if(data==null||"".equals(data)){
                                            cellValue.add(0);
                                        }else{
                                            // 数字
                                            cellValue.add(String.valueOf(new DecimalFormat("#").format(cell.getNumericCellValue())));
                                        }
                                    }
                                    break;
                                 // 布尔型
                                case Cell.CELL_TYPE_BOOLEAN:
                                    cellValue.add(String.valueOf(cell.getBooleanCellValue()));
                                    break;
                                 // 公式
                                case Cell.CELL_TYPE_FORMULA:
                                    cellValue.add(null);
                                    break;
                                 // 空白
                                case Cell.CELL_TYPE_BLANK:
                                    cellValue.add(null);
                                    break;
                                 // 错误
                                case Cell.CELL_TYPE_ERROR:
                                    cellValue.add(null);
                                    break;
                                default:
                                    cellValue.add(null);
                            }
                        }
                        
                    }
                    // 用于判断是否是存进合法数据的 list 里
                    boolean flag = true;
                    // 获取了一条数据
                    if(!cellValue.isEmpty() && cellValue.size() == cellCount) {
                        for(int k = 0; k < cellCount; k++) {
                            // 启动必填项校验
                            if(xmlColumnNotNullArr[k]) {
                                // 记录为空的单元格
                                if(cellValue.get(k) == null || "".equals(cellValue.get(k))) {
                                    xyOfNullList.add(xslEngCodeArr[k] + r);
                                    flag = false;
                                }
                                
                            }
                        
                        }
                        // 启动唯一校验
                        if(xmlColumnUnique != null){
                            Object object = cellValue.get(xmlColumnUnique);
                            if(object != null){
                                // 利用set存不重复的值的特性
                                int size = uniqueIdSet.size();
                                uniqueIdSet.add((String)object);
                                if(uniqueIdSet.size() == size){
                                    //与数据库唯一标识重复，不存数据！记录单元格
                                    xyOfRepList.add(xslEngCodeArr[xmlColumnUnique] + r);
                                    flag = false;
                                }
                            
                            }
                            
                        }
                    }
                    // 存储没有问题的数据
                    if(flag){
                        dataList.add(cellValue);
                    }
                }
                
            }
            
        } catch (Exception e) {
            e.printStackTrace();
            throw new ExcelException("读取失败！");
        }
        infoMap.put("xyOfNullList", xyOfNullList);
        infoMap.put("xyOfRepList", xyOfRepList);
        return infoMap;
	}
	
	
	/**
	 * 
	 * 方法描述：  比较两个list存储的元素、个数及顺序是否一致
	 * @Author: Liuzy 
	 * @Version: V1.0
	 */
	private static boolean compare(List<String> a,List<String> b){
	    int size = a.size();
	    if(size != b.size()){
	        return false;
	    }
	    for(int i = 0; i < size; i++ ){
	       if(!a.get(i).equals(b.get(i))){
	           return false;
	       }
	    }
	    return true;
	}
	
}

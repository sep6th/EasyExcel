package com.github.sep6th.util;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;

import com.github.sep6th.constant.MsgConstant;
import com.github.sep6th.core.XmlCURD;
import com.github.sep6th.exception.ExcelException;

/** 
  * @ClassName : ExcelExportUtil
  * @Description : TODO (导出数据，或模板) 
  * @Author : Liuzy
  * @Date : 2017年11月14日 上午10:25:50 
  * @Version : V1.0  
  */

public class ExcelExportUtil {

	// 导出Excel的文件名
	private static String fileName;
	
	static{
		fileName = XmlCURD.getXmlStringValue("fileName");
	}

	
	/**
	 * 方法描述：  List<Object[]> 为null 时，导出数据模板； 有数据时，导出数据。
	 * Map<Integer,List<Object[]>> 的 key : xml 里 sheet 序号(从0开始); value : 对应 sheet 的数据。
	 * @Author: Liuzy 
	 * @Version: V1.0
	 */
	@SuppressWarnings({ "deprecation", "resource" })
	public static void export(Map<Integer,List<Object[]>> map, HttpServletResponse response) 
			throws Exception {
		// 创建工作簿对象
		HSSFWorkbook workbook = new HSSFWorkbook(); 
		if(map == null){
			throw new ExcelException(MsgConstant.PARAM_XMLSHEETINDEX_ERROR);
		}
		// 创建工作表
		HSSFSheet sheet = null;
		int columnCount = 0;
		for(Integer xmlSheetIndex : map.keySet()){
			
			String title = XmlCURD.getXmlStringValue("sheets.sheet("+xmlSheetIndex+").title");
			if("".equals(title) || title == null){
				title = fileName;
			}
			// title中含有0x0000 、 0x0003、冒号 (:)、反斜杠 (\)、星号 (*)、问号 (？ )、正斜杠 (/)、 [、 ]用空格代替
			String safeName = WorkbookUtil.createSafeSheetName(title);
			// 创建工作表
			sheet = workbook.createSheet(safeName); 
			// 创建行（标题）
			HSSFRow rowm = sheet.createRow(0);
			// 创建单元格（标题）
			HSSFCell cellTiltle = rowm.createCell(0);
			
			// 导出表的列数
			columnCount = XmlCURD.getColumns(xmlSheetIndex).getChildrenCount("column");
			if(columnCount == 0){
				throw new ExcelException(MsgConstant.MODEL_ROWNAME_ERROR);
			}
			
			// 设置标题的样式
			HSSFCellStyle columnTopStyle = getColumnTopStyle(workbook);
			HSSFCellStyle style = getStyle(workbook); 
			// 合并单元格CellRangeAddress方法参数
			// [第一行 (基于 0),最后一行 (0-based),第一列 (基于 0),最后一列 (从 0 开始 )]
			sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, (columnCount - 1)));
			String freeze = XmlCURD.getXmlStringValue("sheets.sheet("+xmlSheetIndex+").freeze");
			if(freeze!=null){
				if(freeze.equals("true")){
					// colSplit:冻结的列数、rowSplit:冻结的行数、leftmostColumn:右边区域[可见]的首列序号 、topRow:下边区域[可见]的首行序号
					sheet.createFreezePane(columnCount, 3, columnCount, 4);;
				}
			}
			cellTiltle.setCellStyle(columnTopStyle);
			// 将标题名称写入单元格
			cellTiltle.setCellValue(title);
			
			
			// 创建行（列名）
			HSSFRow rowRowName = sheet.createRow((short)2); 
			
			// 创建单元格（列名）,并将列名写入
			for (int n = 0; n < columnCount; n++) {
				HSSFCell cellRowName = rowRowName.createCell(n);
				// 设置单元格值的类型
				cellRowName.setCellType(CellType.STRING);
				//获取列名
				HSSFRichTextString text = new HSSFRichTextString(
				        XmlCURD.getXmlStringValue("sheets.sheet(" + xmlSheetIndex + ").columns.column(" + n + ").header"));
				cellRowName.setCellValue(text);
				// 设置单元格样式
				cellRowName.setCellStyle(columnTopStyle); 
			}
			// 获得要导出的数据
			List<Object[]> dataList = map.get(xmlSheetIndex);
			if(dataList != null && dataList.size() != 0){
				// 获取数据及数据条数，创建相应的行，并填充数据
				for (int i = 0; i < map.get(xmlSheetIndex).size(); i++) {
					// 遍历每个对象
					Object[] obj = dataList.get(i);
					// 创建行（数据）
					HSSFRow row = sheet.createRow(i + 3);
	
					for (int j = 0; j < obj.length; j++) {
						HSSFCell cell = null;
						cell = row.createCell(j, CellType.STRING);
						if (!"".equals(obj[j]) && obj[j] != null) {
							// 设置单元格的值
							cell.setCellValue(obj[j].toString());
						}else{
							cell.setCellValue("");
						}
						cell.setCellStyle(style);
					}
				}
			}
			
			// 让列宽随着导出的列长自动适应
			for (int colNum = 0; colNum < columnCount; colNum++) {
				int columnWidth = sheet.getColumnWidth(colNum) / 256;
				for (int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++) {
					HSSFRow currentRow;
					// 当前行未被使用过
					if (sheet.getRow(rowNum) == null) {
						currentRow = sheet.createRow(rowNum);
					} else {
						currentRow = sheet.getRow(rowNum);
					}
					if (currentRow.getCell(colNum) != null) {
						HSSFCell currentCell = currentRow.getCell(colNum);
						if (currentCell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
							int length = currentCell.getStringCellValue().getBytes().length;
							if (columnWidth < length) {
								columnWidth = length;
							}
						}
					}
				}
				
				if (colNum == 0) {
					sheet.setColumnWidth(colNum, (columnWidth - 2) * 256);
				} else {
					sheet.setColumnWidth(colNum, (columnWidth + 4) * 256);
				}
			}
			
			FileOutputStream fileOut = new FileOutputStream("workbook.xls");
			workbook.write(fileOut);
		    fileOut.close();
			if (workbook != null) {
				try {
					String fileFullName = fileName + ".xls";
					fileFullName = new String(fileFullName.getBytes("GBK"), "ISO_8859_1");
					response.setContentType("application/octet-stream;charset=GBK");
					response.setHeader("Content-Disposition", "attachment; filename=" + fileFullName);
					OutputStream out = response.getOutputStream();
					workbook.write(out);
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			
		}
		
	}
	
	/**
	 * 列头单元格样式
	 * @param workbook
	 * @return
	 */
	public static HSSFCellStyle getColumnTopStyle(HSSFWorkbook workbook) {

		// 设置字体
		HSSFFont font = workbook.createFont();
		// 设置字体大小
		font.setFontHeightInPoints((short) 11);
		// 字体加粗
		font.setBold(true);
		// 设置字体名字
		font.setFontName("Courier New");
		// 设置样式;
		HSSFCellStyle style = workbook.createCellStyle();
		// 设置底边框;
		style.setBorderBottom(BorderStyle.THIN);
		// 设置底边框颜色;
		style.setBottomBorderColor(HSSFColorPredefined.BLACK.getIndex());
		// 设置左边框;
		style.setBorderLeft(BorderStyle.THIN);
		// 设置左边框颜色;
		style.setLeftBorderColor(HSSFColorPredefined.BLACK.getIndex());
		// 设置右边框;
		style.setBorderRight(BorderStyle.THIN);
		// 设置右边框颜色;
		style.setRightBorderColor(HSSFColorPredefined.BLACK.getIndex());
		// 设置顶边框;
		style.setBorderTop(BorderStyle.THIN);
		// 设置顶边框颜色;
		style.setTopBorderColor(HSSFColorPredefined.BLACK.getIndex());
		// 在样式用应用设置的字体;
		style.setFont(font);
		// 设置自动换行;
		style.setWrapText(false);
		// 设置水平对齐的样式为居中对齐;
		style.setAlignment(HorizontalAlignment.CENTER);
		// 设置垂直对齐的样式为居中对齐;
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		return style;

	}

	
	/**
	 * 列数据信息单元格样式
	 * @param workbook
	 * @return
	 */
	public static HSSFCellStyle getStyle(HSSFWorkbook workbook) {
		// 设置字体
		HSSFFont font = workbook.createFont();
		// 设置字体大小
		font.setFontHeightInPoints((short)10);
		// 设置字体名字
		font.setFontName("Courier New");
		// 设置样式;
		HSSFCellStyle style = workbook.createCellStyle();
		// 设置底边框;
		style.setBorderBottom(BorderStyle.THIN);
		// 设置底边框颜色;
		style.setBottomBorderColor(HSSFColorPredefined.BLACK.getIndex());
		// 设置左边框;
		style.setBorderLeft(BorderStyle.THIN);
		// 设置左边框颜色;
		style.setLeftBorderColor(HSSFColorPredefined.BLACK.getIndex());
		// 设置右边框;
		style.setBorderRight(BorderStyle.THIN);
		// 设置右边框颜色;
		style.setRightBorderColor(HSSFColorPredefined.BLACK.getIndex());
		// 设置顶边框;
		style.setBorderTop(BorderStyle.THIN);
		// 设置顶边框颜色;
		style.setTopBorderColor(HSSFColorPredefined.BLACK.getIndex());
		// 在样式用应用设置的字体;
		style.setFont(font);
		// 设置自动换行;
		style.setWrapText(false);
		// 设置水平对齐的样式为居中对齐;
		style.setAlignment(HorizontalAlignment.CENTER);
		// 设置垂直对齐的样式为居中对齐;
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		return style;

	}
	
	public static void main(String[] args) throws Exception {
		Map<Integer,List<Object[]>> info = new HashMap<Integer, List<Object[]>>();
		List<Object[]> dataList = new ArrayList<Object[]>();
		Object[] a = new Object[]{"小白","",10.2,39};
		Object[] b = new Object[]{"小黑","",11.6,30};
		dataList.add(a);
		dataList.add(a);
		dataList.add(b);
		info.put(0, dataList);
		info.put(1, dataList);
		ExcelExportUtil.export(info, null);
	}
	
	
}

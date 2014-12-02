package office;

import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class OfficeUtils {

	public void createFile(String filePath) {
		String excelFile = "myexcel.xls";
		FileOutputStream fos = new FileOutputStream(excelFile);
		HSSFWorkbook wb = new HSSFWorkbook();// 创建工作薄
		HSSFSheet sheet = wb.createSheet();// 创建工作表
		wb.setSheetName(0, "sheet0");// 设置工作表名

		HSSFRow row = null;
		HSSFCell cell = null;
		for (int i = 0; i < 10; i++) {
			row = sheet.createRow(i);// 新增一行
			cell = row.createCell((short) 0);// 新增一列
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);// 设置单元格的字符集
			cell.setCellType(i);// 向单元格中写入数据
			cell = row.createCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("第" + i + "行");
		}
		wb.write(fos);
		fos.close();
	}

	public void readFile(String filePath, OutputStream out) {

	}

	public static void main(String[] args) {

	}
}
package way.tool;

import java.awt.image.BufferStrategy;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.Value;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class InputExcelTranslate {
	
	//翻譯Excel讀取位置
	static String inputExcelPath = "C:\\Users\\way82\\Desktop\\Tool\\excel\\測試功能.xlsx";
	
	
	public static void main(String[] args) throws IOException {
		Workbook wb = null;
		String extString = inputExcelPath.substring(inputExcelPath.lastIndexOf("."));
		InputStream is = new FileInputStream(inputExcelPath);
		if(".xls".equals(extString)){
		   wb = new HSSFWorkbook(is);
		}else if(".xlsx".equals(extString)){
		   wb = new XSSFWorkbook(is);
		}
		is.close();
		
		List<ArrayList<String>> excelData = parseExcel(wb);
		for(int index = 0; index < excelData.size(); index ++) {
			transateJpWord(excelData.get(index).get(0), excelData.get(index).get(1), excelData.get(index).get(2));
		}
		
	}

	public static List<ArrayList<String>> parseExcel(Workbook workbook) {
		List<ArrayList<String>> excelDataList = new ArrayList<ArrayList<String>>();
		//遍歷每一個sheet
		for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
			Sheet sheet = workbook.getSheetAt(sheetNum);

			if (sheet == null) {
				continue;
			}

			int firstRowNum = sheet.getFirstRowNum();
			Row firstRow = sheet.getRow(firstRowNum);
			if (null == firstRow) {
				System.out.println("解析Excel失敗");
			}

			int rowStart = firstRowNum + 1;
			int rowEnd = sheet.getPhysicalNumberOfRows();
			for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
				Row row = sheet.getRow(rowNum);
				if (null == row) {
					continue;
				}
				ArrayList<String> excelData = convertRowToData(row);
				if (null == excelData) {
					continue;
				}
				excelDataList.add(excelData);
			}
		}
		return excelDataList;
	}
	
	private static ArrayList<String> convertRowToData(Row row) {
		ArrayList<String> rowData = new ArrayList<String>();
		Cell cell;
		int cellNum = 0;
		String a = "";
		for(int i = 0; i < 3; i++) {
		cell = row.getCell(cellNum++);
		if(cell == null) {
			System.err.println("譯文欄位異常，不可為空");
		}
		cell.setCellType(CellType.STRING);
		rowData.add(cell.getStringCellValue());	
		}
		return rowData;
	}

	static public void transateJpWord(String jpWord, String twWord, String path) throws IOException {	
		System.out.println("開始讀取:"+path);
		//方法1
		File file1=new File(path);
		byte[] b=new byte[(int)file1.length()];
		FileInputStream in=null;
		FileOutputStream out=null;
		in=new FileInputStream(file1);

		byte[] buffer = new byte[1024];
		int len = 0;
		ArrayList<byte[]> container = new ArrayList<byte[]>();
		while((len = in.read(buffer)) != -1) {  
			if(new String(buffer).contains(jpWord)) {
				container.add(new String(buffer).replace(jpWord, twWord).getBytes());
			}else {
				container.add(new String(buffer, 0, len).getBytes());
			}
		}
		in.close();
		System.out.println("開始寫入");
		File file2=new File(path);
		out=new FileOutputStream(file2);
		for(int i = 0; i < container.size() ; i++) {
			out.write(container.get(i));
		}
			
		out.flush();
		out.close();
		System.out.println("寫入成功");
	    }
	
}

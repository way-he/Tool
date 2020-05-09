package way.tool;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CCBOutputWordToExcel {
	
	//CCB檔案路徑，起始需為資料夾
	static String inputResourcePath = "C:\\Users\\way82\\Desktop\\Tool\\excel\\123\\SpriteBuilder Resources";
	
	//Excel輸出位置
	static String outputExcelPath = "C:\\Users\\way82\\Desktop\\Tool\\excel\\測試功能.xlsx";
	
	static String parserSpecWord = "wayAddRule";		
	static Workbook wb = null; 	
	static Sheet sheet = null;
	static Row row = null;
	static Cell cell = null;
	static ArrayList<String> jpWord;
	static public void pathFile(String path) throws IOException {
		File file = new File(path);
		String[] filenames;
		String fullpath = file.getAbsolutePath(); // 取得路徑
		if (file.isDirectory()){
			filenames = file.list();
			for (int i = 0; i < filenames.length; i++){
					File tempFile = new File(fullpath + "\\" + filenames[i]);
					if (tempFile.isDirectory()){
						pathFile(fullpath + "\\" +filenames[i]);
					}else {
						//只處理ccb檔案
						if(!filenames[i].substring(filenames[i].length()-3).equals("ccb")) {
							continue;
						}
						parserWord(fullpath + "\\" +filenames[i]);
					}
	
				}
		}else {
			System.out.println("[" + file + "]不是目錄");
		}
		
	}
	
	public static void main(String[] args) throws IOException {
		jpWord = new ArrayList<String>();
		pathFile(inputResourcePath);   
        //excel生成
        String extString = outputExcelPath.substring(outputExcelPath.lastIndexOf("."));
        if(".xls".equals(extString)){
            wb = new HSSFWorkbook();
        }else if(".xlsx".equals(extString)){
            wb = new XSSFWorkbook();
        }else{
            System.out.println("無效檔案");
            return;
        }
        sheet = wb.createSheet();
        //表格寬度調整
        sheet.setColumnWidth(0, 150 * 125);
        sheet.setColumnWidth(1, 150 * 125);
        sheet.setColumnWidth(2, 150 * 125); 
        //標題
        String[] titleStrings = new String[] {"原文","譯文","路徑(不可修改)"};
        saveData(titleStrings);
        for(int listIndex = 0 ; listIndex < jpWord.size(); listIndex++) {
        	saveData(new String[] {jpWord.get(listIndex), "", jpWord.get(++listIndex)});     	
        }     
        //輸出excel
        FileOutputStream fos = new FileOutputStream(new File(outputExcelPath));
        wb.write(fos);
        fos.flush();
        fos.close();
    	
    }
	static boolean isNotAll = false;
	static public void parserWord(String inputResourcePath) throws IOException {
		Pattern r = Pattern.compile("<string>[\\x{2E80}-\\x{9FFF}]+?.*?</string>");
		ArrayList<String> ccbWord = new ArrayList<String>();
        InputStreamReader isr = new InputStreamReader(new FileInputStream(inputResourcePath), "UTF-8");
        BufferedReader br = new BufferedReader(isr);
        StringBuilder lineBuffer = new StringBuilder();
        while (br.ready()) {
        	String line = br.readLine();
        	if(!isNotAll && (line.contains("<string>") && !line.contains("</string>"))) {
        		
        		isNotAll = true;
        	}
        	if(isNotAll && (line.toString().contains("</string>"))) {
        			lineBuffer.append(parserSpecWord);
        			lineBuffer.append(line);
        			ccbWord.add(lineBuffer.toString()); 
        			System.out.println(lineBuffer.toString());
        			lineBuffer.setLength(0);
        			isNotAll = false;
        	}else if(isNotAll && (!line.toString().contains("</string>"))){ 		
        		lineBuffer.append(line);
        	}else {
        		ccbWord.add(line);
        	}
        	
        }
        isr.close();

        for(int i = 0 ; i<ccbWord.size(); i++) {
        	Matcher m = r.matcher(ccbWord.get(i));
            if (m.find( )) { 
            	String splitLine = m.group();
            	if(splitLine.contains(parserSpecWord)) {
            		String[] lineArray = splitLine.split(parserSpecWord);
            		for(int index = 0; index < lineArray.length; index++) {
            			saveDataToJpwordContainer(lineArray[index], inputResourcePath, String.valueOf(i+1));
            		}
            	}else {
            		saveDataToJpwordContainer(splitLine, inputResourcePath, String.valueOf(i+1));
            	}
             } 

        }
	}
	
	private static void saveDataToJpwordContainer(String string, String path, String wordLine) {
		System.out.println(string);
		jpWord.add(string.replace("<string>", "").replace("</string>", ""));
    	jpWord.add(path);
//    	jpWord.add(wordLine); 取消記錄行數
	}
	
	
	static int mRowNumber = 0;
	static boolean isFirstRow = true;
	static public void saveData(String[] data) {
		
		CellStyle firstRowStyle = null;
        if(isFirstRow) {
            firstRowStyle = wb.createCellStyle();
    	   	Font font = wb.createFont();
    	   	font.setColor(HSSFColor.GREEN.index);
    	   	firstRowStyle.setFont(font);
    	   	firstRowStyle.setLocked(true);
            isFirstRow = false;
        }
		
		CellStyle lockstyle =  wb.createCellStyle();
        lockstyle.setLocked(true);//设置锁定
        
        CellStyle unlockStyle= wb.createCellStyle();
        unlockStyle.setLocked(false);

	        row = sheet.createRow(mRowNumber);
	        for(int i = 0; i < data.length; i++) {
		        cell = row.createCell(i);
		        cell.setCellValue(data[i]);
		        
		        if(mRowNumber == 0) {
		        	cell.setCellStyle(firstRowStyle);
		        }else{
			        if(i >1 ) {
			            cell.setCellStyle(lockstyle);
			        }else {
			        	cell.setCellStyle(unlockStyle);	        	
			        }
		        }

	        }
	        //上鎖需要設置密碼
//	        sheet.protectSheet("123456");
	        mRowNumber++;
	}
	
}

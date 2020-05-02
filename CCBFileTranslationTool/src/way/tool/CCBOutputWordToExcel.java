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

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CCBOutputWordToExcel {
	//CCB�ɮ׸��|�A�_�l�ݬ���Ƨ�
	static String inputResourcePath = "C:\\Users\\way82\\Desktop\\���";
	//Excel��X��m
	static String outputExcelPath = "C:\\Users\\way82\\Desktop\\���\\ccb½Ķ���.xlsx";
	static Workbook wb = null; 
	
	static Sheet sheet = null;
	static Row row = null;
	static Cell cell = null;
	static ArrayList<String> jpWord;
	static public void pathFile(String path) throws IOException {
		File file = new File(path);
		String[] filenames;
		String fullpath = file.getAbsolutePath(); // ���o���|
		if (file.isDirectory()){
			filenames = file.list();
			for (int i = 0; i < filenames.length; i++){
					File tempFile = new File(fullpath + "\\" + filenames[i]);
					if (tempFile.isDirectory()){
						pathFile(fullpath + "\\" +filenames[i]);
					}else {
						//�u�B�zccb�ɮ�
						if(!filenames[i].substring(filenames[i].length()-3).equals("ccb")) {
							continue;
						}
						parserWord(fullpath + "\\" +filenames[i]);
					}
	
				}
		}else {
			System.out.println("[" + file + "]���O�ؿ�");
		}
		
	}
	
	public static void main(String[] args) throws IOException {
		jpWord = new ArrayList<String>();
		pathFile(inputResourcePath);   
        //excel�ͦ�
        String extString = outputExcelPath.substring(outputExcelPath.lastIndexOf("."));
        if(".xls".equals(extString)){
            wb = new HSSFWorkbook();
        }else if(".xlsx".equals(extString)){
            wb = new XSSFWorkbook();
        }else{
            System.out.println("�L���ɮ�");
            return;
        }
        sheet = wb.createSheet();
        //���e�׽վ�
        sheet.setColumnWidth(0, 150 * 125);
        sheet.setColumnWidth(1, 150 * 125);
        sheet.setColumnWidth(2, 150 * 125); 
        //���D
        String[] titleStrings = new String[] {"���","Ķ��","���|","���"};
        saveData(titleStrings);
        for(int listIndex = 0 ; listIndex < jpWord.size(); listIndex++) {
        	System.out.println(jpWord.get(listIndex));
        	saveData(new String[] {jpWord.get(listIndex), "", jpWord.get(++listIndex), jpWord.get(++listIndex)});     	
        }     
        //��Xexcel
        FileOutputStream fos = new FileOutputStream(new File(outputExcelPath));
        wb.write(fos);
        fos.flush();
        fos.close();
    	
    }
	
	static public void parserWord(String inputResourcePath) throws IOException {
		Pattern r = Pattern.compile("<string>[\\x{2E80}-\\x{9FFF}]+?.*?</string>");
		ArrayList<String> ccbWord = new ArrayList<String>();
        InputStreamReader isr = new InputStreamReader(new FileInputStream(inputResourcePath), "UTF-8");
        BufferedReader br = new BufferedReader(isr);
        while (br.ready()) {
        	ccbWord.add(br.readLine());
        }
        isr.close();
        for(int i = 0 ; i<ccbWord.size(); i++) {

        	Matcher m = r.matcher(ccbWord.get(i));
            if (m.find( )) { 
            	jpWord.add(m.group().replace("<string>", "").replace("</string>", ""));
            	jpWord.add(inputResourcePath);
            	jpWord.add(String.valueOf(i+1));
             } 

        }
	}
	
	static int mRowNumber = 0;
	static public void saveData(String[] data) {
	        row = sheet.createRow(mRowNumber++);
	        for(int i = 0; i < data.length; i++) {
		        cell = row.createCell(i);
		        cell.setCellValue(data[i]);
		        if(i == 4) {
		        	CellStyle style = wb.createCellStyle();
		        	 Font font = wb.createFont();
		        	 font.setColor(HSSFColor.RED.index);
		        	 style.setFont(font);
		        	 cell.setCellStyle(style);
		        }
	        }
	}
	
}

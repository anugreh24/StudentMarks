package utils;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	static ArrayList<Student> al = new ArrayList<Student>();



	public static void percentile() {
		int total_students = al.size();
		for (int i=0;i<5;i++) {
			int count = 0;
			double a = al.get(i).percentage;
			for (int j=0;j<5;j++) {
				if (al.get(j).percentage <= a) {
					count++;	
				}
			}
			double percentile = count*100/total_students;
			System.out.println("Percentile for "+al.get(i).name+" is "+percentile);
		}
	}

	public static void math_max() {
		double max = 0;
		for (int i=0;i<5;i++) {
			if(al.get(i).maths_score>max) {
				max = al.get(i).maths_score;
			}
		}

		for (int i=0;i<5;i++) {
			if(al.get(i).maths_score==max) {
				System.out.print(al.get(i).name+" has highest score in Maths with "+al.get(i).maths_score+" marks.");
			}
		}
		System.out.println();
	}

	public static void science_max() {
		double max = 0;
		for (int i=0;i<5;i++) {
			if(al.get(i).science_score>max) {
				max = al.get(i).science_score;
			}
		}

		for (int i=0;i<5;i++) {
			if(al.get(i).science_score==max) {
				System.out.print(al.get(i).name+" has highest score in Science with "+al.get(i).science_score+" marks.");
			}
		}
		System.out.println();
	}

	public static void language_max() {
		double max = 0;
		for (int i=0;i<5;i++) {
			if(al.get(i).language_score>max) {
				max = al.get(i).language_score;
			}
		}

		for (int i=0;i<5;i++) {
			if(al.get(i).language_score==max) {
				System.out.print(al.get(i).name+" has highest score in Language with "+al.get(i).language_score+" marks.");
			}
		}
		System.out.println();
		System.out.println();

	}

	public static void storeData() throws IOException {
		String excelPath = "./data/Scores.xlsx";
		String sheetName = "Sheet1";
		workbook = new XSSFWorkbook(excelPath);
		sheet = workbook.getSheet(sheetName);
		Iterator<Row> iterator = sheet.iterator();
		while(iterator.hasNext()) {
			Row row = iterator.next();
			Iterator<Cell> cellIterator = row.iterator();
			Cell cell = cellIterator.next();
			int rowNum = cell.getRowIndex();
			int colNum = cell.getColumnIndex();
			if ( rowNum!=0  ) {
				String a = returnName(rowNum,colNum);
				double b = returnMarks(rowNum,colNum+1);
				double c = returnMarks(rowNum,colNum+2);
				double d = returnMarks(rowNum,colNum+3);	
				double e = (b+c+d)/3;
				Student s1=new Student(a,b,c,d,e); 
				al.add(s1);
			}
		}
		workbook.close();
		DecimalFormat numberFormat = new DecimalFormat("#.00");
		Iterator itr=al.iterator();  
		while(itr.hasNext()){  
			Student st=(Student)itr.next();  
			System.out.println(st.name+" has percentage "+numberFormat.format(st.percentage) );  
		} 
		System.out.println();
	}

	public static void printExcel() {
		try {
			String excelPath = "./data/Scores.xlsx";
			String sheetName = "Sheet1";
			workbook = new XSSFWorkbook(excelPath);
			sheet = workbook.getSheet(sheetName);
			Iterator<Row> iterator = sheet.iterator();
			while(iterator.hasNext()) {
				Row row = iterator.next();
				Iterator<Cell> cellIterator = row.iterator();
				while(cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					printCellData(cell.getRowIndex(), cell.getColumnIndex());		
				}
				System.out.println();
			}
			workbook.close();
			System.out.println();
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}

	public static void printCellData(int rowNum, int colNum) throws IOException {
		if ( rowNum==0 | colNum==0 ) {
			String value = sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
			System.out.print(value+"\t");
		}
		else {
			try {
				double value = sheet.getRow(rowNum).getCell(colNum).getNumericCellValue();
				if (value < 0) {
					throw new IllegalArgumentException("Negative numbers are not allowed."); 
				}      
				System.out.print(value+"\t");
			} catch (IllegalArgumentException e) {
				System.out.println(e.getMessage());
			}
		}
	}

	public static String returnName(int rowNum, int colNum) {
		String value = sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
		return value;
	}

	public static double returnMarks(int rowNum, int colNum) {
		double value = sheet.getRow(rowNum).getCell(colNum).getNumericCellValue();
		return value;
	}
}

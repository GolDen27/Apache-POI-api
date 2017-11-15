package main.java.com.epam.testExcel.main;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteTest {

	public static void test() {

		Workbook workbook = new XSSFWorkbook(); //Создание экселевской книги
		
		Sheet sheet0 = workbook.createSheet(); //Создание листа
		Sheet workSheet = workbook.createSheet("MySheet"); //Создание именованного листа
		Sheet sheet2 = workbook.createSheet(WorkbookUtil.createSafeSheetName("??>!@31!@$P{()^Yer?>#@645t")); //Создание листа с заменой запрещённых символов пробелами
		
		Row row = workSheet.createRow(1); //создание рабочей строки. 1=0, 2=1, 3=2, 4=3
		Cell cell = row.createCell(5); //создание рабочей ячейки. A=0, B=1, C=2, D=3 
		
		cell.setCellValue("Hi!!!"); //запись значения в ячейку
		
		workSheet.createRow(3).createCell(3).setCellValue("Так короче"); //упрощённая запись
		
		try {
			FileOutputStream output = new FileOutputStream("Test1.xlsx"); 
			workbook.write(output); //книга -> поток
			workbook.close();
			output.close();
		}catch (IOException e) {
			e.printStackTrace();
		}
		

	}

}

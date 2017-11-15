package main.java.com.epam.testExcel.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

public class ReadTest {
	
	public static void test() {
		
		try {
			FileInputStream input = new FileInputStream(new File("Test1.xlsx")); 
			

			Workbook workbook = new XSSFWorkbook(input); //Создание экселевской книги
			Sheet workSheet = workbook.getSheet("MySheet"); //Создание листа
			
			Iterator<Row> iterator = workSheet.iterator();

            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                while (cellIterator.hasNext()) {

                    Cell currentCell = cellIterator.next();
                    
                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        System.out.println("Data [" + currentRow.getRowNum() + ";" + currentCell.getColumnIndex() + "] -" + currentCell.getStringCellValue());
                    }

                }
            } // это стандартный и простейший способ обхода листа в ексельке
            
            
            //если мы знаем где, что и как заполнено, то можно так))
            Row workRow = workSheet.getRow(3);
            Cell workCell = workRow.getCell(3);
            String workData = workCell.getRichStringCellValue().toString();
            System.out.println(workData);
            
			
			input.close();
		}catch (IOException e) {
			e.printStackTrace();
		}
		

	}
	
}

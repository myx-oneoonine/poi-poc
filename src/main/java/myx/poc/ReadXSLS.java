package myx.poc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadXSLS {
	public static void main(String[] args) {
		String fileName = "/Users/myx4play/Downloads/Template_Justify_New_Community_2018_Test_v3_Pico.xlsx";

		try {
			FileInputStream excelFile = new FileInputStream(new File(fileName));
			Workbook workbook = new XSSFWorkbook(excelFile);

			System.out.println(workbook.getNumberOfSheets());

			Sheet datatypeSheet = workbook.getSheetAt(0);

			List<String> sheetName = new ArrayList<>();
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				sheetName.add(workbook.getSheetName(i));
				System.out.println("Sheet name: " + workbook.getSheetName(i));
			}

			System.out.println("xxx :" + sheetName.contains("Details"));

			Iterator<Row> rows = datatypeSheet.iterator();

			while (rows.hasNext()) {
				Row currentRow = rows.next();
				Iterator<Cell> cells = currentRow.iterator();

				while (cells.hasNext()) {
					Cell currentCell = cells.next();

					if (currentCell.getCellTypeEnum() == CellType.STRING) {
						System.out.print(currentCell.getStringCellValue() + "--");
					} else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
						System.out.print(currentCell.getNumericCellValue() + "--");
					}

				}

			}
		} catch (FileNotFoundException e) {

			e.printStackTrace();
		} catch (IOException e) {

			e.printStackTrace();
		}

	}
}

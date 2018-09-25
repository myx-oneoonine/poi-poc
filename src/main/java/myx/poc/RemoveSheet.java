package myx.poc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class RemoveSheet {

	public static void main(String[] args) {

		String fileName = "/Users/myx4play/Downloads/Template_Justify_New_Community_2018_Test_v3_Pico.xlsx";

		try {
			FileInputStream excelFile = new FileInputStream(new File(fileName));
			Workbook workbook = new XSSFWorkbook(excelFile);

			System.out.println(workbook.getNumberOfSheets());

			List<String> sheetName = new ArrayList<>();
			HashMap<String, Integer> sheetNameMap = new HashMap<>();
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				sheetName.add(workbook.getSheetName(i));
				sheetNameMap.put(workbook.getSheetName(i), i);
				System.out.println("Sheet name: " + workbook.getSheetName(i));
			}
			System.out.println("sheet name :" + sheetName.contains("Details"));

			workbook.removeSheetAt(sheetNameMap.get("Details").intValue());

			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				System.out.println("Sheet name: " + workbook.getSheetName(i));
			}

			OutputStream file = new FileOutputStream("/Users/myx4play/Downloads/Template_Justify_New_Community_2018_Test_v3_Pico.xlsx");
			workbook.write(file);

		} catch (FileNotFoundException e) {

			e.printStackTrace();
		} catch (IOException e) {

			e.printStackTrace();
		}

	}

}

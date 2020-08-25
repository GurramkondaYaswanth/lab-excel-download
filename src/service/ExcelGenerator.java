package service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import com.sun.rowset.internal.Row;

import model.Prograd;

//			Progression -2
//Go to src/service. Open the ExcelGenerator and fill the logic inside the excelGenerate method.
//Your excelGenerate method should return the HSSFWorkbook object as output
//Stick to the instructions clearly. If you face any issue contact your mentor to get the guidance. 

public class ExcelGenerator {
	
	FileOutputStream out;
	public HSSFWorkbook excelGenerate(Prograd prograd, List<Prograd> list) throws IOException {
		try {


			
			// Type your code here

			//object of HSSFWorkbook class
			 HSSFWorkbook hwb=new HSSFWorkbook();
			 //creation of sheet
			 HSSFSheet sheet = hwb.createSheet("ProGrad Details");
			 //creating a row head
			 HSSFRow hRow=sheet.createRow(0);
			  hRow.createCell(0).setCellValue("Prograd ID");
			  hRow.createCell(1).setCellValue("Name");
			  hRow.createCell(2).setCellValue("Rating");
			  hRow.createCell(3).setCellValue("Comment");
			  hRow.createCell(4).setCellValue("Recommendation");
			   
			  //reading inputs from list and adding to each row 
			  int i=0;
			  for(Prograd prograd1 :list)
			  {
				  int j=i+1;
				  HSSFRow hrow = sheet.createRow(j);
				  hrow.createCell(0).setCellValue(prograd1.getId());
					hrow.createCell(1).setCellValue(prograd1.getName());
					hrow.createCell(2).setCellValue(prograd1.getRate());
					hrow.createCell(3).setCellValue(prograd1.getComment());
					hrow.createCell(4).setCellValue(prograd1.getRecommend());
					i++;
			  }
		         
			
		
			
			// Do not modify the lines given below
			out = new FileOutputStream(new File("D:/excelgen.xlsx"));
			hwb.write(out);
		
			return hwb;
			}
		catch (Exception e) {
				e.printStackTrace();
			}
		finally {
			out.close();
			System.out.println("sucessfully created");
		}
		return null;
		
	}
}

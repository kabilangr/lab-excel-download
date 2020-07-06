package service;

import java.io.FileOutputStream;
import java.io.IOException;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import model.Prograd;

//			Progression -1 
//Go to src/service. Open the ExcelGenerator and fill the logic inside the excelGenerate method.
//
//Stick to the instructions clearly. If you face any issue contact your mentor to get the guidance. 

public class ExcelGenerator {
	
	FileOutputStream out;
	public HSSFWorkbook excelGenerate(Prograd prograd, List<Prograd> list) throws IOException {
		try {

              HSSFWorkbook ob=new HSSFWorkbook();
			HSSFSheet sheet=ob.createSheet("ProGrad Details");
			HSSFRow h=sheet.createRow(0);
			h.createCell(0).setCellValue("ProGrad ID");
			h.createCell(1).setCellValue("Name");
			h.createCell(2).setCellValue("Rating");
			h.createCell(3).setCellValue("Comments");
			h.createCell(4).setCellValue("Recommendation");
		int i=0;
		for(Prograd p : list)
		{
			int j=i+1;
			HSSFRow hr=sheet.createRow(j);
			hr.createCell(0).setCellValue(p.getId());
			hr.createCell(1).setCellValue(p.getName());
			hr.createCell(2).setCellValue(p.getRate());
			hr.createCell(3).setCellValue(p.getComment());
			hr.createCell(4).setCellValue(p.getRecommend());
			i++;
			}
			// Do not modify the lines given below
			 out = new FileOutputStream("C:\\Users\\Dagger\\labs\\javapro\\lab-excel-download\\chumma.xlsx");
			ob.write(out);
		
			return ob;
			}
		catch (Exception e) {
				e.printStackTrace();
			}
		finally {
			out.close();
		}
		return null;
		
	}
}

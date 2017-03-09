package com.mpastor.app.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import com.mpastor.app.bean.FPlatform;
import com.mpastor.app.bean.KPlatform;


public class ExcelUtil {
	
	
	public ExcelUtil(){
	 	
	}
	
	public void generateExcelFileKPlatform() {
		
		try {
		
			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFSheet sheet = workbook.createSheet("Hoja1");
			
			CellStyle style = workbook.createCellStyle();
			Font font = workbook.createFont();
		    font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		    style.setFont(font);
		    
			Row row = sheet.createRow(0);
			
			Cell cell0 = row.createCell(0);
			Cell cell1 = row.createCell(1);
			Cell cell2 = row.createCell(2);
			Cell cell3 = row.createCell(3);
			Cell cell4 = row.createCell(4);
			Cell cell5 = row.createCell(5);
			Cell cell6 = row.createCell(6);
			Cell cell7 = row.createCell(7);
			Cell cell8 = row.createCell(8);
			Cell cell9 = row.createCell(9);
			Cell cell10 = row.createCell(10);
			Cell cell11 = row.createCell(11);
			Cell cell12 = row.createCell(12);
			Cell cell13 = row.createCell(13);
			Cell cell14 = row.createCell(14);
			Cell cell15 = row.createCell(15);
			Cell cell16 = row.createCell(16);
			Cell cell17 = row.createCell(17);
			Cell cell18 = row.createCell(18);
			Cell cell19 = row.createCell(19);
			Cell cell20 = row.createCell(20);
			Cell cell21 = row.createCell(21);
			Cell cell22 = row.createCell(22);
			Cell cell23 = row.createCell(23);
			Cell cell24 = row.createCell(24);
			Cell cell25 = row.createCell(25);
			Cell cell26 = row.createCell(26);
			Cell cell27 = row.createCell(27);
			
			cell0.setCellStyle(style);
			cell1.setCellStyle(style);
			cell2.setCellStyle(style);
			cell3.setCellStyle(style);
			cell4.setCellStyle(style);
			cell5.setCellStyle(style);
			cell6.setCellStyle(style);
			cell7.setCellStyle(style);
			cell8.setCellStyle(style);
			cell9.setCellStyle(style);
			cell10.setCellStyle(style);
			cell11.setCellStyle(style);
			cell12.setCellStyle(style);
			cell13.setCellStyle(style);
			cell14.setCellStyle(style);
			cell15.setCellStyle(style);
			cell16.setCellStyle(style);
			cell17.setCellStyle(style);
			cell18.setCellStyle(style);
			cell19.setCellStyle(style);
			cell20.setCellStyle(style);
			cell21.setCellStyle(style);
			cell22.setCellStyle(style);
			cell23.setCellStyle(style);
			cell24.setCellStyle(style);
			cell25.setCellStyle(style);
			cell26.setCellStyle(style);
			cell27.setCellStyle(style);
			
			cell0.setCellValue("Nombre");
			cell1.setCellValue("Apellido 1");
			cell2.setCellValue("Apellido 2");
			cell3.setCellValue("Edad");
			cell4.setCellValue("Peso");
			cell5.setCellValue("Talla");
			cell6.setCellValue("Número pie");
			cell7.setCellValue("Hora medición");
			cell8.setCellValue("Tipo Plataforma");
			cell9.setCellValue("Ojos");
			cell10.setCellValue("Intento");
			cell11.setCellValue("X max (mm)");
			cell12.setCellValue("X min (mm)");
			cell13.setCellValue("Y max (mm)");
			cell14.setCellValue("Y min (mm)");
			cell15.setCellValue("X SD (mm)");
			cell16.setCellValue("Y SD (mm)");
			cell17.setCellValue("X Avg (mm)");
			cell18.setCellValue("Dev CG x (mm)");
			cell19.setCellValue("Y Avg (mm)");
			cell20.setCellValue("Dev CG Y (mm)");
			cell21.setCellValue("Area circular (mm2)");
			cell22.setCellValue("Area rectangular (mm2)");
			cell23.setCellValue("Area efectiva (mm2)");
			cell24.setCellValue("Area Elipse 95th Percentil (mm)");
			cell25.setCellValue("Velocidad promedio (mm/seg)");
			cell26.setCellValue("X D Avg (mm)");
			cell27.setCellValue("Y D Avg (mm)");
			
			FileOutputStream out = new FileOutputStream(new File("/Users/mpastor/Downloads/mediciones_plataformaK.xls"));
		    workbook.write(out);
		    out.close();
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void writeKPlatform(KPlatform kplatform, File kPlatformExcelFile, int pos, String color) {
		
		try {
		
			FileInputStream fileInputStream = new FileInputStream(kPlatformExcelFile);
			
			HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
			HSSFSheet sheet = workbook.getSheetAt(0);
			
			CellStyle style = workbook.createCellStyle();
			style.setFillPattern(CellStyle.SOLID_FOREGROUND);
			
			if (color.equals("blue")) {
				style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.index);
			} else if (color.equals("lemon")) {
				style.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.index);	
			}
			
			Row row = sheet.createRow(pos);
			
			Cell cell = row.createCell(0);
			cell.setCellValue(kplatform.getName());
			cell.setCellStyle(style);
			
			cell = row.createCell(1);
			cell.setCellValue(kplatform.getLastName());
			cell.setCellStyle(style);
			
			cell = row.createCell(2);
			cell.setCellValue(kplatform.getSecondLastName());
			cell.setCellStyle(style);
			
			cell = row.createCell(3);
			cell.setCellValue("");
			cell.setCellStyle(style);
			
			cell = row.createCell(4);
			cell.setCellValue("");
			cell.setCellStyle(style);
			
			cell = row.createCell(5);
			cell.setCellValue("");
			cell.setCellStyle(style);
			
			cell = row.createCell(6);
			cell.setCellValue("");
			cell.setCellStyle(style);
			
			cell = row.createCell(7);
			cell.setCellValue("");
			cell.setCellStyle(style);
			
			cell = row.createCell(8);
			cell.setCellValue("PLATAFORMA K");
			cell.setCellStyle(style);
			
			cell = row.createCell(9);
			cell.setCellValue(kplatform.getEyesMethod());
			cell.setCellStyle(style);
			
			cell = row.createCell(10);
			cell.setCellValue(kplatform.getNumeroIntento());
			cell.setCellStyle(style);
			
			cell = row.createCell(11);
			cell.setCellValue(kplatform.getxMax());
			cell.setCellStyle(style);
			
			cell = row.createCell(12);
			cell.setCellValue(kplatform.getxMin());
			cell.setCellStyle(style);
			
			cell = row.createCell(13);
			cell.setCellValue(kplatform.getyMax());
			cell.setCellStyle(style);
			
			cell = row.createCell(14);
			cell.setCellValue(kplatform.getyMin());
			cell.setCellStyle(style);
			
			cell = row.createCell(15);
			cell.setCellValue(kplatform.getxSD());
			cell.setCellStyle(style);
			
			cell = row.createCell(16);
			cell.setCellValue(kplatform.getySD());
			cell.setCellStyle(style);
			
			cell = row.createCell(17);
			cell.setCellValue(kplatform.getxAvg());
			cell.setCellStyle(style);
			
			cell = row.createCell(18);
			cell.setCellValue(kplatform.getDevCGx());
			cell.setCellStyle(style);
			
			cell = row.createCell(19);
			cell.setCellValue(kplatform.getyAvg());
			cell.setCellStyle(style);
			
			cell = row.createCell(20);
			cell.setCellValue(kplatform.getDevCGy());
			cell.setCellStyle(style);
			
			cell = row.createCell(21);
			cell.setCellValue(kplatform.getAreaCircular());
			cell.setCellStyle(style);
			
			cell = row.createCell(22);
			cell.setCellValue(kplatform.getAreaRectangular());
			cell.setCellStyle(style);
			
			cell = row.createCell(23);
			cell.setCellValue(kplatform.getAreaEfectiva());
			cell.setCellStyle(style);
			
			cell = row.createCell(24);
			cell.setCellValue(kplatform.getAreaElipse());
			cell.setCellStyle(style);
			
			cell = row.createCell(25);
			cell.setCellValue(kplatform.getvAvg());
			cell.setCellStyle(style);
			
			cell = row.createCell(26);
			cell.setCellValue(kplatform.getXdAvg());
			cell.setCellStyle(style);
			
			cell = row.createCell(27);
			cell.setCellValue(kplatform.getYdAvg());
			cell.setCellStyle(style);
			
			for (int j = 0; j < 28; j++) { 
				sheet.autoSizeColumn((short)j); 
			}
			
			FileOutputStream fileOutputStream = new FileOutputStream(kPlatformExcelFile);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
		
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void writeFPlatform(FPlatform fplatform, File fPlatformExcelFile, int pos, String color) {
		
		try {
			
			FileInputStream fileInputStream = new FileInputStream(fPlatformExcelFile);
			
			HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
			HSSFSheet sheet = workbook.getSheetAt(0);
			
			CellStyle style = workbook.createCellStyle();
			style.setFillPattern(CellStyle.SOLID_FOREGROUND);
			
			if (color.equals("blue")) {
				style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.index);
			} else if (color.equals("lemon")) {
				style.setFillForegroundColor(IndexedColors.LEMON_CHIFFON.index);	
			}
			
			Row row = sheet.createRow(pos);
			
			Cell cell = row.createCell(0);
			cell.setCellValue(fplatform.getName());
			cell.setCellStyle(style);
			
			cell = row.createCell(1);
			cell.setCellValue(fplatform.getLastName());
			cell.setCellStyle(style);
			
			cell = row.createCell(2);
			cell.setCellValue(fplatform.getSecondLastName());
			cell.setCellStyle(style);
			
			cell = row.createCell(3);
			cell.setCellValue("");
			cell.setCellStyle(style);
			
			cell = row.createCell(4);
			cell.setCellValue("");
			cell.setCellStyle(style);
			
			cell = row.createCell(5);
			cell.setCellValue("");
			cell.setCellStyle(style);
			
			cell = row.createCell(6);
			cell.setCellValue("");
			cell.setCellStyle(style);
			
			cell = row.createCell(7);
			cell.setCellValue("");
			cell.setCellStyle(style);
			
			cell = row.createCell(8);
			cell.setCellValue("PLATAFORMA F");
			cell.setCellStyle(style);
			
			cell = row.createCell(9);
			cell.setCellValue(fplatform.getEyesMethod());
			cell.setCellStyle(style);
			
			cell = row.createCell(10);
			cell.setCellValue(fplatform.getNumeroIntento());
			cell.setCellStyle(style);
			
			cell = row.createCell(11);
			cell.setCellValue(fplatform.getOscilacionLatDer());
			cell.setCellStyle(style);
			
			cell = row.createCell(12);
			cell.setCellValue(fplatform.getOscilacionLatIzq());
			cell.setCellStyle(style);
			
			cell = row.createCell(13);
			cell.setCellValue(fplatform.getOscilacionAnterior());
			cell.setCellStyle(style);
			
			cell = row.createCell(14);
			cell.setCellValue(fplatform.getOscilacionPosterior());
			cell.setCellStyle(style);
			
			cell = row.createCell(15);
			cell.setCellValue(fplatform.getDesviacionEstPosEjeX());
			cell.setCellStyle(style);
			
			cell = row.createCell(16);
			cell.setCellValue(fplatform.getDesviacionEstPosEjeY());
			cell.setCellStyle(style);
			
			cell = row.createCell(17);
			cell.setCellValue(fplatform.getBaricentroMedioResX());
			cell.setCellStyle(style);
			
			cell = row.createCell(18);
			cell.setCellValue(fplatform.getBaricentroMedioResY());
			cell.setCellStyle(style);
			
			cell = row.createCell(19);
			cell.setCellValue(fplatform.getSuperficieElipse());
			cell.setCellStyle(style);
			
			cell = row.createCell(20);
			cell.setCellValue(fplatform.getVelocidadTotal());
			cell.setCellStyle(style);
			
			cell = row.createCell(21);
			cell.setCellValue(fplatform.getSwayArea());
			cell.setCellStyle(style);
			
			for (int j = 0; j < 22; j++) { 
				sheet.autoSizeColumn((short)j); 
			}
			
			FileOutputStream fileOutputStream = new FileOutputStream(fPlatformExcelFile);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void generateExcelFileFPlatform() {
		
		try {
		
			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFSheet sheet = workbook.createSheet("Hoja1");
			
			CellStyle style = workbook.createCellStyle();
			Font font = workbook.createFont();
		    font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		    style.setFont(font);
			
			Row row = sheet.createRow(0);
			
			Cell cell0 = row.createCell(0);
			Cell cell1 = row.createCell(1);
			Cell cell2 = row.createCell(2);
			Cell cell3 = row.createCell(3);
			Cell cell4 = row.createCell(4);
			Cell cell5 = row.createCell(5);
			Cell cell6 = row.createCell(6);
			Cell cell7 = row.createCell(7);
			Cell cell8 = row.createCell(8);
			Cell cell9 = row.createCell(9);
			Cell cell10 = row.createCell(10);
			Cell cell11 = row.createCell(11);
			Cell cell12 = row.createCell(12);
			Cell cell13 = row.createCell(13);
			Cell cell14 = row.createCell(14);
			Cell cell15 = row.createCell(15);
			Cell cell16 = row.createCell(16);
			Cell cell17 = row.createCell(17);
			Cell cell18 = row.createCell(18);
			Cell cell19 = row.createCell(19);
			Cell cell20 = row.createCell(20);
			Cell cell21 = row.createCell(21);
			
			cell0.setCellStyle(style);
			cell1.setCellStyle(style);
			cell2.setCellStyle(style);
			cell3.setCellStyle(style);
			cell4.setCellStyle(style);
			cell5.setCellStyle(style);
			cell6.setCellStyle(style);
			cell7.setCellStyle(style);
			cell8.setCellStyle(style);
			cell9.setCellStyle(style);
			cell10.setCellStyle(style);
			cell11.setCellStyle(style);
			cell12.setCellStyle(style);
			cell13.setCellStyle(style);
			cell14.setCellStyle(style);
			cell15.setCellStyle(style);
			cell16.setCellStyle(style);
			cell17.setCellStyle(style);
			cell18.setCellStyle(style);
			cell19.setCellStyle(style);
			cell20.setCellStyle(style);
			cell21.setCellStyle(style);
			
			cell0.setCellValue("Nombre");
			cell1.setCellValue("Apellido 1");
			cell2.setCellValue("Apellido 2");
			cell3.setCellValue("Edad");
			cell4.setCellValue("Peso");
			cell5.setCellValue("Talla");
			cell6.setCellValue("Número pie");
			cell7.setCellValue("Hora medición");
			cell8.setCellValue("Tipo Plataforma");
			cell9.setCellValue("Ojos");
			cell10.setCellValue("Intento");
			cell11.setCellValue("X max (mm)");
			cell12.setCellValue("X min (mm)");
			cell13.setCellValue("Y max (mm)");
			cell14.setCellValue("Y min (mm)");
			cell15.setCellValue("X SD (mm)");
			cell16.setCellValue("Y SD (mm)");
			cell17.setCellValue("Baricentro medio X (mm)");
			cell18.setCellValue("Baricentro medio Y (mm)");
			cell19.setCellValue("Superficie elipse (mm2)");
			cell20.setCellValue("Velocidad (mm/seg)");
			cell21.setCellValue("Sway Area (mm)");
			
			FileOutputStream out = new FileOutputStream(new File("/Users/mpastor/Downloads/mediciones_plataformaF.xls"));
		    workbook.write(out);
		    out.close();
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
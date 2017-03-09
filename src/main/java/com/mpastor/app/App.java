package com.mpastor.app;

import java.io.File;
import com.mpastor.app.bean.FPlatform;
import com.mpastor.app.bean.KPlatform;
import com.mpastor.app.utils.DataExtractor;
import com.mpastor.app.utils.ExcelUtil;


public class App {
	private static final String PDF_EXT = "pdf";
	
	public static void main( String[] args ) {
		
		String pathPDFsPlataformaK = args[0];
		String pathPDFsPlataformaF = args[1];
		
		try {
		
			System.out.println("***** ***** Comienzo del proceso ***** *****");
			
			ExcelUtil excelUtil = new ExcelUtil();
			// Generamos el excel para la plataforma K :
			excelUtil.generateExcelFileKPlatform();
			// Generamos el excel para la plataforma F :
			excelUtil.generateExcelFileFPlatform();
			
			// Leemos los pdfs de la plataforma K y los vamos insertando en el excel creado :
			File kPdfsfilePath = new File(pathPDFsPlataformaK);
			File[] listOfKpdfs = kPdfsfilePath.listFiles();
			
			DataExtractor dataExtractor = new DataExtractor();
			
			int pos = 1;
			
			String nombrePacienteKIni = "";
			String nombrePacienteFIni = "";
			String color = "blue";
			
			for (File file : listOfKpdfs) {
				if (file.isFile()) {
					if (file.getName().contains(PDF_EXT)){

						// Dar color a las celdas por Paciente :
						String nombrePacienteK = "";
						if (file.getName().contains("CE")) { 
							nombrePacienteK = file.getName().substring(0, file.getName().indexOf("CE"));	
						} else if (file.getName().contains("ce")) {
							nombrePacienteK = file.getName().substring(0, file.getName().indexOf("ce"));	
						} else if (file.getName().contains("OE")) {
							nombrePacienteK = file.getName().substring(0, file.getName().indexOf("OE"));	
						} else if (file.getName().contains("oe")) {
							nombrePacienteK = file.getName().substring(0, file.getName().indexOf("oe"));	
						}
						if (nombrePacienteKIni.equals("")){
							nombrePacienteKIni = nombrePacienteK;	
						}
						
						if (!nombrePacienteKIni.equals(nombrePacienteK)) {
							nombrePacienteKIni = nombrePacienteK;
							if (color.equals("blue"))
								color = "lemon";
							else if (color.equals("lemon"))
								color = "blue";
						} 
						// Fin color a celdas
						
						System.out.println("***** Extrayendo los datos del fichero : "+file.getAbsolutePath());
						KPlatform kPlatformData = dataExtractor.getDataFromKPlatformPDF(file.getAbsolutePath());
						kPlatformData.setNumeroIntento(getNumeroIntento(file.getName()));
						excelUtil.writeKPlatform(kPlatformData, new File("/Users/mpastor/Downloads/mediciones_plataformaK.xls"), pos, color);
						System.out.println("***** Los datos se han insertado correctamente en el fichero excel...");
						pos++;
					}
				}
			}
			
			pos = 1;
			
			// Leemos los pdfs de la plataforma F y los vamos insertando en el excel creado :
			File fPdfsfilePath = new File(pathPDFsPlataformaF);
			File[] listOfFpdfs = fPdfsfilePath.listFiles();
			
			for (File file : listOfFpdfs) {
				if (file.isFile()) {
					if (file.getName().contains(PDF_EXT)){
						
						// Dar color a las celdas por Paciente :
						String nombrePacienteF = "";
						if (file.getName().contains("CE")) { 
							nombrePacienteF = file.getName().substring(0, file.getName().indexOf("CE"));	
						} else if (file.getName().contains("ce")) {
							nombrePacienteF = file.getName().substring(0, file.getName().indexOf("ce"));	
						} else if (file.getName().contains("OE")) {
							nombrePacienteF = file.getName().substring(0, file.getName().indexOf("OE"));	
						} else if (file.getName().contains("oe")) {
							nombrePacienteF = file.getName().substring(0, file.getName().indexOf("oe"));	
						}
						if (nombrePacienteFIni.equals("")){
							nombrePacienteFIni = nombrePacienteF;	
						}
						
						if (!nombrePacienteFIni.equals(nombrePacienteF)) {
							nombrePacienteFIni = nombrePacienteF;
							if (color.equals("blue"))
								color = "lemon";
							else if (color.equals("lemon"))
								color = "blue";
						} 
						// Fin color a celdas
						
						System.out.println("***** Extrayendo los datos del fichero : "+file.getAbsolutePath());
						FPlatform fPlatformData = dataExtractor.getDataFromFPlatformPDF(file.getAbsolutePath());
						fPlatformData.setNumeroIntento(getNumeroIntento(file.getName()));
						excelUtil.writeFPlatform(fPlatformData, new File("/Users/mpastor/Downloads/mediciones_plataformaF.xls"), pos, color);
						System.out.println("***** Los datos se han insertado correctamente en el fichero excel...");
						pos++;
					}
				}
			}
			
			System.out.println("***** ***** Fin del proceso ***** *****");
			
		} catch (Exception e){
			e.printStackTrace();
		}
	}
	
	private static String getNumeroIntento(String fileName) {
		String name = fileName.substring(0, fileName.indexOf(".pdf"));
		return name.substring(name.length()-1, name.length());
	}
	
}
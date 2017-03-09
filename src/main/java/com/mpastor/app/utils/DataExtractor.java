package com.mpastor.app.utils;

import java.io.File;
import java.io.FileInputStream;
import java.text.DecimalFormat;
import org.apache.pdfbox.cos.COSDocument;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.util.PDFTextStripper;
import com.mpastor.app.bean.FPlatform;
import com.mpastor.app.bean.KPlatform;


public class DataExtractor {
	
	public DataExtractor(){
		
	}
	
	public KPlatform getDataFromKPlatformPDF(String filePath) throws Exception {
	
		KPlatform kPlatform = new KPlatform();
		
		String data = getData(filePath);
		
		//System.out.println(data);
		
		String completeName = getKPlatformName(data);
		String[] arrayName = completeName.split(" ");
		
		kPlatform.setName(arrayName[0]);
		
		if (arrayName.length>1)
			kPlatform.setLastName(arrayName[1]);
		else
			kPlatform.setLastName("");
		
		if (arrayName.length>2)
			kPlatform.setSecondLastName(arrayName[2]);
		else
			kPlatform.setSecondLastName("");
		
		kPlatform.setxMax(convertCmIntoMm(getGeneralStatics(data, PDFConstants.K_PLATFORM_XMAX, 3, PDFConstants.CM)));
		kPlatform.setxMin(convertCmIntoMm(getGeneralStatics(data, PDFConstants.K_PLATFORM_XMIN, 3, PDFConstants.CM)));
		kPlatform.setyMax(convertCmIntoMm(getGeneralStatics(data, PDFConstants.K_PLATFORM_YMAX, 3, PDFConstants.CM)));
		kPlatform.setyMin(convertCmIntoMm(getGeneralStatics(data, PDFConstants.K_PLATFORM_YMIN, 3, PDFConstants.CM)));
		kPlatform.setxSD(convertCmIntoMm(getGeneralStatics(data, PDFConstants.K_PLATFORM_XSD, 3, PDFConstants.CM)));
		kPlatform.setySD(convertCmIntoMm(getGeneralStatics(data, PDFConstants.K_PLATFORM_YSD, 3, PDFConstants.CM)));
		kPlatform.setxAvg(convertCmIntoMm(getGeneralStatics(data, PDFConstants.K_PLATFORM_XAVG, 3, PDFConstants.CM)));
		kPlatform.setyAvg(convertCmIntoMm(getGeneralStatics(data, PDFConstants.K_PLATFORM_YAVG, 3, PDFConstants.CM)));
		kPlatform.setXdAvg(convertCmIntoMm(getGeneralStatics(data, PDFConstants.K_PLATFORM_XDISPAVG, 4, PDFConstants.CM)));
		kPlatform.setYdAvg(convertCmIntoMm(getGeneralStatics(data, PDFConstants.K_PLATFORM_YDISPAVG, 4, PDFConstants.CM)));
		kPlatform.setDevCGx(convertCmIntoMm(getGeneralStatics(data, PDFConstants.K_PLATFORM_DEVCGX, 4, PDFConstants.CM)));
		kPlatform.setDevCGy(convertCmIntoMm(getGeneralStatics(data, PDFConstants.K_PLATFORM_DEVCGY, 4, PDFConstants.CM)));
		kPlatform.setAreaCircular(convertCmIntoMm(getGeneralStatics(data, PDFConstants.K_PLATFORM_AREACIRCULAR, 3, PDFConstants.CM)));
		kPlatform.setAreaRectangular(convertCmIntoMm(getGeneralStatics(data, PDFConstants.K_PLATFORM_AREARECTANG, 3, PDFConstants.CM)));
		kPlatform.setAreaEfectiva(convertCmIntoMm(getGeneralStatics(data, PDFConstants.K_PLATFORM_AREAEFFECT, 3, PDFConstants.CM)));
		kPlatform.setAreaElipse(convertCmIntoMm(getGeneralStatics(data, PDFConstants.K_PLATFORM_AREAELLIPSE, 2, PDFConstants.CM)));
		kPlatform.setvAvg(convertCmIntoMm(getGeneralStatics(data, PDFConstants.K_PLATFORM_VAVG, 3, PDFConstants.CM)));
		kPlatform.setEyesMethod(getKPlatformEyesMethod(data));
		
		return kPlatform;
	}
	
	public FPlatform getDataFromFPlatformPDF(String filePath) throws Exception {
		
		FPlatform fPlatform = new FPlatform();
		
		String data = getData(filePath);
		
		data = data.replaceAll("\n", "");
		
		//System.out.println(data);
		
		String completeName = getFPlatformName(data);
		String[] arrayName = completeName.split(" ");
		fPlatform.setName(arrayName[0]);
		fPlatform.setLastName(arrayName[1]);
		
		if (arrayName.length == 3)
			fPlatform.setSecondLastName(arrayName[2]);
		else
			fPlatform.setSecondLastName("");
		
		fPlatform.setOscilacionLatDer(getGeneralStatics(data, PDFConstants.F_PLATFORM_OSCLATDER, 3, PDFConstants.MM));
		fPlatform.setOscilacionLatIzq(getGeneralStatics(data, PDFConstants.F_PLATFORM_OSCLATIZQ, 3, PDFConstants.MM));
		fPlatform.setOscilacionAnterior(getGeneralStatics(data, PDFConstants.F_PLATFORM_OSCANTERIOR, 2, PDFConstants.MM));
		fPlatform.setOscilacionPosterior(getGeneralStatics(data, PDFConstants.F_PLATFORM_OSCPOSTERIOR, 2, PDFConstants.MM));
		fPlatform.setDesviacionEstPosEjeX(getGeneralStaticsSpecial(data, PDFConstants.F_PLATFORM_DESESTEJEX, 6, "D"));
		fPlatform.setDesviacionEstPosEjeY(getGeneralStaticsSpecial(data, PDFConstants.F_PLATFORM_DESESTEJEY, 5, "D"));
		fPlatform.setBaricentroMedioResX(getGeneralStatics(data, PDFConstants.F_PLATFORM_BARIMEDRES, 4, PDFConstants.MM));
		fPlatform.setBaricentroMedioResY(getBaricentroMedioResY(data));
		fPlatform.setSwayArea(getGeneralStatics(data, PDFConstants.F_PLATFORM_SWAYAREA, 2, PDFConstants.MM));
		fPlatform.setSuperficieElipse(getGeneralStatics(data, PDFConstants.F_PLATFORM_SUPELIPSE, 5, PDFConstants.MM));
		
		// Para los ficheros de David : fPlatform.setEyesMethod(getFPlatformEyesMethod(data));
		fPlatform.setEyesMethod(getEyesMethodFromFileName(filePath));
		
		fPlatform.setVelocidadTotal(getGeneralStatics(data, PDFConstants.F_PLATFORM_VELCOMPLE, 5, PDFConstants.MM_SEC));
		
		return fPlatform;
	}
	
	private String getKPlatformName(String data) {
		String result = "";
		String[] tmp;
		result = data.substring(data.indexOf(PDFConstants.K_PLATFORM_FIRSTNAME), data.indexOf(PDFConstants.K_PLATFORM_DATE));
		result = result.replaceAll("\n", "");
		result = result.replace("Middle Name", ":");
		tmp = result.split(":");
		System.out.println("***** tmp length : "+tmp.length);
		if(tmp.length>4)
			result = tmp[1].trim() + " " + tmp[4].trim();
		else
			result = tmp[1].trim();
			
		return result;
	}
	
	private String getFPlatformName(String data) {
		
		String result = "";
		String[] tmp;
		result = data.substring(data.indexOf(PDFConstants.F_PLATFORM_PACIENTE), data.indexOf(PDFConstants.F_PLATFORM_FECHANAC));
		tmp = result.split(":");
		result = tmp[1];
		tmp = result.split(" ");
		
		if (tmp.length == 4) {
			result = tmp[3] + " " + tmp[1] + " " + tmp[2];	
		} else if (tmp.length == 3) {
			result = tmp[2] + " " + tmp[1];	
		} else if (tmp.length == 2) {
			result = tmp[3] + " " + tmp[1];	
		}
		
		return result;
	}
	
	private String getKPlatformEyesMethod(String data) {
		if (data.contains("Eyes Closed")){
			return "CE";
		} else {
			return "OE";
		} 
	}
	
	private String getFPlatformEyesMethod(String data) {
		String result = "";
		result = data.substring(data.indexOf(PDFConstants.F_PLATFORM_NOTAANALISIS), data.length());
		result = result.substring(0, result.indexOf(PDFConstants.F_PLATFORM_LORANENG));
		if (result.contains("ce")){
			return "CE";
		} else {
			return "OE";
		}
	}
	
	private String getEyesMethodFromFileName(String fileName) {
		if (fileName.contains(" OE "))
			return "OE";
		else
			return "CE";
	}
	
	private String getBaricentroMedioResY(String data) {
		String result = "";
		String[] tmp;
		result = data.substring(data.indexOf(PDFConstants.F_PLATFORM_BARIMEDRES), data.length());
		result = result.substring(0,result.indexOf("D"));
		tmp = result.split(PDFConstants.MM);
		result = tmp[1];
		tmp = result.trim().split(":");
		result = tmp[1];
		//System.out.println("***** result : "+result.trim());
		return result.trim();
	}
	
	private String getGeneralStatics(String data, String key, int pos, String delimitter) {
		String result = "";
		String[] tmp;
		result = data.substring(data.indexOf(key), data.length());
		result = result.substring(0, result.indexOf(delimitter));
		tmp = result.split(" ");
		result = tmp[pos];
		//System.out.println("***** result : "+result);
		return result;	
	}
	
	private String getGeneralStaticsSpecial(String data, String key, int pos, String delimitter) {
		String result = "";
		String[] tmp;
		result = data.substring(data.indexOf(key), data.length());
		result = result.substring(1, result.length());
		result = result.substring(1, result.indexOf(delimitter));
		tmp = result.split(" ");
		result = tmp[pos];
		//System.out.println("***** result : "+result);
		return result;	
	}
	
	private String getData(String filePath) {
		
		String parsedText = "";
		
		try {

			PDDocument pdDoc = null;
			PDFTextStripper pdfStripper = null;
			COSDocument cosDoc = null;
			File file = new File(filePath);
	
			PDFParser parser = new PDFParser(new FileInputStream(file));
			parser.parse();
			cosDoc = parser.getDocument();
			pdfStripper = new PDFTextStripper();
			pdDoc = new PDDocument(cosDoc);
			pdfStripper.setStartPage(1);
			pdfStripper.setEndPage(5);
			parsedText = pdfStripper.getText(pdDoc);
			
			pdDoc.close();
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		return parsedText;
	}

	private String convertCmIntoMm(String cmValue) {
		String result = "";
		DecimalFormat decFormat = new DecimalFormat("0.00");
		double cmDoubleValue = Double.parseDouble(cmValue);
		double mmValue = cmDoubleValue * 10;
		result = String.valueOf(decFormat.format(mmValue));
		if (result.substring(result.length()-2, result.length()).equals("00")) {
			result = result.substring(0, result.length()-3);
		}
		return result;
	}
	
}
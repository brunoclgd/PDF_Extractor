
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.pdfbox.pdmodel.font.PDCIDFontType2;
import org.apache.pdfbox.pdmodel.font.PDSimpleFont;

public class Geral_PDF_GIA {

	public static int line = 1;
	public static int indice = 0;
	
	public static String getData(String data) {
		String formatedData;
		int i;
		i = data.indexOf(" ");
		formatedData = data.substring(0,i)+"/"+data.substring(i+1, data.length());
		return formatedData;
	}
	
	public static int checkTypeOf(String[] text) {
		int type = 0, i = 0;


		while(text[i].equals(" ") || text[i].contains("ADE Conjunto Cotec/Corat")) {
			i++;
		}
		if(text[i].contains("CFOPs Entradas")) {
			type = 1;
		}
		else if(text[i].contains("Guia de Informação de Apuração do ICMS")) {
			type = 2;
		}
		else if(text[i].contains("NO ESTADO")) {
			type = 3;
		}
		else if(text[i].contains("OPERAÇÕES PRÓPRIAS")) {
			type = 4;
		}
		
		return type;
	}
	public static String getReferencia(String lines[]) {
		String ref = null;
		String aux;
		
		aux = lines[1];
		
		ref = aux.substring(0, 7);
		
		return ref;
		
		
	}
	public static String getCNPJ(String lines[]) {
		String ref = null;
		String aux;
		aux = lines[1];
		ref = aux.substring(10, aux.length());
		
		return ref;
	}
	public static void convertPDFToExcel(File inputFile, Sheet sheet) {
		PDDocument document;
		try {
			document = PDDocument.load(inputFile);

			document.getClass();
			if (!document.isEncrypted()) {
				Row rowData = sheet.createRow(line);
				PDFTextStripperByArea stripper = new PDFTextStripperByArea();
				stripper.setSortByPosition(true);
				
				PDFTextStripper tStripper = new PDFTextStripper();
				String pdfFileInText = tStripper.getText(document);
				//String pdfFileInText = tStripper.toString();
				
				String lines[] = pdfFileInText.split("\\r?\\n");
				String auxLines[] = null;
				String IRST1[] = null;
				String IRST2[] = null;
				String IRST3[] = null;
				String IRST4[] = null;
				String aux = null;
				int type = checkTypeOf(lines);
				if(type == 1) {
					for(String line : lines) {
						if ( line.contains("1.403")){
							IRST1 = line.split(" ");
							
						}
						if(line.contains("2.403")) {
							IRST2 = line.split(" ");
						}
						if(line.contains("1.406")) {
							IRST3 = line.split(" ");
						}
						if(line.contains("1.407")) {
							IRST4 = line.split(" ");
						}
					}
					System.out.println("a");
					String data = getData(lines[5]);
					System.out.println("NOME DO CONTRIBUINTE: "+lines[7]);
					rowData.createCell(0).setCellValue(lines[7]);
					sheet.autoSizeColumn(0);
					indice++;
					
					auxLines = lines[9].split(" ");
					aux = auxLines[1];
					
					System.out.println("CNPJ: "+aux);
					rowData.createCell(1).setCellValue(aux);
					sheet.autoSizeColumn(1);
					indice++;
					
					auxLines = lines[11].split(" ");
					aux = auxLines[1];
					
					System.out.println("REFERÊNCIA: "+aux);
					rowData.createCell(2).setCellValue(aux);
					sheet.autoSizeColumn(2);
					indice++;
					if(IRST1 != null) {
						if(IRST1[1] != "a") {
							aux = IRST1[7];
							System.out.println("IRST 1.403: "+aux);
							rowData.createCell(3).setCellValue(aux);
							sheet.autoSizeColumn(3);
							indice++;
							sheet.getRow(0).createCell(7).setCellValue("Outros Impostos 1.403");
							aux = IRST1[8];
							rowData.createCell(7).setCellValue(aux);
							System.out.println("Outros Impostos 1.403: "+aux);
							sheet.autoSizeColumn(7);
						}
						
					}
					
					if(IRST2 != null) {
						if(IRST2[1] != "a") {
							aux = IRST2[7];
							System.out.println("IRST 2.403: "+aux);
							rowData.createCell(4).setCellValue(aux);
							sheet.autoSizeColumn(4);
							indice++;
							sheet.getRow(0).createCell(8).setCellValue("Outros Impostos 2.403");
							aux = IRST2[8];
							rowData.createCell(8).setCellValue(aux);
							System.out.println("Outros Impostos 2.403: "+aux);
							sheet.autoSizeColumn(8);
						}
						
					}
					
					if(IRST3 != null) {
						if(IRST3[1] != "a") {
							aux = IRST3[7];
							System.out.println("IRST 1.406: "+aux);
							rowData.createCell(5).setCellValue(aux);
							sheet.autoSizeColumn(5);
							indice++;
							sheet.getRow(0).createCell(9).setCellValue("Outros Impostos 1.406");
							aux = IRST3[8];
							rowData.createCell(9).setCellValue(aux);
							System.out.println("Outros Impostos 1.406: "+aux);
							sheet.autoSizeColumn(9);
						}
						
					}
					if(IRST4 != null) {
						if(IRST4[1] != "a") {
							aux = IRST4[7];
							System.out.println("IRST 1.407: "+aux);
							rowData.createCell(6).setCellValue(aux);
							sheet.autoSizeColumn(6);
							indice++;
							sheet.getRow(0).createCell(10).setCellValue("Outros Impostos 1.407");
							aux = IRST4[8];
							rowData.createCell(10).setCellValue(aux);
							System.out.println("Outros Impostos 1.407: "+aux);
							sheet.autoSizeColumn(10);
						}
						
					}
					
					
					indice = 0;

				}else if(type == 2) {
					System.out.println("implementando agora...");
					
					System.out.println("a");
					System.out.println("NOME DO CONTRIBUINTE: "); //não há nome do contribuinte no documento
					indice++;
					
					auxLines = lines[2].split(" ");
					aux = auxLines[1];
					
					System.out.println("CNPJ: "+aux);
					rowData.createCell(indice).setCellValue(aux);
					sheet.autoSizeColumn(indice);
					indice++;
					
					auxLines = lines[4].split(" ");
					aux = auxLines[1];
					
					System.out.println("REFERÊNCIA: "+aux);
					rowData.createCell(indice).setCellValue(aux);
					sheet.autoSizeColumn(indice);
					indice++;
					
					auxLines = lines[7].split(" ");
					aux = auxLines[8];
					
					System.out.println("IMPOSTO SAÍDA C/ DÉBITO - 051: "+aux);
					rowData.createCell(5).setCellValue(aux);
					sheet.autoSizeColumn(5);
					
					indice = 0;
				}
				
				else if(type == 3){
					System.out.println("entra no type 3-----------------");
					for(String line : lines) {
						if ( line.contains("1.403")){
							IRST1 = line.split(" ");
							
						}
						if(line.contains("2.403")) {
							IRST2 = line.split(" ");
						}
						if(line.contains("1.406")) {
							
							IRST3 = line.split(" ");
						}
						if(line.contains("1.407")) {
							IRST4 = line.split(" ");
						}
					}
					System.out.println("a");
					String data = getData(lines[22]);
					System.out.println("NOME DO CONTRIBUINTE: "+lines[22]);
					rowData.createCell(0).setCellValue(lines[22]);
					sheet.autoSizeColumn(0);
					indice++;
					
					auxLines = lines[32].split(" ");
					aux = getCNPJ(auxLines);
					
					System.out.println("CNPJ: "+aux);
					rowData.createCell(1).setCellValue(aux);
					sheet.autoSizeColumn(1);
					indice++;
					
					auxLines = lines[32].split(" ");
					aux = getReferencia(auxLines);
					
					System.out.println("REFERÊNCIA: "+aux);
					rowData.createCell(2).setCellValue(aux);
					sheet.autoSizeColumn(2);
					indice++;
					
					if(IRST1 != null) {
						if(IRST1[1] != "a") {
							aux = IRST1[7];
							System.out.println("IRST 1.403: "+aux);
							rowData.createCell(3).setCellValue(aux);
							sheet.autoSizeColumn(3);
							indice++;
							sheet.getRow(0).createCell(7).setCellValue("Outros Impostos 1.403");
							aux = IRST1[8];
							rowData.createCell(7).setCellValue(aux);
							System.out.println("Outros Impostos 1.403: "+aux);
							sheet.autoSizeColumn(7);
						}
						
					}
					
					if(IRST2 != null) {
						if(IRST2[1] != "a") {
							aux = IRST2[7];
							System.out.println("IRST 2.403: "+aux);
							rowData.createCell(4).setCellValue(aux);
							sheet.autoSizeColumn(4);
							indice++;
							sheet.getRow(0).createCell(8).setCellValue("Outros Impostos 2.403");
							aux = IRST2[8];
							rowData.createCell(8).setCellValue(aux);
							System.out.println("Outros Impostos 2.403: "+aux);
							sheet.autoSizeColumn(8);
						}
						
					}
					
					if(IRST3 != null) {
						if(IRST3[1] != "a") {
							aux = IRST3[37];
							System.out.println("IRST 1.406: "+aux);
							sheet.getRow(0).createCell(5).setCellValue("IRST 1.406");
							rowData.createCell(5).setCellValue(aux);
							sheet.autoSizeColumn(5);
							indice++;
							sheet.getRow(0).createCell(9).setCellValue("Outros Impostos 1.406");
							aux = IRST3[38];
							rowData.createCell(9).setCellValue(aux);
							System.out.println("Outros Impostos 1.406: "+aux);
							sheet.autoSizeColumn(9);
						}
						
					}
					if(IRST4 != null) {
						if(IRST4[1] != "a") {
							aux = IRST4[37];
							System.out.println("IRST 1.407: "+aux);
							sheet.getRow(0).createCell(6).setCellValue("IRST 1.407");
							rowData.createCell(6).setCellValue(aux);
							sheet.autoSizeColumn(6);
							indice++;
							sheet.getRow(0).createCell(10).setCellValue("Outros Impostos 1.407");
							aux = IRST4[38];
							rowData.createCell(10).setCellValue(aux);
							System.out.println("Outros Impostos 1.407: "+aux);
							sheet.autoSizeColumn(10);
						}
						
					}
				}else if(type == 4) {
					aux = lines[1].substring(14, lines[1].length());
	
					System.out.println("NOME DO CONTRIBUINTE: "+ aux);
					sheet.getRow(0).createCell(0).setCellValue("NOME DO CONTRIBUINTE");
					rowData.createCell(0).setCellValue(aux);
					sheet.autoSizeColumn(0);
					
					auxLines = lines[2].split(" ");
					aux = auxLines[1];
					System.out.println("CPF/CNPJ: "+aux);
					sheet.getRow(0).createCell(1).setCellValue("CPF/CNPJ");
					rowData.createCell(1).setCellValue(aux);
					sheet.autoSizeColumn(1);
					
					auxLines = lines[3].split(" ");
					aux = auxLines[0];
					System.out.println("PERÍODO DE ESCRITURAÇÃO: "+aux);
					sheet.getRow(0).createCell(2).setCellValue("PERÍODO DE ESCRITURAÇÃO");
					rowData.createCell(2).setCellValue(aux);
					sheet.autoSizeColumn(2);
					
					auxLines = lines[6].split(" ");
					aux = auxLines[7];
					System.out.println("SAÍDA COM DÉBITO: "+aux);
					sheet.getRow(0).createCell(3).setCellValue("SAÍDA COM DÉBITO");
					rowData.createCell(3).setCellValue(aux);
					sheet.autoSizeColumn(3);
					
				}
				else {
				
					System.out.println("Esse tipo ainda não foi implementado!");
				}
				
				}
				
			document.close();
		}catch(Exception e) {
			System.out.println(e);
		}

	}
	public static void searchFilesOnDirectory(String directory, ArrayList<String> directoryList ) {
		File file = new File(directory);
		File files[] = file.listFiles();
		directoryList.add(directory);
		int i = 0;
		while(i != directoryList.size()) {
			File f = new File(directoryList.get(i));
			files = f.listFiles();
			for(int x = 0; x < files.length; x++) {
				if(files[x].isDirectory()) {
					directoryList.add(files[x].toString());
					System.out.println(files[x].toString());
				}
			}
			i++;
		}
		
		
	}
	
	
	public static void main(String args[]){  
		try {
			System.setProperty("sun.java2d.cmm", "sun.java2d.cmm.kcms.KcmsServiceProvider");
			Workbook wb = new XSSFWorkbook();
			String extensionToFind = ".pdf";
			OutputStream fileOut = new FileOutputStream("/treinamento_cloudged/teste novaApuracaoICMS.xlsx");
			Sheet sheet = wb.createSheet();
			
			Row rowTitle = sheet.createRow(0);
			rowTitle.createCell(0).setCellValue("NOME DO CONTRIBUINTE");
			sheet.autoSizeColumn(0);
			rowTitle.createCell(1).setCellValue("CNPJ");
			sheet.autoSizeColumn(1);
			rowTitle.createCell(2).setCellValue("REFERÊNCIA");
			sheet.autoSizeColumn(2);
			rowTitle.createCell(3).setCellValue("IRST 1.403");
			sheet.autoSizeColumn(3);
			rowTitle.createCell(4).setCellValue("IRST 2.403");
			sheet.autoSizeColumn(4);
			rowTitle.createCell(5).setCellValue("IMPOSTO SAÍDA C/ DÉBITO - 051");
			sheet.autoSizeColumn(5);
			
			
			ArrayList<String> directoryList = new ArrayList<>();
			searchFilesOnDirectory("\\\\192.168.0.7\\Controle\\5- PIS - COFINS\\BA\\LUTAN DISTRIBUIDORA DE ALIMENTOS LTDA - 05.156.7130001-62\\APURAÇÃO ICMS\\pages\\1 - para extrair", directoryList);
			
			for(String dir : directoryList) {
				File diretorio = new File(dir);
				File arquivos[] = diretorio.listFiles();
				System.out.println("aaa"+arquivos[0]);
				for(int i = 0; i < arquivos.length; i++) {
					
					int fileSize = arquivos[i].getName().length();
					String nameFile = arquivos[i].getName();
					String extensao = nameFile.substring(fileSize-4, fileSize);
					
					if(extensao.equals(extensionToFind)) {
						System.out.println("-----------"+arquivos[i].getName()+"-----------");
						File inputFile = new File(diretorio.getAbsolutePath()+"\\\\"+arquivos[i].getName());
						convertPDFToExcel(inputFile, sheet);
						line++;
					}
		
				}
			}
			wb.write(fileOut);
			wb.close();
			fileOut.close();
		}catch(Exception e) {
			System.out.println(e);
		}


	}
}


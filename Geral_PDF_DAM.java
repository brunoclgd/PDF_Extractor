
import java.io.File;
import java.io.FileOutputStream;
//import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Queue;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Geral_PDF_DAM {
	public static int line = 1;
	public static int indice = 0;
	
	public static String getData(String data) {
		String formatedData;
		int i;
		i = data.indexOf(" ");
		formatedData = data.substring(0,i)+"/"+data.substring(i+1, data.length());
		return formatedData;
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
				//System.out.println("Text:" + st);
				// split by whitespace
				String lines[] = pdfFileInText.split("\\r?\\n");
				if(lines.length == 236) {
					System.out.println("a");
					String data = getData(lines[5]);
					System.out.println("Inscrição Estadual/CNPJ: "+lines[10]);
					rowData.createCell(indice).setCellValue(lines[10]);
					indice++;
					System.out.println("Período de referência: "+data);
					rowData.createCell(indice).setCellValue(data);
					indice++;
					System.out.println("Saída com débito: "+lines[174]);
					rowData.createCell(indice).setCellValue(lines[174]);
					sheet.autoSizeColumn(indice);
					indice++;

				}
				else if(lines.length == 238 || lines.length == 476) {
					if(lines.length == 238) {
						System.out.println("b");	
					}
					else {
						System.out.println("d");
					}
					String data = getData(lines[5]);
					System.out.println("Inscrição Estadual/CNPJ: "+lines[12]);
					rowData.createCell(indice).setCellValue(lines[12]);
					indice++;
					System.out.println("Período de referência: "+data);
					rowData.createCell(indice).setCellValue(data);
					indice++;
					System.out.println("Saída com débito: "+lines[176]);
					rowData.createCell(indice).setCellValue(lines[176]);
					sheet.autoSizeColumn(indice);
					indice++;
				}
				else if(lines.length == 235){
					System.out.println("c");
					String data = getData(lines[5]);
					System.out.println("Inscrição Estadual/CNPJ: "+lines[10]);
					rowData.createCell(indice).setCellValue(lines[10]);
					indice++;
					System.out.println("Período de referência: "+data);
					rowData.createCell(indice).setCellValue(data);
					indice++;
					System.out.println("Saída com débito: "+lines[173]);
					rowData.createCell(indice).setCellValue(lines[173]);
					sheet.autoSizeColumn(indice);
					indice++;
				}
				else if(lines.length == 237) {
					System.out.println("e");
					String data = getData(lines[5]);
					System.out.println("Inscrição Estadual/CNPJ: "+lines[12]);
					rowData.createCell(indice).setCellValue(lines[12]);
					indice++;
					System.out.println("Período de referência: "+data);
					rowData.createCell(indice).setCellValue(data);
					indice++;
					System.out.println("Saída com débito: "+lines[175]);
					rowData.createCell(indice).setCellValue(lines[175]);
					sheet.autoSizeColumn(indice);
					indice++;
				}
				else {
					System.out.println("não entrou em nenhuma condição!");
				}
				indice = 0;
				


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
			//File diretorio = new File("C:\\Users\\micro\\Desktop\\treinamento\\DAM\\2018");
			//File arquivos[] = diretorio.listFiles();
			String extensionToFind = ".pdf";
			OutputStream fileOut = new FileOutputStream("/treinamento_cloudged/DISFRI 03.345.9350002-89 DAM.xlsx");
			Sheet sheet = wb.createSheet();
			
			Row rowTitle = sheet.createRow(0);
			rowTitle.createCell(0).setCellValue("INSCRIÇÃO ESTADUAL/CNPJ");
			sheet.autoSizeColumn(0);
			rowTitle.createCell(1).setCellValue("PERÍODO DE REFERÊNCIA");
			sheet.autoSizeColumn(1);
			rowTitle.createCell(2).setCellValue("Saída com débito");
			sheet.autoSizeColumn(2);
			
			ArrayList<String> directoryList = new ArrayList<>();
			searchFilesOnDirectory("\\\\192.168.0.7\\Controle\\5- PIS - COFINS\\AC\\DISFRI IMPORTAÇÃO e EXPORTAÇÃO LTDA\\DAM\\03.345.9350002-89\\pages", directoryList);
			
			for(String dir : directoryList) {
				File diretorio = new File(dir);
				File arquivos[] = diretorio.listFiles();
				
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


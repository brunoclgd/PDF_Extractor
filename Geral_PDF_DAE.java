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

public class Geral_PDF_DAE {
	private static int line = 1;
	
	public static void convertPDFToExcel(File inputFile, Sheet sheet) {
		PDDocument document;
		try {
			document = PDDocument.load(inputFile);

			document.getClass();
			if (!document.isEncrypted()) {
				PDFTextStripperByArea stripper = new PDFTextStripperByArea();
				stripper.setSortByPosition(true);
				
				PDFTextStripper tStripper = new PDFTextStripper();
				String pdfFileInText = tStripper.getText(document);
				String lines[] = pdfFileInText.split("\\r?\\n");
				
				String auxLines[] = null;
				String aux = null;
				String razaoSocial = lines[7].substring(14, lines[7].length());
				auxLines = lines[6].split(" ");
				String cnpj = auxLines[4];				
				
				for(int i = 0; i < lines.length; i++) {
					
					if(lines[i].contains("1145 ICMS")) {
						Row rowData = sheet.createRow(line);
						System.out.println("RAZÃO SOCIAL: "+razaoSocial);
						rowData.createCell(0).setCellValue(razaoSocial);
						
						System.out.println("CNPJ: "+cnpj);
						rowData.createCell(1).setCellValue(cnpj);
						
						auxLines = lines[i].split(" ");
						if(auxLines[0].contains("/")) {
							System.out.println("PAGAMENTO: "+auxLines[0]);
							rowData.createCell(2).setCellValue(auxLines[0]);
							
							System.out.println("REFERÊNCIA: "+auxLines[1]);
							rowData.createCell(3).setCellValue(auxLines[1]);
							
							System.out.println("VALOR PRINCIPAL: "+auxLines[6]);
							rowData.createCell(4).setCellValue(auxLines[6]);
						}
						else {
							System.out.println("PAGAMENTO: "+auxLines[1]);
							rowData.createCell(2).setCellValue(auxLines[1]);
							
							System.out.println("REFERÊNCIA: "+auxLines[2]);
							rowData.createCell(3).setCellValue(auxLines[2]);
							
							System.out.println("VALOR PRINCIPAL: "+auxLines[7]);
							rowData.createCell(4).setCellValue(auxLines[7]);
						}
						line++;
						
					}
				}
				
				
				
			}
			document.close();
		}
		catch(Exception e) {
			System.out.println(e.getStackTrace());
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
			OutputStream fileOut = new FileOutputStream("/treinamento_cloudged/LUTAN DAE.xlsx");
			Sheet sheet = wb.createSheet();
			
			Row rowTitle = sheet.createRow(0);
			rowTitle.createCell(0).setCellValue("RAZÃO SOCIAL");
			sheet.autoSizeColumn(0);
			rowTitle.createCell(1).setCellValue("CNPJ");
			sheet.autoSizeColumn(1);
			rowTitle.createCell(2).setCellValue("PAGAMENTO");
			sheet.autoSizeColumn(2);
			rowTitle.createCell(3).setCellValue("REFERÊCIA");
			sheet.autoSizeColumn(3);
			rowTitle.createCell(4).setCellValue("VALOR PRINCIPAL");
			sheet.autoSizeColumn(4);			
			
			ArrayList<String> directoryList = new ArrayList<>();
			searchFilesOnDirectory("\\\\192.168.0.7\\Controle\\5- PIS - COFINS\\BA\\LUTAN DISTRIBUIDORA DE ALIMENTOS LTDA - 05.156.7130001-62\\DAEs\\documentos para extração", directoryList);
			
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
						sheet.autoSizeColumn(0);
						sheet.autoSizeColumn(1);
						sheet.autoSizeColumn(2);
						sheet.autoSizeColumn(3);
						sheet.autoSizeColumn(4);
						
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
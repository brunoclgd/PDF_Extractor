
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Geral_ApuracaoICMS {
	public static int line = 1;
	public static int indice = 0;

	public static String getData(String data) {
		String formatedData;
		int i;
		i = data.indexOf(" ");
		formatedData = data.substring(0,i)+"/"+data.substring(i+1, data.length());
		return formatedData;
	}

	public static String concatName(String[] line, int index) {
		String name = "";
		for(int i = index; i< line.length; i++) {
			name = name.concat(" "+line[i]);
		}
		return name;
	}
	public static int countByIndexOf(String text, String key) {
		int pos = 0, count = 0;

		if(text != null) {
			pos = text.indexOf(key);
			while(pos >= 0) {
				count++;
				pos = text.indexOf(key, pos+1);
			}
		}
		return count;
	}

	public static double checkValue(String value) {
		double newValue = 0;



		return newValue;
	}

	public static int checkTypeOf(String[] text) {
		int type = 0, i = 0;


		while(text[i].equals(" ") || text[i].contains("ADE Conjunto Cotec/Corat")) {
			i++;
		}
		if(text[i].contains("Ministério da Fazenda")) {
			type = 1;
		}
		else if(text[i].contains("Data de Vencimento")) {
			type = 2;
		}
		return type;
	}

	public static void convertPDFToExcel(File inputFile, Sheet sheet) {
		PDDocument document;
		try {
			document = PDDocument.load(inputFile);

			document.getClass();
			if (!document.isEncrypted()) {

				PDFTextStripperByArea stripper = new PDFTextStripperByArea();
				stripper.setSortByPosition(true);

				PDFTextStripper tStripper = new PDFLayoutTextStripper();
				String pdfFileInText = tStripper.getText(document);
				//System.out.println(pdfFileInText);
				String lines[] = pdfFileInText.split("\\r?\\n");
				Row rowData = sheet.createRow(line);
				String auxLines[] = null;
				String aux = null;

				for(String line: lines) {
					line = line.trim().replaceAll("\\s{2,}", " ");
					if(!line.equals(" ")) {
						if(line.contains("CONTRIBUINTE:")) {
							System.out.println(line);
							auxLines = line.split(" ");
							aux = concatName(auxLines, 6);
							System.out.println("a");
							System.out.println("NOME DO CONTRIBUINTE: "+aux);
							rowData.createCell(0).setCellValue(aux);
							sheet.autoSizeColumn(0);

						}

						else if(line.contains("CNPJ/MF:")) {
							auxLines = line.split(" ");
							aux = auxLines[9];
							System.out.println("CNPJ: "+aux);
							rowData.createCell(1).setCellValue(aux);
							sheet.autoSizeColumn(1);

						}

						else if(line.contains("1403")) {
							auxLines = line.split(" ");
							aux = auxLines[10];
							System.out.println("FONTE CFOP 1403: "+aux);
							rowData.createCell(2).setCellValue(aux);
							sheet.autoSizeColumn(2);
							//indice++;
						}
						else if(line.contains("1407")) {
							auxLines = line.split(" ");
							aux = auxLines[10];
							System.out.println("FONTE CFOP 1407: "+aux);
							rowData.createCell(3).setCellValue(aux);
							sheet.autoSizeColumn(3);

						}
						else if(line.contains("2403")) {
							auxLines = line.split(" ");
							aux = auxLines[10];
							System.out.println("FONTE CFOP 2403: "+aux);
							rowData.createCell(4).setCellValue(aux);
							sheet.autoSizeColumn(4);

						}
						else if(line.contains("2407")) {
							auxLines = line.split(" ");
							aux = auxLines[10];
							System.out.println("FONTE CFOP 2407: "+aux);
							rowData.createCell(5).setCellValue(aux);
							sheet.autoSizeColumn(5);

						}
						else if(line.contains("SAÍDAS E PRESTAÇÕES COM DÉBITO DO IMPOSTO")) {
							auxLines = line.split(" ");
							aux = auxLines[9];
							System.out.println("SAÍDA C/ DÉB. DO IMPOSTO :" + aux);
							rowData.createCell(6).setCellValue(aux);
							sheet.autoSizeColumn(6);
						}


						indice = 0;
					}
				}

			}

			document.close();

		}catch(Exception e) {
			System.out.println(e.getStackTrace()[0].getLineNumber());
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
			OutputStream fileOut = new FileOutputStream("/treinamento_cloudged/Workbook_PDF.xlsx");
			Sheet sheet = wb.createSheet();

			Row rowTitle = sheet.createRow(0);
			rowTitle.createCell(0).setCellValue("NOME DO CONTRIBUINTE");
			sheet.autoSizeColumn(0);
			rowTitle.createCell(1).setCellValue("CNPJ");
			sheet.autoSizeColumn(1);
			rowTitle.createCell(2).setCellValue("FONTE CFOP 1403");
			sheet.autoSizeColumn(2);
			rowTitle.createCell(3).setCellValue("FONTE CFOP 1407");
			sheet.autoSizeColumn(3);
			rowTitle.createCell(4).setCellValue("FONTE CFOP 2403");
			sheet.autoSizeColumn(4);
			rowTitle.createCell(5).setCellValue("FONTE CFOP 2407");
			sheet.autoSizeColumn(5);
			rowTitle.createCell(6).setCellValue("SAÍDA C/ DÉB. DO IMPOSTO");
			sheet.autoSizeColumn(6);


			ArrayList<String> directoryList = new ArrayList<>();
			searchFilesOnDirectory("C:\\Users\\micro\\Desktop\\treinamento\\Apuração ICMS", directoryList);

			for(String dir : directoryList) {
				File diretorio = new File(dir);
				File arquivos[] = diretorio.listFiles();


				for(int i = 0; i < arquivos.length; i++) {

					int fileSize = arquivos[i].getName().length();
					String nameFile = arquivos[i].getName();
					System.out.println(nameFile);
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


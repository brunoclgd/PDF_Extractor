
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

public class Geral_PDF_DARF {

	public static int line = 1;
	public static int indice = 0;

	public static String getData(String data) {
		String formatedData;
		int i=0;
		try {
		i = data.indexOf(" ");
		}
		catch(Exception e) {
			System.out.println("espaços não encontrados");
		}
		formatedData = data.substring(0,i)+"/"+data.substring(i+1, data.length());
		return formatedData;
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

	public static String concatName(String[] line, int index) {
		String name = "";
		for(int i = index; i< line.length; i++) {
			name = name.concat(" "+line[i]);
		}
		return name;
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
		String cliente = "";
		try {
			document = PDDocument.load(inputFile);

			document.getClass();
			if (!document.isEncrypted()) {

				PDFTextStripperByArea stripper = new PDFTextStripperByArea();
				stripper.setSortByPosition(true);

				PDFTextStripper tStripper = new PDFTextStripper();
				String pdfFileInText = tStripper.getText(document);
				document.close();
				String lines[] = pdfFileInText.split("\\r?\\n");
				int qtdDARF = 0;
				int inicialIndex = 0, finalIndex = 0, c = 0, i = 0;
				int type = checkTypeOf(lines);

				if(type == 1) {//COFINS TIPO 1
					qtdDARF = countByIndexOf(pdfFileInText, lines[6]);
					System.out.println("aa"+qtdDARF);
					for(int x = 0; x < qtdDARF; x++) {
						Row rowData = sheet.createRow(line);
												
						inicialIndex = finalIndex+1;
						for (int a= inicialIndex; a < lines.length; a ++) {
							
							if(lines[a].contains("Comprovante emitido com base no ADE Conjunto Cotec/Corat no 02, de 07 de novembro de 2006.")) {
								finalIndex = a;
								System.out.println("inicial = "+inicialIndex+"\nfinal = "+finalIndex);
								
								break;

							}

						}
						i = inicialIndex;
						while(i < finalIndex) {
							if(!lines[i].equals(" ")) {
								
								String auxLines[] = null;
								String IRST1[] = null;
								String IRST2[] = null;
								String aux = null;

								if(lines[i].contains("Contribuinte")) {
									i++;
									while(lines[i].equals(" ")) {
										i++;
									}
									aux = lines[i];
									System.out.println("a");
									System.out.println("NOME DO CONTRIBUINTE: "+aux);
									cliente = aux;
									rowData.createCell(0).setCellValue(aux);
									sheet.autoSizeColumn(0);
									

								}
								else if(lines[i].contains("CNPJ")) {
									if(lines[i].length() >= 30) {
										
										auxLines = lines[i].split(" ");
										aux = auxLines[6];
									}else {
										aux = lines[i+1];
									}

									System.out.println("CNPJ: "+aux);
									rowData.createCell(1).setCellValue(aux);
									sheet.autoSizeColumn(1);
									

								}
								else if(lines[i].contains("Período de Apuração")) {
									if(lines[i].length() >= 21) {
										auxLines = lines[i].split(" ");
										aux = auxLines[3];
									}else {
										aux = lines[i+1];
									}

									System.out.println("PERÍODO DE APURAÇÃO: "+aux);
									rowData.createCell(2).setCellValue(aux);
									sheet.autoSizeColumn(2);
									
								}

								else if(lines[i].contains("Data de Arrecadação")) {
									if(lines[i].length() >= 21) {
										auxLines = lines[i].split(" ");
										aux = auxLines[3];
									}else {
										aux = lines[i+1];
									}
									System.out.println("DATA DE ARRECADAÇÃO: "+aux);
									rowData.createCell(3).setCellValue(aux);
									sheet.autoSizeColumn(3);
									
								}

								if(lines[i].contains("6912:")) {
									if(lines[i].length() >= 33) {
										auxLines = lines[i].split(" ");
										aux = auxLines[6];
									}else {
										aux = lines[i+1];
									}

									System.out.println("VAL. IMPOSTO C. 6912: "+aux);
									rowData.createCell(4).setCellValue(aux);
									sheet.autoSizeColumn(4);
									

								}


								if(lines[i].contains("5856:")) {
									if(lines[i].length() >= 33) {
										auxLines = lines[i].split(" ");
										aux = auxLines[6];
									}else {
										aux = lines[i+1];
									}

									System.out.println("VAL. IMPOSTO C. 5856: "+aux);
									rowData.createCell(5).setCellValue(aux);
									sheet.autoSizeColumn(5);
									

								}

								indice = 0;

							}
							i++;
						}
						line++;
					}
				}
				else if(type == 2) {//COFINS TIPO 2
					String auxQtdL[] = lines[1].split(" ");
					String auxQtdOcurrency = concatName(auxQtdL, 1);
					qtdDARF = countByIndexOf(pdfFileInText, auxQtdOcurrency);
					System.out.println(qtdDARF);
					inicialIndex = 0;
					for(int x = 0; x < qtdDARF; x++) {
						Row rowData = sheet.createRow(line);
						
						inicialIndex = finalIndex+1;
						for (int a= inicialIndex; a < lines.length; a ++) {

							if(lines[a].contains("ADE Conjunto Cotec/Corat")) {
								finalIndex = a+7;
								System.out.println("inicial = "+inicialIndex+"\nfinal = "+finalIndex);
								break;

							}

						}
						i = inicialIndex;

						
						String auxLines[] = null;
						String IRST1[] = null;
						String IRST2[] = null;
						String aux = null;

						System.out.println(lines[inicialIndex]);
							auxLines = lines[inicialIndex].split(" ");
							aux = auxLines[1];
						
						
						for(int s = 2; s < auxLines.length; s++) {
							aux = aux.concat(" "+auxLines[s]);
						}


						System.out.println("NOME DO CONTRIBUINTE: "+aux);
						cliente = aux;
						rowData.createCell(0).setCellValue(aux);
						sheet.autoSizeColumn(0);
						

						aux = auxLines[0];
						System.out.println("CNPJ: "+aux);
						rowData.createCell(1).setCellValue(aux);
						sheet.autoSizeColumn(1);
						

						auxLines = lines[inicialIndex+1].split(" ");
						aux = auxLines[0];

						System.out.println("PERÍODO DE APURAÇÃO: "+aux);
						rowData.createCell(2).setCellValue(aux);
						sheet.autoSizeColumn(2);
						

						aux = auxLines[1];
						System.out.println("DATA DE ARRECADAÇÃO: "+aux);
						rowData.createCell(3).setCellValue(aux);
						sheet.autoSizeColumn(3);
						
						
						for(int f = inicialIndex; f < finalIndex; f++) {
							if(lines[f].contains("6912")) {
								auxLines = lines[f].split(" ");
								aux = auxLines[7];
								System.out.println("VAL. IMPOSTO C. 6912: "+aux);
								rowData.createCell(4).setCellValue(aux);
								sheet.autoSizeColumn(4);
								

							}
							
							if(lines[f].contains("8109")) {
								auxLines = lines[f].split(" ");
								aux = auxLines[4];
								System.out.println("VAL. IMPOSTO C. 8109: "+aux);
								rowData.createCell(5).setCellValue(aux);
								sheet.autoSizeColumn(5);
								
							}
							if(lines[f].contains("5856")) {
								auxLines = lines[f].split(" ");
								aux = auxLines[3];
								System.out.println("VAL. IMPOSTO C. 5856: "+aux);
								rowData.createCell(6).setCellValue(aux);
								sheet.autoSizeColumn(6);
								

							}
							if(lines[f].contains("2172")) {
								auxLines = lines[f].split(" ");
								aux = auxLines[8];
								System.out.println("VAL. IMPOSTO C. 2172: "+aux);
								rowData.createCell(7).setCellValue(aux);
								sheet.autoSizeColumn(7);
								
							}
							
						}
						
						
					}
				}
				else {//TIPO NÃO IDENTIFICADO!!
					System.out.println("tipo não identificado!");
				}
			}

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
	public static void main(String args[]) throws IOException{  
	//	try {
			System.setProperty("sun.java2d.cmm", "sun.java2d.cmm.kcms.KcmsServiceProvider");
			Workbook wb = new XSSFWorkbook();
			//File diretorio = new File("C:\\Users\\micro\\Desktop\\treinamento\\DAM\\2018");
			//File arquivos[] = diretorio.listFiles();
			String extensionToFind = ".pdf";
			OutputStream fileOut = new FileOutputStream("/treinamento_cloudged/DISFRI PIS.xlsx");
			Sheet sheet = wb.createSheet();

			Row rowTitle = sheet.createRow(0);
			rowTitle.createCell(0).setCellValue("NOME DO CONTRIBUINTE");
			sheet.autoSizeColumn(0);
			rowTitle.createCell(1).setCellValue("CNPJ");
			sheet.autoSizeColumn(1);
			rowTitle.createCell(2).setCellValue("PERÍODO DE APURAÇÃO");
			sheet.autoSizeColumn(2);
			rowTitle.createCell(3).setCellValue("DATA DE ARRECADAÇÃO");
			sheet.autoSizeColumn(3);
			rowTitle.createCell(4).setCellValue("VAL. IMPOSTO C.6912");
			sheet.autoSizeColumn(4);
			rowTitle.createCell(5).setCellValue("VAL. IMPOSTO C. 8109");
			sheet.autoSizeColumn(5);
			rowTitle.createCell(6).setCellValue("VAL. IMPOSTO C. 5856");
			sheet.autoSizeColumn(6);
			rowTitle.createCell(7).setCellValue("VAL.IMPOSTO C. 2172");
			sheet.autoSizeColumn(7);

			ArrayList<String> directoryList = new ArrayList<>();
			searchFilesOnDirectory("\\\\192.168.0.7\\Controle\\5- PIS - COFINS\\AC\\DISFRI IMPORTAÇÃO e EXPORTAÇÃO LTDA\\DARF\\PIS\\pages", directoryList);
			
			for(String dir : directoryList) {
				File diretorio = new File(dir);
				File arquivos[] = diretorio.listFiles();
			
				
				for(int i = 0; i < arquivos.length; i++) {

					int fileSize = arquivos[i].getName().length();
					String nameFile = arquivos[i].getName();
					//System.out.println(nameFile);
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
		//}catch(Exception e) {
		//	System.out.println(e);
		//}


	//}
}

}

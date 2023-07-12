package cdp;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class BuscaModulo {
	
	
	private static List<String> programasCobol = new ArrayList<>();

    public static void main(String[] args) throws Exception {
        File pastaPrincipal = new File("C:\\Viegas\\desenvolvimento\\BoaVista\\demandas\\buscaModulos\\code-as400-main");
        String relatorio = "C:\\Viegas\\desenvolvimento\\BoaVista\\demandas\\buscaModulos\\Panilha.xlsx";
        lerPlanilha(relatorio);
        if (programasCobol != null && !programasCobol.isEmpty()) {
        	
        	for (String programaCobol: programasCobol) {
        		
        		procurarArquivo(pastaPrincipal, programaCobol);
        	}
        	
        }
        
    }

    public static void procurarArquivo(File pasta, String nomeArquivo) {
        if (pasta.isDirectory()) {
            File[] arquivos = pasta.listFiles();
            if (arquivos != null) {
                for (File arquivo : arquivos) {
                    if (arquivo.isDirectory()) {
                        procurarArquivo(arquivo, nomeArquivo);
                    } else if (arquivo.getName().equalsIgnoreCase(nomeArquivo)) {
                        //System.out.println("Arquivo encontrado: " + arquivo.getAbsolutePath());
                    	encontrarPalavras(arquivo);
                    }
                }
            }
        }
    }
    
    public static List<String> encontrarPalavras(File arquivo) {
        List<String> palavrasEncontradas = new ArrayList<>();
        Pattern padrao = Pattern.compile("\\bT[\\p{Alnum}]{7}\\b");

        try (BufferedReader br = new BufferedReader(new FileReader(arquivo))) {
            String linha;
            while ((linha = br.readLine()) != null) {
                Matcher matcher = padrao.matcher(linha);
                while (matcher.find()) {
                    String palavraEncontrada = matcher.group();
                    System.out.println(arquivo.getName() + " " + palavraEncontrada);
                    palavrasEncontradas.add(palavraEncontrada);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return palavrasEncontradas;
    }
    
    @SuppressWarnings("resource")
	public static void lerPlanilha(String relatorio) throws Exception {

        try {
               FileInputStream arquivo = new FileInputStream(new File(
            		   relatorio));
               
               File arquivoExcel = new File(relatorio);
               
               DataFormatter dataFormatter = new DataFormatter();
               
               if (arquivoExcel.exists() && arquivoExcel.isFile() && arquivoExcel.length() > 0) {
            	   
            	   OPCPackage pkg = OPCPackage.open(new File(relatorio));
            	   
            	   XSSFWorkbook workbook = new XSSFWorkbook(pkg);
            	   
            	   XSSFSheet sheetPedidos =  workbook.getSheetAt(0);
            	   
            	   Iterator<Row> rowIterator = sheetPedidos.iterator();
            	   
            	   while (rowIterator.hasNext()) {
            		   
            		   Row row = rowIterator.next();
            		   
            		   if (row.getRowNum() == 0) {
            			   continue;
            		   }
            		   
            		   Iterator<Cell> cellIterator = row.cellIterator();
            		   while (cellIterator.hasNext()) {
        				   Cell cell = cellIterator.next();
        				   
        				   // Na última linha do relatório existe a frase: Accenture Highly Confidential.  Copyright ©2022
        				   // Daí dá erro porque estou tratando a célula com o cell.getNumericCellValue() e como é uma sequência de caracteres, dá erro.
        				   // Então ignoro esse erro, pois sei que é só uma frase que está ao final do arquivo
        				   /*
        				   if (cell.getColumnIndex() == 0) {
        					   try {
        						   String cDPTrackerID = String.valueOf(cell.getNumericCellValue());
        					   } catch (Exception e) {
        						   continue;
        					   }
        				   }
        				   */
        				   
        				   switch (cell.getColumnIndex()) {
        				   case 0:
        					   programasCobol.add(cell.getStringCellValue());
        					   break;
        					   
        				   }

        			   }
            		   
            	   }
            	   arquivo.close();
            	   
               }

		} catch (Exception e) {
			throw new Exception("Arquivo Excel nao encontrado! : " + e);
		}
        
  }

}

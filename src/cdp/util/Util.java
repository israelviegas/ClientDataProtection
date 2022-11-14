package cdp.util;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Calendar;
import java.util.Date;
import java.util.Properties;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.WebDriverWait;


public class Util {
	
    public static String getValor(String chave) throws IOException{
    	Properties props = getProp();
        return (String)props.getProperty(chave);
    }
    
    // Arquivo de properties sendo usado fora do projeto
    public static Properties getProp() throws IOException {
        Properties props = new Properties();
        
        // String arquivoProperties = "D:/JOBS/AutomacaoAdquira/configuracoes/propriedadesAdquira.properties"; 
        // String arquivoProperties = "C:/AutomacaoClientDataProtection/configuracoes/propriedadesClientDataProtection.properties"; 
         String arquivoProperties = "C:/Viegas/desenvolvimento/Selenium/ClientDataProtection/configuracoes/propriedadesClientDataProtection.properties";
         
        FileInputStream file = new FileInputStream(arquivoProperties);
        props.load(file);
        return props;
    }
    
    /*
	public static void converteValorNullParaEspacoEmBranco(Pedido pedido) {
		
		if(pedido.getContractNumber().getContrato()==null || pedido.getContractNumber().getContrato().isEmpty()	){          pedido.getContractNumber().setContrato(" ");   				}
		if(pedido.getContractNumber().getFrente()==null	|| pedido.getContractNumber().getFrente().isEmpty() ){          pedido.getContractNumber().setFrente(" ");   					}
		if(pedido.getContractNumber().getNumero()==null	|| pedido.getContractNumber().getNumero().isEmpty() ){         	pedido.getContractNumber().setNumero(" ");    		}
		if(pedido.getContractNumber().getWbs()==null || pedido.getContractNumber().getWbs().isEmpty() ){			        pedido.getContractNumber().setWbs(" ");  }
		if(pedido.getNumero()==null || pedido.getNumero().isEmpty()	){          pedido.setNumero(" ");   					}
		if(pedido.getData()==null || pedido.getData().isEmpty() 	){          pedido.setData(" ");   			}
		if(pedido.getCnpjCliente()==null || pedido.getCnpjCliente().isEmpty() ){          pedido.setCnpjCliente(" ");   			}
		if(pedido.getComprador()==null	|| pedido.getComprador().isEmpty() 	){          pedido.setComprador(" ");	      			}
		if(pedido.getMensagemErroRegraPreenchimento()==null	|| pedido.getMensagemErroRegraPreenchimento().isEmpty()	){          pedido.setMensagemErroRegraPreenchimento(" ");	    	}
		if(pedido.getObservacaoSharepoint()==null || pedido.getObservacaoSharepoint().isEmpty()	){          pedido.setObservacaoSharepoint(" ");	      			}

	}
	*/
    
    public static boolean possuiValor(String valor){
    	
    	boolean possuiValor = false;

    	if (valor != null && !valor.isEmpty() && !"null".equals(valor)) {
    		
    		possuiValor = true;
    		
    	}
		return possuiValor;
    }

    
    public static void criaDiretorio(String caminhoDiretorio){
        File diretorio = new File(caminhoDiretorio);
        if (!diretorio.exists()) {
        	diretorio.mkdirs();
        }
    }

	   public static void criaDiretorioTemp(){
	    	
		   	String str1 = "echo %temp%";
		   	String command = "C:\\WINDOWS\\system32\\cmd.exe /y /c " + str1;
		    	
		   	try {
		    		
		   		Process processo = Runtime.getRuntime().exec(command);
		   		String line;
		   		String caminhoTemp = "";
		    		
		   		//pega o retorno do processo
		   		BufferedReader stdInput = new BufferedReader(new 
		   				InputStreamReader(processo.getInputStream()));
		    		
		   		//printa o retorno
		   		while ((line = stdInput.readLine()) != null) {
		   			caminhoTemp = line;
		   		}
		    		
		   		criaDiretorio(caminhoTemp);
		    		
		   	} catch (Exception e) {
		   		System.out.println("Deu erro na criacao do diretorio Temp: " + e.getMessage());
		   	}

	   }
	   
	    public static void gravarArquivo(String caminhoDiretorio, String nomeArquivo, String extensaoArquivo, String conteudoArquivo, String mensagem) throws IOException {
	    	
	    	String arquivo = caminhoDiretorio + "/" + nomeArquivo + extensaoArquivo; 
	    	File file = new File(arquivo);
			BufferedWriter writer = new BufferedWriter(new FileWriter(file));
			writer.write(mensagem + conteudoArquivo);
			writer.newLine();
			//Criando o conteudo do arquivo
			writer.flush();
			//Fechando conexao e escrita do arquivo.
			writer.close();
			
	    }
	    
		  public static void apagaArquivosDiretorioDeRelatorios(String caminhoDiretorio) throws Exception{
		    	
		    	boolean sucesso = false;
		        File diretorio = new File(caminhoDiretorio);
		        if (diretorio.exists() && diretorio.isDirectory()) {
		        	sucesso = true;
		        	
		        	//lista os nomes dos arquivos
					String arquivos [] = diretorio.list();
					
					if (arquivos != null && arquivos.length > 0) {
						
						for (String item : arquivos){
							
							File arquivo = new File(caminhoDiretorio + "/" + item);
							// Se existirem arquivos, os deleto
							if (arquivo.exists() && arquivo.isFile()) {
								arquivo.delete();
							}

						}
					}
		        	
		        }
		        
		        if (!sucesso) {
		        	throw new Exception("Nao existe o diretorio: " + caminhoDiretorio);
		        }
		        
		    }
		  
		   public static void moverArquivosEntreDiretorios(String caminhoArquivoOrigem, String caminhoDiretorioDestino) throws Exception{
		    	
		    	boolean sucesso = true;
		    	File arquivoOrigem = new File(caminhoArquivoOrigem);
		        File diretorioDestino = new File(caminhoDiretorioDestino);
		        if (arquivoOrigem.exists() && diretorioDestino.exists()) {
		        	sucesso = arquivoOrigem.renameTo(new File(diretorioDestino, arquivoOrigem.getName()));
		        }
		        
		        if (!sucesso) {
		        	throw new Exception("Ocorreu um erro no momento de mover o relat�rio " + caminhoArquivoOrigem + " para " + caminhoDiretorioDestino);
		        	
		        }
		        
		    }
		  
		  public static void apagaDiretoriosDeRelatorios(String caminhoDiretorio) throws Exception{
		    	
		    	boolean sucesso = false;
		        File diretorio = new File(caminhoDiretorio);
		        Date dataAtual = new Date();
		        Calendar cal = Calendar.getInstance();
		        cal.setTime(dataAtual);
		        cal.add(Calendar.DATE, -7);
		        Date dataAntes7Dias = cal.getTime();
		        
		        if (diretorio.exists() && diretorio.isDirectory()) {
		        	sucesso = true;
		        	
		        	//lista os nomes dos diret�rios
					String itens [] = diretorio.list();
					
					if (itens != null && itens.length > 0) {
						
						for (String item : itens){
							
							File pasta = new File(caminhoDiretorio + "/" + item);
							
							if (pasta.exists() && pasta.isDirectory()) {
								
								Long dataModificacaoPasta =  FileUtils.lastModified(pasta);
								Date dataModificacaoPasta2 = new Date(dataModificacaoPasta);
								
								// Se existirem diret�rios com a data anterior � data de 7 dias atr�s, os deleto
								if (dataModificacaoPasta2.before(dataAntes7Dias)) {
									FileUtils.deleteQuietly(pasta);
								}
								
							}

						}

					}
		        	
		        }
		        
		        if (!sucesso) {
		        	throw new Exception("Nao existe o diretorio: " + caminhoDiretorio);
		        }
		        
		    }

		  public static boolean existeArquivo(String caminhoArquivo) throws Exception{
		    	
		    	boolean existeArquivo = false;
		        File arquivo = new File(caminhoArquivo);
		        if (arquivo.exists() && !arquivo.isDirectory()) {
		        	existeArquivo = true;
		        }
				return existeArquivo;
		    }



}
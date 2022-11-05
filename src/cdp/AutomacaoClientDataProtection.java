package cdp;

import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import java.awt.Color;
import java.awt.Robot;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.math.BigDecimal;
import java.nio.file.attribute.BasicFileAttributes;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.GregorianCalendar;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxOptions;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.ie.InternetExplorerOptions;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.interactions.PointerInput;
import org.openqa.selenium.interactions.PointerInput.Origin;
import org.openqa.selenium.interactions.Sequence;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import cdp.dao.HistoricoExecucaoDao;
import cdp.util.Util;

public class AutomacaoClientDataProtection {
	
	private static String nomeRelatorioBaixado;
	private static String nomeZipBaixado;
	private static boolean extracaoPossuiPedidos;
	private static boolean existemPedidos = false;
	private static String dataAtual = new SimpleDateFormat("yyyy_MM_dd HH_mm_ss").format(new Date());
	private static String dataAtualPlanilhaFinal = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss").format(new Date());
	private static String dataAtualSharepoint = new SimpleDateFormat("MM/dd/yyyy").format(new Date());
	private static String diretorioLogs = "";
	private static String subdiretorioPdfsBaixados = "";
//	private static List<ContractNumber> listaContractNumbers = null;
//	private static List<ContractNumber> listaContractNumbersTemporaria = null;
	private static Set<String> listaNumerosContractNumbersDistintos = null;
//	private static List<Pedido> listaPedidos = null;
//	private static List<Pedido> listaPedidosFaturados = null;
//	private static List<Pedido> listaPedidosNaoFaturados = null;
//	private static List<Pedido> listaPedidosNaoFaturadosAuxiliar = null;
	private static String listaPedidosComErrosNasRegraDePreenchimentoNoSharePoint = "";
	private static int contadorErros;
	private static int contadorErroslerRelatorioExcel = 0;
	private static int contadorfazerDownlodRelatorioPorPeriodo = 0;
	private static int contadorErrosMoverArquivos = 0;
	private static int contadorErrosLogin = 0;
	private static int contadorErrosLogout = 0;
	private static int contadorErrosRecuperaContractNumbersSharepoint = 0;
	private static int contadorErrosRecuperaPedidosSharepoint = 0;
	private static int contadorErrosPreencherCamposBiling = 0;
	private static int contadorExecutaAutomacaoClientDataProtection = 0;
	private static int contadorLogin = 0;
	private static String diretorioRelatorio = null;
	private static String subdiretorioRelatoriosBaixados = null;
	private static String subdiretorioRelatoriosBaixados2 = null;
	private static String subdiretorioRelatorioFinal = null;
	private static String subdiretorioRelatorioFinal2  = null;
	private static String subdiretorioRelatorioIncremental = null;   
	private static String subdiretorioPdfsBaixados2  = null;
	private static String caminhoExecutavelPlanilhaContractNumbers = null;
	private static String caminhoExecutavelPlanilhaPedidosFaturados = null;
	
	
	@SuppressWarnings("unused")
	public static void main(String[] args) throws Exception {
		
		WebDriver driver = null;

    		try {
    			diretorioLogs = Util.getValor("caminho.diretorio.relatorios") + "/" + dataAtual;
    			subdiretorioRelatoriosBaixados = Util.getValor("caminho.download.relatorios") + "\\" + dataAtual;
    			Util.criaDiretorio(subdiretorioRelatoriosBaixados);
    			
    			// As vezes o diretorio que armazena dados temporarios do Chome simplesmente some, dai o Selenium da pau na hora de chamar o browser
    			// Com o metodo abaixo, crio essa pasta se ela nao existir
    			Util.criaDiretorioTemp();

    			// Deleta os diretorios que possuirem data de criacao anterior a data de 7 dias atras
    			Util.apagaDiretoriosDeRelatorios(Util.getValor("caminho.download.relatorios"));
    			
    			executaAutomacaoClientDataProtection(driver);
            
    		} catch (Exception e) {
    			Util.gravarArquivo(diretorioLogs, "Erro ClientDataProtection" + " " + dataAtual, ".txt", e.getMessage(), "Ocorreu um erro na automacao de extracao de relatorios: ");
    			inserirStatusExecucaoNoBanco("ClientDataProtection", dataAtualPlanilhaFinal, "Erro de execucao do robo");
    		} finally {
    			if (driver != null) {
    				driver.quit();
    			}
    			
    			mataProcessosGoogle();
    			mataProcessosFirefox();
				
			}
    		
	}
	
    public static void executaAutomacaoClientDataProtection(WebDriver driver) throws Exception{
    	
    	try {
    		
    		// Deleto arquivos que existirem no diretorio de relatorios
    		Util.apagaArquivosDiretorioDeRelatorios(Util.getValor("caminho.download.relatorios"));
    		
    		// Deleto arquivos que existirem no subdiretorio de relatorios
    		Util.apagaArquivosDiretorioDeRelatorios(subdiretorioRelatoriosBaixados);
    		
    		if (driver != null) {
    			driver.quit();
    		}
    		
    		mataProcessosGoogle();
    		mataProcessosFirefox();
    		
    		System.out.println("Inicio: " + new SimpleDateFormat("dd_MM_yyyy HH_mm_ss").format(new Date()));
    		String mensagemResultadoClientDataProtection = "Sucesso na extração dos relatórios!";
    		
    		driver = getWebDriver();
    		JavascriptExecutor js = (JavascriptExecutor) driver;
    		WebDriverWait wait = new WebDriverWait(driver, 60);
    		
    		// Acessa o site Client Data Protection
    		acessarClientDataProtection(driver, wait, js);
    		
    		// Clicando no link TELEFONICA GROUP
    		clicarLinkTelefonicaGroup(driver, wait, js);
    		
    		// Faz o download do relatório Controls Due This Month
    		fazerDownloadRelatorioControlsDueThisMonth(driver, wait, js);
    		
    		// Acessa o site Client Data Protection na aba original
    		//((JavascriptExecutor) driver).executeScript("window.open()");
    		List<String> windowHandles = new ArrayList<>(driver.getWindowHandles());
    		driver.switchTo().window(windowHandles.get(0));
    		acessarClientDataProtection(driver, wait, js);
    		
    		// Faz o download do relatório Operational Risk Index By Client
    		fazerDownloadRelatorioOperationalRiskIndexByClient(driver, wait, js);
    		
    		// Faz o download do relatório 
    		/*
    		fazerDownloadRelatorioOperationalRiskIndexByClient(driver, wait, js);
    		
    		if (extracaoPossuiPedidos) {
    			existemPedidos = true;
    			//Move o relatorio baixado do diretorio relatorios para o diretorio correto
    			//Util.moverArquivosEntreDiretorios(driver, wait, js, Util.getValor("caminho.download.relatorios") + "\\" + nomeRelatorioBaixado, subdiretorioRelatoriosBaixados);
    			//Thread.sleep(5000);
    			
    			// Le o relatorio baixado
    			//lerRelatorioExcel(driver, wait, js, subdiretorioRelatoriosBaixados2 + "/" + nomeRelatorioBaixado, subdiretorioRelatoriosBaixados);
    			
    		}
    		
    		// Gravo em um arquivo os Pedidos que tiveram problemas nas regras de preenchimento
    		if (listaPedidosComErrosNasRegraDePreenchimentoNoSharePoint != null && !listaPedidosComErrosNasRegraDePreenchimentoNoSharePoint.isEmpty()) {
    			//Util.gravarArquivo(subdiretorioRelatorioFinal2, "Pedidos com Erros nas Regras de Preenchimento no Sharepoint" + " " + dataAtual, ".txt", listaPedidosComErrosNasRegraDePreenchimentoNoSharePoint, "");
    		}
    		*/
    		
    		Util.gravarArquivo(diretorioLogs, "Sucesso ClientDataProtection" + " " + dataAtual, ".txt", "", mensagemResultadoClientDataProtection);
    		
    		// Grava na tabela Tb_Historico_Execucao_Robos o servi�o, data e hora e status da execu��o
    		//inserirStatusExecucaoNoBanco("ClientDataProtection", dataAtualPlanilhaFinal, mensagemResultadoAdquira);
    		
    		//mensagemSucesso();
    		
    		System.out.println("Fim: " + new SimpleDateFormat("dd_MM_yyyy HH_mm_ss").format(new Date()));
    		
    		if (driver != null) {
    			driver.quit();
    		}
    		
    		mataProcessosGoogle();
    		mataProcessosFirefox();
			
		} catch (Exception e) {
			contadorExecutaAutomacaoClientDataProtection ++;
			// Executo ate 20 vezes se der erro no executaAutomacaoAdquiraSharepoint
			if (contadorExecutaAutomacaoClientDataProtection <= 20) {
				
				System.out.println("Deu erro no metodo executaAutomacaoClientDataProtection, tentativa de acerto: " + contadorExecutaAutomacaoClientDataProtection);
				executaAutomacaoClientDataProtection(driver);
			
			} else {
				throw new Exception("Ocorreu um erro no metodo executaAutomacaoClientDataProtection: " + e);
		    }

		}
    	
    }
    
    public static void acessarClientDataProtection(WebDriver driver, WebDriverWait wait, JavascriptExecutor js) throws Exception {

    	// Abrindo a URl do Adquira
    	driver.manage().window().maximize();
    	driver.get(Util.getValor("url.cdp"));
    	Thread.sleep(3000);
  
    }
    
    public static void clicarLinkTelefonicaGroup(WebDriver driver, WebDriverWait wait, JavascriptExecutor js) throws Exception {
    	// Clicando no link TELEFONICA GROUP
    	String textoTelefonicaGroup = "TELEFONICA GROUP";
    	wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span [text()='"+textoTelefonicaGroup+"']"))).click();
    	Thread.sleep(5000);
    }
    
    public static void fazerDownloadRelatorioControlsDueThisMonth(WebDriver driver, WebDriverWait wait, JavascriptExecutor js) throws Exception {
    	
    	// Clicando no link  Show Summary Statistics & CDP Operational Compliance
    	String idOlhinho = "showSummaryTableText";
    	wait.until(ExpectedConditions.elementToBeClickable(By.id(idOlhinho))).click();
    	Thread.sleep(5000);
    	
    	// Clicando no link Controls Due this Month:
    	String textoControlsDueThisMonth = "Controls Due this Month:";
    	wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span [text()='"+textoControlsDueThisMonth+"']"))).click();
    	
    	// Não consegui interagir com nenhum elemento da página do relatório usando o Selenium
    	// Tive que apelar para o Robot ;)
    	clickBotaoSalvarRelatorio();
    	
    	String nomeRelatorioControlsDueThisMonth = "ControlsDueThisMonth.xlsx";
    	//Move o excel baixado do diretorio relatorios para o diretorio correto
    	Util.moverArquivosEntreDiretorios(Util.getValor("caminho.download.relatorios") + "\\" + nomeRelatorioControlsDueThisMonth, subdiretorioRelatoriosBaixados);
    	Thread.sleep(1000);
    	
    }
    
    public static void fazerDownloadRelatorioOperationalRiskIndexByClient(WebDriver driver, WebDriverWait wait, JavascriptExecutor js) throws Exception {
    	
    	// Clicando no link REPORTS
    	String idReports = "//*[@id=\"oppMobDelTabs\"]/li[2]/span";
    	wait.until(ExpectedConditions.elementToBeClickable(By.xpath(idReports))).click();
    	Thread.sleep(3000);
    	
    	// Clicando no item Operational Risk Index Report do menu
    	String textoOperationalRiskIndexReport = "Operational Risk Index Report";
    	wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//li [text()='"+textoOperationalRiskIndexReport+"']"))).click();
    	
    	// Não consegui interagir com nenhum elemento da página do relatório usando o Selenium
    	// Tive que apelar para o Robot ;)
    	clickBotaoSalvarRelatorio();
    	
    	String nomeRelatorioOperationalRiskIndexByClient = "Operational Risk Index By Client.xlsx";
    	//Move o excel baixado do diretorio relatorios para o diretorio correto
    	Util.moverArquivosEntreDiretorios(Util.getValor("caminho.download.relatorios") + "\\" + nomeRelatorioOperationalRiskIndexByClient, subdiretorioRelatoriosBaixados);
    	Thread.sleep(1000);
    	
    }

    
    public static void moverArquivosEntreDiretorios(WebDriver driver, WebDriverWait wait, JavascriptExecutor js, String caminhoArquivoOrigem, String caminhoDiretorioDestino) throws Exception{
    	
    	boolean sucesso = true;
    	File arquivoOrigem = new File(caminhoArquivoOrigem);
        File diretorioDestino = new File(caminhoDiretorioDestino);
        if (arquivoOrigem.exists() && diretorioDestino.exists()) {
        	sucesso = arquivoOrigem.renameTo(new File(diretorioDestino, arquivoOrigem.getName()));
        }
        
        if (!sucesso) {
        	contadorErrosMoverArquivos++;
        	
            // Tento mover o arquivo por at� 20 vezes
            if (contadorErrosMoverArquivos <= 20) {
            	
				System.out.println("Deu erro no metodo moverArquivosEntreDiretorios, tentativa de acerto: " + contadorErrosMoverArquivos);
				// Est� dando erro de logout no servidor
				// O bot�o de logout est� ficando escondido
				// ent�o retirarei o logout e o login por enquanto
				fazerLogoutAdquira(driver, wait);
				fazerLoginAdquira(driver, wait, js);
				//acessarPaginaInicial(driver, wait);
				fazerDownlodRelatorioPorPeriodo(driver, wait, js);
				moverArquivosEntreDiretorios(driver, wait, js, caminhoArquivoOrigem, caminhoDiretorioDestino);
            
            } else {
            	throw new Exception("Ocorreu um erro no momento de mover o relat�rio " + caminhoArquivoOrigem + " para " + caminhoDiretorioDestino);
            }
        	
        }
        
    }

    
    public static void clickBotaoSalvarRelatorio1(WebDriver driver, int posicaoX, int posicaoY) throws Exception {
    	
    	PointerInput mouse = new PointerInput(PointerInput.Kind.MOUSE, "default mouse");
    	
    	Sequence actions = new Sequence(mouse, 0)
    			.addAction(mouse.createPointerMove(Duration.ofMillis(500), Origin.pointer(), posicaoX, 50))
    			.addAction(mouse.createPointerDown(PointerInput.MouseButton.RIGHT.asArg()))
    			.addAction(mouse.createPointerUp(PointerInput.MouseButton.RIGHT.asArg()));
    	
    	((RemoteWebDriver) driver).perform(Collections.singletonList(actions));
    	
    }
    
    public static void clickBotaoSalvarRelatorio2(WebDriver driver, int posicaoX, int posicaoY) throws Exception {
    	
        Actions builder = new Actions(driver);
        Action moveM = builder.moveByOffset(posicaoX, posicaoY).build();
        moveM.perform();

        Action click = builder.contextClick().build();
        click.perform();
            	
    }
    
    public static void clickBotaoSalvarRelatorio() throws Exception {
    	
    	Robot robot = new Robot();
 //   	robot.mouseMove(0, 0); 
    	
    	Thread.sleep(30000);
    	// Apertando o tab 6 vezes até chegar no botão de salvar
    	for (int i = 1; i <= 6; i++) {
    		robot.keyPress(KeyEvent.VK_TAB);
    		robot.keyRelease(KeyEvent.VK_TAB);
    		Thread.sleep(1000);
    	}
    	
    	// Enter no botão de salvar
    	robot.keyPress(KeyEvent.VK_ENTER);
    	robot.keyRelease(KeyEvent.VK_ENTER);
    	Thread.sleep(1000);
    	
    	// Descendo até a opção de Excel
    	for (int i = 1; i <= 4; i++) {
    		robot.keyPress(KeyEvent.VK_DOWN);
    		robot.keyRelease(KeyEvent.VK_DOWN);
    		Thread.sleep(1000);
    	}
    	
    	// Enter no botão de gerar relatório em excel
    	robot.keyPress(KeyEvent.VK_ENTER);
    	robot.keyRelease(KeyEvent.VK_ENTER);
    	Thread.sleep(20000);
    	
        /*
    	Robot bot = new Robot();
        bot.mouseMove(posicaoX, posicaoY);    
        bot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
        bot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        */
    }

    
    @SuppressWarnings("resource")
	public static void lerRelatorioExcel(WebDriver driver, WebDriverWait wait, JavascriptExecutor js, String relatorio, String subdiretorioRelatoriosBaixados) throws Exception {

        try {
               FileInputStream arquivo = new FileInputStream(new File(
            		   relatorio));
               
               File arquivoExcel = new File(relatorio);
               
               if (arquivoExcel.exists() && arquivoExcel.isFile() && arquivoExcel.length() > 0) {
            	   
            	   OPCPackage pkg = OPCPackage.open(new File(relatorio));
            	   
            	   XSSFWorkbook workbook = new XSSFWorkbook(pkg);
            	   
            	   XSSFSheet sheetPedidos =  workbook.getSheetAt(0);
            	   
            	   Iterator<Row> rowIterator = sheetPedidos.iterator();
            	   
            	   Pedido pedido = null;
            	   String azul = "FF99CCFF";
            	   String verde = "FFCCFFCC";
            	   
            	   List<Pedido> listaPedidosDeLinhaAzul = new ArrayList<Pedido>();
            	   List<Pedido> listaPedidosDeLinhaVerde = new ArrayList<Pedido>();
            	   
            	   while (rowIterator.hasNext()) {
            		   
            		   Row row = rowIterator.next();
            		   
            		   if (row.getRowNum() == 0 || row.getRowNum() == 1 || row.getRowNum() == 2) {
            			   continue;
            		   }
            		   
            		   Iterator<Cell> cellIterator = row.cellIterator();
        			   pedido = new Pedido();
        			   boolean isAzul = false;
        			   boolean isVerde = false;
        			   
        			   while (cellIterator.hasNext()) {
        				   Cell cell = cellIterator.next();
        				   
        				   XSSFCellStyle cellStyle = (XSSFCellStyle) cell.getCellStyle();
        				   XSSFColor color = cellStyle.getFillForegroundColorColor();
        				   
        				   // Na planilha baixada do Adquira, um pedido � composto por informa��es que est�o na linha azul e na linha verde.
        				   // Separo essas informa��es em duas listas para depois criar uma lista �nica

        				   if (cell.getColumnIndex() == 0 && azul.equals(((XSSFColor)color).getARGBHex()) || cell.getColumnIndex() == 1 && azul.equals(((XSSFColor)color).getARGBHex()) ||
        					   cell.getColumnIndex() == 2 && azul.equals(((XSSFColor)color).getARGBHex()) || cell.getColumnIndex() == 4 && azul.equals(((XSSFColor)color).getARGBHex()) ||
        					   cell.getColumnIndex() == 5 && azul.equals(((XSSFColor)color).getARGBHex()) || cell.getColumnIndex() == 6 && azul.equals(((XSSFColor)color).getARGBHex()) ||
        					   cell.getColumnIndex() == 11 && azul.equals(((XSSFColor)color).getARGBHex())) {
        					   
        					   isAzul = true;
        					   
        					   switch (cell.getColumnIndex()) {
        					   case 0:
        						   pedido.setNumero(cell.getStringCellValue());
        						   break;
        					   case 1:
        						   pedido.setComprador(cell.getStringCellValue());
        						   break;
        					   case 2:
        						   pedido.setCnpjCliente(cell.getStringCellValue());
        						   break;       
        					   case 4:
        						   pedido.setData(formatarDataPedido(cell.getStringCellValue()));
        						   break;
        					   case 5:
        						   pedido.setValor(formatarValorPedido(cell.getStringCellValue()));
        						   break;
        					   case 6:
        						   pedido.setEstado(cell.getStringCellValue());
        						   break;
        					   case 11:
        						   pedido.setPrazoPagamento(cell.getStringCellValue());
        						   break;	   
        					   }
        					   
        				   } else if (cell.getColumnIndex() == 0 && verde.equals(((XSSFColor)color).getARGBHex()) || cell.getColumnIndex() == 1 && verde.equals(((XSSFColor)color).getARGBHex()) ||
        						      cell.getColumnIndex() == 2 && verde.equals(((XSSFColor)color).getARGBHex())) {
        					   
        					   isVerde = true;
        					   
        					   switch (cell.getColumnIndex()) {
        					   case 0:
        						   pedido.setNumero(cell.getStringCellValue());
        						   break;
        					   case 1:
        						   pedido.setItem(Integer.valueOf(cell.getStringCellValue()));
        						   break;  
        					   case 2:
        						   pedido.setCap(cell.getStringCellValue());
        						   break;       
        						   
        					   }
        					   
        				   }
        				   
        			   }
        			   
		        	   // Data Extra��o
		        	    pedido.setDataExtracao(dataAtualPlanilhaFinal);
            		   
        			   if (isAzul) {
        				   
        				   listaPedidosDeLinhaAzul.add(pedido);
        			   
        			   } else if (isVerde) {
        				   
        				   listaPedidosDeLinhaVerde.add(pedido);
        			   }
            		   
            	   }
            	   arquivo.close();
            	   
            	   // Unifica os pedidos de linha azul com os pedidos de linha verde na lista listaPedidos
            	   unificaPedidosdeLinhaAzulComPedidosDeLinhaVerde(listaPedidosDeLinhaAzul, listaPedidosDeLinhaVerde, listaPedidos);
               }

		} catch (Exception e) {
     	   contadorErroslerRelatorioExcel ++;
            Thread.sleep(3000);
            System.out.println("Arquivo Excel n�o encontrado! Tentando resolver, tentativa de n�mero: " + contadorErroslerRelatorioExcel);
            
            // Tento ler o arquivo por at� 20 vezes
            if (contadorErroslerRelatorioExcel <= 20) {
            	
				System.out.println("Deu erro no m�todo lerRelatorioExcel, tentativa de acerto: " + contadorErroslerRelatorioExcel);
				// Est� dando erro de logout no servidor
				// O bot�o de logout est� ficando escondido
				// ent�o retirarei o logout e o login por enquanto
				fazerLogoutAdquira(driver, wait);
				fazerLoginAdquira(driver, wait, js);
				//acessarPaginaInicial(driver, wait);
				fazerDownlodRelatorioPorPeriodo(driver, wait, js);
				moverArquivosEntreDiretorios(driver, wait, js, Util.getValor("caminho.download.relatorios") + "\\" + nomeRelatorioBaixado, subdiretorioRelatoriosBaixados);
         	    lerRelatorioExcel(driver, wait, js, relatorio, subdiretorioRelatoriosBaixados);
            
            } else {
         	   throw new Exception("Arquivo Excel n�o encontrado! : " + e);
            }
		}
        
        if (listaPedidos.size() == 0) {
        	   throw new Exception("Lista de pedidos est� vazia");
        }
        
  }


    
	   public static void mataProcessosFirefox() throws IOException, SQLException, InterruptedException{
		   
		   // Mata o Firefox
		   Runtime.getRuntime().exec(Util.getValor("caminho.matar.firefox"));
		   Thread.sleep(3000);
	   
	   }

	   public static void mataProcessosGoogle() throws IOException, SQLException, InterruptedException{
			  
		   // Mata o Google
		   // Viegas
		   Runtime.getRuntime().exec(Util.getValor("caminho.matar.google"));
		   Thread.sleep(3000);
		   
		   // Mata o chromedriver
		   Runtime.getRuntime().exec(Util.getValor("caminho.matar.chromedriver"));
		   Thread.sleep(3000);
	   
	   }

    
    // Propiedades do driver para abrir no IE, Chrome ou Firefox
    public static WebDriver getWebDriver() throws InterruptedException {
    	
    	WebDriver driver = null;
    		
            try {
				
            	if ("Chrome".equals(Util.getValor("navegador"))) {
				    
					File file = new File(Util.getValor("driver.Chrome.selenium"));
					System.setProperty(Util.getValor("propriedade.sistema.para.driver.Chrome.selenium"), file.getAbsolutePath());
				    DesiredCapabilities caps = DesiredCapabilities.chrome();
				    caps.setJavascriptEnabled(true);
				    caps.setCapability("ignoreZoomSetting", true);
				    caps.setCapability("nativeEvents",false);
				    ChromeOptions chromeOptions = new ChromeOptions(); 
				    Map<String, Object> chromePreferences = new HashMap<String, Object>();
					chromePreferences.put("profile.default_content_settings.popups", 0);
				    chromePreferences.put("download.default_directory",Util.getValor("caminho.download.relatorios"));
				    chromePreferences.put("browser.helperApps.neverAsk.saveToDisk", "text/plain, application/vnd.ms-excel, text/csv, text/comma-separated-values, application/octet-stream");
				    chromeOptions.setExperimentalOption("prefs", chromePreferences);
				    
				    // Argumento que faz com que o navegador use os dados do usu�rio salvos
				    // Com isso n�o ser� necess�rio digitar os dados de login no sharepoint, pois ele pegar� as informa��es do usu�rio salvas na m�quina
				    // Um ponto importante � que n�o poderemos ter mais de uma sess�o do Chrome aberta
				    // Outro ponto importante � que a op��o acima browser.helperApps.neverAsk.saveToDisk que permite que o browser salve um arquivo sem perguntar aonde salvar,
				    // n�o funcionar� por conta do trecho abaixo.
				    // Neste caso deveremos setar manualmente essa op��o no Chrome antes de rodar o rob�
				    // Ser� necess�rio fazer aparecer essa pasta no explorer do usu�rio
				    chromeOptions.addArguments("user-data-dir=" + Util.getValor("caminho.dados.usuario.Chrome"));
				    chromeOptions.addArguments("--lang=pt");

				    // Com essa op��o, o Selenium executa tudo sem mostrar o navegador
				    // Por�m no Adquira n�o funciona
				    //chromeOptions.addArguments("--headless");
				    
				    driver = new ChromeDriver(chromeOptions);
				    // Limpa o cache usando m�todo do driver
				    driver.manage().deleteAllCookies();
				    
				    
				} else if ("internetExplorer".equals(Util.getValor("navegador"))) {
				
					File file = new File(Util.getValor("driver.internetExplorer.selenium"));
					System.setProperty(Util.getValor("propriedade.sistema.para.driver.internetExplorer.selenium"), file.getAbsolutePath());
				    DesiredCapabilities caps = DesiredCapabilities.internetExplorer();
				    caps.setJavascriptEnabled(true);
				    //caps.setPlatform(org.openqa.selenium.Platform.WINDOWS);
				    caps.setCapability("ignoreZoomSetting", true);
				    caps.setCapability("nativeEvents",false);
					InternetExplorerOptions ieOptions = new InternetExplorerOptions();
					ieOptions.setCapability("ignoreZoomSetting", true);
					ieOptions.setCapability("nativeEvents",false);
					ieOptions.setCapability("browser.download.folderList", 2);
					ieOptions.setCapability("browser.helperApps.alwaysAsk.force", false);
					ieOptions.setCapability("browser.download.manager.showWhenStarting",false);
					//ieOptions.setCapability("browser.download.dir",getValor("caminho.download.relatorios"));
					//ieOptions.setCapability("browser.helperApps.neverAsk.saveToDisk", "text/plain, application/vnd.ms-excel, text/csv, text/comma-separated-values, application/octet-stream");
					//ieOptions.setCapability("browser.helperApps.alwaysAsk.force", true);
					// Limpando o cache com propriedades do Internet Explorer
					//caps.setCapability(InternetExplorerDriver.IE_ENSURE_CLEAN_SESSION,true);
					//driver = new InternetExplorerDriver(caps);
					//driver.manage().timeouts().implicitlyWait(0, TimeUnit.SECONDS);
					
					driver = new InternetExplorerDriver(ieOptions);
				    // Limpa o cache usando m�todo do driver
				    driver.manage().deleteAllCookies();
				
				} else if ("Firefox".equals(Util.getValor("navegador"))) {
					
					File file = new File(Util.getValor("driver.Firefox.selenium"));
					System.setProperty(Util.getValor("propriedade.binario.Firefox.selenium"),Util.getValor("binario.Firefox")); 
					System.setProperty(Util.getValor("propriedade.sistema.para.driver.Firefox.selenium"),file.getAbsolutePath());
     				File profileDirectory = new File(Util.getValor("caminho.dados.usuario.Firefox"));
				    //FirefoxProfile fxProfile = new FirefoxProfile(profileDirectory);
     				// Viegas
     				//Na minha máquina só funciona assim
     				FirefoxProfile fxProfile = new FirefoxProfile();
				    fxProfile.setPreference("browser.download.folderList",2);
				    fxProfile.setPreference("browser.download.manager.showWhenStarting",false);
				    fxProfile.setPreference("browser.download.dir",Util.getValor("caminho.download.relatorios"));
				    fxProfile.setPreference("browser.helperApps.neverAsk.saveToDisk", "text/plain, application/vnd.ms-excel, application/zip, text/csv, text/comma-separated-values, application/octet-stream");
				    // Limpando o cache com propriedades do Firefox 
				    /*
				    fxProfile.setPreference("browser.cache.disk.enable", false);
				    fxProfile.setPreference("browser.cache.memory.enable", false);
				    fxProfile.setPreference("browser.cache.offline.enable", false);
				    fxProfile.setPreference("network.http.use-cache", false);
				    fxProfile.setPreference("network.cookie.cookieBehavior", 2);
				    */
				    
				    FirefoxOptions fxOptions = new FirefoxOptions();
				    fxOptions.setProfile(fxProfile);
				    driver = new FirefoxDriver(fxOptions);
				    // Limpa o cache usando m�]etodo do driver
				    driver.manage().deleteAllCookies();
				}
			
            } catch (IOException e) {
				System.out.println("Ocorreu um erro no metodo getWebDriver: " + e.getMessage());
			}
            
            return driver;
    } 


	
	
	   public static void inserirStatusExecucaoNoBanco(String servico, String dataHora, String status) throws IOException, SQLException{
			  
		   HistoricoExecucaoDao historicoExecucaoDao = new HistoricoExecucaoDao();
		   historicoExecucaoDao.inserirStatusExecucao(servico, dataHora, status);
	   
	   }



}

package cdp;

import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
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
import org.openqa.selenium.support.ui.WebDriverWait;

import bean.RelatorioControlsDueThisMonth;
import bean.RelatorioOperationalRiskIndexByClient;
import cdp.dao.ClientDataProtectionDao;
import cdp.dao.HistoricoExecucaoDao;
import cdp.util.Util;

public class AutomacaoClientDataProtection {
	
	private static String dataAtual = new SimpleDateFormat("yyyy_MM_dd HH_mm_ss").format(new Date());
	private static String dataAtualPlanilhaFinal = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss").format(new Date());
	private static String diretorioLogs = "";
	private static List<RelatorioControlsDueThisMonth> listaRelatorioControlsDueThisMonth = new ArrayList<RelatorioControlsDueThisMonth>();
	private static List<RelatorioOperationalRiskIndexByClient> listaRelatorioOperationalRiskIndexByClient = new ArrayList<RelatorioOperationalRiskIndexByClient>();
	
	private static int contadorErroslerRelatorioExcel = 0;
	private static int contadorExecutaAutomacaoClientDataProtection = 0;
	private static String subdiretorioRelatoriosBaixados = null;
	private static String subdiretorioRelatoriosBaixados2 = null;
	
	@SuppressWarnings("unused")
	public static void main(String[] args) throws Exception {
		
		WebDriver driver = null;

    		try {
    			diretorioLogs = Util.getValor("caminho.diretorio.relatorios") + "/" + dataAtual;
    			subdiretorioRelatoriosBaixados = Util.getValor("caminho.download.relatorios") + "\\" + dataAtual;
    			subdiretorioRelatoriosBaixados2 = diretorioLogs;
    			Util.criaDiretorio(subdiretorioRelatoriosBaixados);
    			
    			// As vezes o diretorio que armazena dados temporarios do Chome simplesmente some, dai o Selenium da pau na hora de chamar o browser
    			// Com o metodo abaixo, crio essa pasta se ela nao existir
    			Util.criaDiretorioTemp();

    			// Deleta os diretorios que possuirem data de criacao anterior a data de 7 dias atras
    			Util.apagaDiretoriosDeRelatorios(Util.getValor("caminho.download.relatorios"));
    			
    			executaAutomacaoClientDataProtection(driver);
            
    		} catch (Exception e) {
    			Util.gravarArquivo(diretorioLogs, "Erro ClientDataProtection" + " " + dataAtual, ".txt", e.getMessage(), "Ocorreu um erro na automacao de extracao de relatorios: ");
    			inserirStatusExecucaoNoBanco("ClientDataProtection", dataAtualPlanilhaFinal, "Erro");
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
    		WebDriverWait wait = new WebDriverWait(driver, 100);
    		
    		// Acessa o site Client Data Protection
    		acessarClientDataProtection(driver, wait, js);
    		
    		// Clicando no link TELEFONICA GROUP
    		clicarLinkTelefonicaGroup(driver, wait, js);
    		
    		// Faz o download do relatório Controls Due This Month
    		String nomeRelatorioControlsDueThisMonth = "ControlsDueThisMonth.xlsx";
    		fazerDownloadRelatorioControlsDueThisMonth(driver, wait, js, nomeRelatorioControlsDueThisMonth);
    		
    		// Acessa o site Client Data Protection na aba original
    		//((JavascriptExecutor) driver).executeScript("window.open()");
    		List<String> windowHandles = new ArrayList<>(driver.getWindowHandles());
    		driver.switchTo().window(windowHandles.get(0));
    		acessarClientDataProtection(driver, wait, js);
    		
    		// Faz o download do relatório Operational Risk Index By Client
    		String nomeRelatorioOperationalRiskIndexByClient = "Operational Risk Index By Client.xlsx";
    		fazerDownloadRelatorioOperationalRiskIndexByClient(driver, wait, js,  nomeRelatorioOperationalRiskIndexByClient);
    		
    		// Lê o relatório Controls Due This Month
    		lerRelatorioControlsDueThisMonth(subdiretorioRelatoriosBaixados2 + "/" + nomeRelatorioControlsDueThisMonth, subdiretorioRelatoriosBaixados);
    		
    		// Deleto o que estiver na tabela CDP_Controls_Due e gravo o relatório Controls Due This Month na mesma tabela
    		deletaEInsereRelatorioControlsDueThisMonthNoBanco();
    		
    		// Lê o relatório Operational Risk Index By Client
    		lerRelatorioOperationalRiskIndexByClient(subdiretorioRelatoriosBaixados2 + "/" + nomeRelatorioOperationalRiskIndexByClient, subdiretorioRelatoriosBaixados);
    		
    		// Deleto o que estiver na tabela CDP_Operational_Risc e gravo o relatório Operational Risk Index By Client na mesma tabela
    		deletaEInsereRelatorioOperationalRiskIndexByClient();
    		
    		Util.gravarArquivo(diretorioLogs, "Sucesso ClientDataProtection" + " " + dataAtual, ".txt", "", mensagemResultadoClientDataProtection);
    		
    		// Grava na tabela Tb_Historico_Execucao_Robos o servico, data e hora e status da execucao
    		inserirStatusExecucaoNoBanco("ClientDataProtection", dataAtualPlanilhaFinal, "Sucesso");
    		
    		System.out.println("Fim: " + new SimpleDateFormat("dd_MM_yyyy HH_mm_ss").format(new Date()));
    		
    		if (driver != null) {
    			driver.quit();
    		}
    		
    		mataProcessosGoogle();
    		mataProcessosFirefox();
			
		} catch (Exception e) {
			contadorExecutaAutomacaoClientDataProtection ++;
			// Executo ate 5 vezes se der erro no executaAutomacaoAdquiraSharepoint
			if (contadorExecutaAutomacaoClientDataProtection <= 3) {
				
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
    
    public static void fazerDownloadRelatorioControlsDueThisMonth(WebDriver driver, WebDriverWait wait, JavascriptExecutor js, String nomeRelatorioControlsDueThisMonth) throws Exception {
    	
    	// Clicando no link Show Summary Statistics & CDP Operational Compliance
    	String idOlhinho = "showSummaryTableText";
    	wait.until(ExpectedConditions.elementToBeClickable(By.id(idOlhinho))).click();
    	Thread.sleep(5000);
    	
    	// Clicando no link Controls Due this Month:
    	String textoControlsDueThisMonth = "Controls Due this Month:";
    	wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span [text()='"+textoControlsDueThisMonth+"']"))).click();
    	
    	// Digitando o usuário e senha com o Robot
    	digitarUsuarioSenha();
    	
    	// Não consegui interagir com nenhum elemento da página do relatório usando o Selenium
    	// Tive que apelar para o Robot ;)
    	clickBotaoSalvarRelatorio();
    	
    	//Move o excel baixado do diretorio relatorios para o diretorio correto
    	Util.moverArquivosEntreDiretorios(Util.getValor("caminho.download.relatorios") + "\\" + nomeRelatorioControlsDueThisMonth, subdiretorioRelatoriosBaixados);
    	Thread.sleep(1000);
    	
    }
    
    public static void fazerDownloadRelatorioOperationalRiskIndexByClient(WebDriver driver, WebDriverWait wait, JavascriptExecutor js, String nomeRelatorioOperationalRiskIndexByClient) throws Exception {
    	
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
    	
    	//Move o excel baixado do diretorio relatorios para o diretorio correto
    	Util.moverArquivosEntreDiretorios(Util.getValor("caminho.download.relatorios") + "\\" + nomeRelatorioOperationalRiskIndexByClient, subdiretorioRelatoriosBaixados);
    	Thread.sleep(1000);
    	
    }
    
    public static void digitarUsuarioSenha() throws Exception {
    	
    	Robot robo = new Robot();
    	RoboCapturaTeclas rob = new RoboCapturaTeclas(robo);
    	if ("S".equals(Util.getValor("digita.automaticamente.usuario"))) {
    		Thread.sleep(10000);
    		//robo.delay(1500);
    		
    		// Digita o e-mail
    		rob.datilografar(Util.getValor("username"));
    		Thread.sleep(1000);
    	
    	}

    	if ("S".equals(Util.getValor("digita.automaticamente.senha"))) {
    		Thread.sleep(3000);
    		// Apertando o tab 1 vez para chegar na senha
    		robo.keyPress(KeyEvent.VK_TAB);
    		robo.keyRelease(KeyEvent.VK_TAB);
    		//robo.delay(1500);
    		Thread.sleep(1000);
    		
    		// Digita a senha
    		rob.datilografar(Util.getValor("senha"));
    		Thread.sleep(1000);
    		
    		// Apertando o tab 1 vez para chegar no botão de fazer login
    		robo.keyPress(KeyEvent.VK_TAB);
    		robo.keyRelease(KeyEvent.VK_TAB);
    		Thread.sleep(1000);
    		
    		// Aperta o ENTER no botão de fazer login
    		robo.keyPress(KeyEvent.VK_ENTER);
    		robo.keyRelease(KeyEvent.VK_ENTER);
    		
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
    	
    	Thread.sleep(40000);
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
    	Thread.sleep(40000);
    	
        /*
    	Robot bot = new Robot();
        bot.mouseMove(posicaoX, posicaoY);    
        bot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
        bot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        */
    }
    
    public static void deletaEInsereRelatorioControlsDueThisMonthNoBanco() throws Exception {
    	
    	if (listaRelatorioControlsDueThisMonth != null && !listaRelatorioControlsDueThisMonth.isEmpty()) {
    		
    		int contador = 0;
    		
    		ClientDataProtectionDao deletarClientDataProtectionDao = new ClientDataProtectionDao();
    		// Deleto o que estiver na tabela CDP_Controls_Due
    		deletarClientDataProtectionDao.deletarRelatorioControlsDueThisMonth();
    		System.out.println("Relatorio Controls Due This Month. Apaguei todos os dados do banco");

    		for (RelatorioControlsDueThisMonth relatorioControlsDueThisMonth : listaRelatorioControlsDueThisMonth) {
    			ClientDataProtectionDao inserirClientDataProtectionDao = new ClientDataProtectionDao();
    			inserirClientDataProtectionDao.inserirRelatorioControlsDueThisMonth(relatorioControlsDueThisMonth);
    			contador++;
    			System.out.println("Relatorio Controls Due This Month. Inseri o registro de numero: " + contador + " de um total de " + listaRelatorioControlsDueThisMonth.size());
    			
    		}
    		
    	} else {
    		
    		System.out.println("O relatorio Controls Due This Month está vazio");
    		
    	}
    
    }
    
    public static void deletaEInsereRelatorioOperationalRiskIndexByClient() throws Exception {
    	
    	if (listaRelatorioOperationalRiskIndexByClient != null && !listaRelatorioOperationalRiskIndexByClient.isEmpty()) {
    		
    		int contador = 0;
    		
    		ClientDataProtectionDao deletarClientDataProtectionDao = new ClientDataProtectionDao();
    		// Deleto o que estiver na tabela CDP_Operational_Risc
    		deletarClientDataProtectionDao.deletarRelatorioOperationalRiskIndexByClient();
    		System.out.println("Relatorio Operational Risk Index By Client. Apaguei todos os dados do banco");

    		for (RelatorioOperationalRiskIndexByClient relatorioOperationalRiskIndexByClient : listaRelatorioOperationalRiskIndexByClient) {
    			ClientDataProtectionDao inserirClientDataProtectionDao = new ClientDataProtectionDao();
    			inserirClientDataProtectionDao.inserirRelatorioOperationalRiskIndexByClient(relatorioOperationalRiskIndexByClient);
    			contador++;
    			System.out.println("Relatorio Operational Risk Index By Client. Inseri o registro de numero: " + contador + " de um total de " + listaRelatorioOperationalRiskIndexByClient.size());
    			
    		}
    		
    	} else {
    		
    		System.out.println("O relatorio Operational Risk Index By Client está vazio");
    		
    	}
    
    }

    @SuppressWarnings("resource")
	public static void lerRelatorioControlsDueThisMonth(String relatorio, String subdiretorioRelatoriosBaixados) throws Exception {

        try {
               FileInputStream arquivo = new FileInputStream(new File(
            		   relatorio));
               
               File arquivoExcel = new File(relatorio);
               
               if (arquivoExcel.exists() && arquivoExcel.isFile() && arquivoExcel.length() > 0) {
            	   
            	   OPCPackage pkg = OPCPackage.open(new File(relatorio));
            	   
            	   XSSFWorkbook workbook = new XSSFWorkbook(pkg);
            	   
            	   XSSFSheet sheetPedidos =  workbook.getSheetAt(0);
            	   
            	   Iterator<Row> rowIterator = sheetPedidos.iterator();
            	   
            	   RelatorioControlsDueThisMonth relatorioControlsDueThisMonth = null;
            	   
            	   while (rowIterator.hasNext()) {
            		   
            		   Row row = rowIterator.next();
            		   
            		   if (row.getRowNum() == 0 || row.getRowNum() == 1) {
            			   continue;
            		   }
            		   
            		   Iterator<Cell> cellIterator = row.cellIterator();
        			   relatorioControlsDueThisMonth = new RelatorioControlsDueThisMonth();
        			   
        			   while (cellIterator.hasNext()) {
        				   Cell cell = cellIterator.next();
        				   
        				   switch (cell.getColumnIndex()) {
        				   case 0:
        					   relatorioControlsDueThisMonth.setClientName(cell.getStringCellValue());
        					   break;
        				   case 1:
        					   relatorioControlsDueThisMonth.setControlName(cell.getStringCellValue());
        					   break;
        				   case 2:
        					   relatorioControlsDueThisMonth.setControlID(String.valueOf(cell.getNumericCellValue()));
        					   break;
        				   case 3:
        					   relatorioControlsDueThisMonth.setControlOwner(cell.getStringCellValue());
        					   break;       
        				   case 4:
        					   relatorioControlsDueThisMonth.setIndicator(cell.getStringCellValue());
        					   break;
        				   case 5:
        					   relatorioControlsDueThisMonth.setNextDueDate(cell.getStringCellValue());
        					   break;
        				   case 6:
        					   relatorioControlsDueThisMonth.setVerified(cell.getStringCellValue());
        					   break;
        				   case 7:
        					   relatorioControlsDueThisMonth.setDeliveryLocationSource(cell.getStringCellValue());
        					   break;
        				   case 8:
        					   relatorioControlsDueThisMonth.setServiceName(cell.getStringCellValue());
        					   break;
           				   case 9:
        					   relatorioControlsDueThisMonth.setClientDataProtectionRequirement(cell.getStringCellValue());
        					   break;
        				   case 10:
        					   relatorioControlsDueThisMonth.setRelatedContractualRequirement(cell.getStringCellValue());
        					   break;
        				   case 11:
        					   relatorioControlsDueThisMonth.setRegulatoryRequirement(cell.getStringCellValue());
        					   break;       
        				   case 12:
        					   //relatorioControlsDueThisMonth.setEngagementLevelProcedure(cell.getStringCellValue());
        					   break;
        				   case 13:
        					   relatorioControlsDueThisMonth.setEngagementLevelProcedure(cell.getStringCellValue());
        					   break;
        				   case 14:
        					   relatorioControlsDueThisMonth.setCompliance(cell.getStringCellValue());
        					   break;
        				   case 15:
        					   relatorioControlsDueThisMonth.setOngoingFrequency(cell.getStringCellValue());
        					   break;
        				   }

        			   }
        			   
		        	   // Data Extracao
		        	    relatorioControlsDueThisMonth.setDataExtracao(dataAtualPlanilhaFinal);
		        	    
		        	    listaRelatorioControlsDueThisMonth.add(relatorioControlsDueThisMonth);
            		   
            	   }
            	   arquivo.close();
            	   
               }

		} catch (Exception e) {
     	   contadorErroslerRelatorioExcel ++;
            Thread.sleep(3000);
            System.out.println("Arquivo Excel nao encontrado! Tentando resolver, tentativa de numero: " + contadorErroslerRelatorioExcel);
            
            // Tento ler o arquivo por ate 5 vezes
            if (contadorErroslerRelatorioExcel <= 3) {
            	
				System.out.println("Deu erro no metodo lerRelatorioControlsDueThisMonth, tentativa de acerto: " + contadorErroslerRelatorioExcel);
         	    lerRelatorioControlsDueThisMonth(relatorio, subdiretorioRelatoriosBaixados);
            
            } else {
         	   throw new Exception("Arquivo Excel nao encontrado! : " + e);
            }
		}
        
  }
    
    @SuppressWarnings("resource")
	public static void lerRelatorioOperationalRiskIndexByClient(String relatorio, String subdiretorioRelatoriosBaixados) throws Exception {

        try {
               FileInputStream arquivo = new FileInputStream(new File(
            		   relatorio));
               
               File arquivoExcel = new File(relatorio);
               
               if (arquivoExcel.exists() && arquivoExcel.isFile() && arquivoExcel.length() > 0) {
            	   
            	   OPCPackage pkg = OPCPackage.open(new File(relatorio));
            	   
            	   XSSFWorkbook workbook = new XSSFWorkbook(pkg);
            	   
            	   XSSFSheet sheetPedidos =  workbook.getSheetAt(0);
            	   
            	   Iterator<Row> rowIterator = sheetPedidos.iterator();
            	   
            	   RelatorioOperationalRiskIndexByClient relatorioOperationalRiskIndexByClient = null;
            	   
            	   while (rowIterator.hasNext()) {
            		   
            		   Row row = rowIterator.next();
            		   
            		   if (row.getRowNum() == 0 || row.getRowNum() == 1 || row.getRowNum() == 2) {
            			   continue;
            		   }
            		   
            		   Iterator<Cell> cellIterator = row.cellIterator();
        			   relatorioOperationalRiskIndexByClient = new RelatorioOperationalRiskIndexByClient();
        			   
        			   while (cellIterator.hasNext()) {
        				   Cell cell = cellIterator.next();
        				   
        				   switch (cell.getColumnIndex()) {
        				   case 0:
        					   relatorioOperationalRiskIndexByClient.setCDPTrackerID(String.valueOf(cell.getNumericCellValue()));
        					   break;
        				   case 1:
        					   relatorioOperationalRiskIndexByClient.setClientName(cell.getStringCellValue());
        					   break;
        				   case 2:
        					   relatorioOperationalRiskIndexByClient.setTier(cell.getStringCellValue());
        					   break;
        				   case 3:
        					   relatorioOperationalRiskIndexByClient.setMasterCustomerNbr(cell.getStringCellValue());
        					   break;       
        				   case 4:
        					   relatorioOperationalRiskIndexByClient.setMarket(cell.getStringCellValue());
        					   break;
        				   case 5:
        					   relatorioOperationalRiskIndexByClient.setMarketUnit(cell.getStringCellValue());
        					   break;
        				   case 6:
        					   relatorioOperationalRiskIndexByClient.setOwningOrganization(cell.getStringCellValue());
        					   break;
        				   case 7:
        					   relatorioOperationalRiskIndexByClient.setAccountableMD(cell.getStringCellValue());
        					   break;
        				   case 8:
        					   relatorioOperationalRiskIndexByClient.setAccountISL(cell.getStringCellValue());
        					   break;
           				   case 9:
        					   relatorioOperationalRiskIndexByClient.setAccountContractManager(cell.getStringCellValue());
        					   break;
        				   case 10:
        					   relatorioOperationalRiskIndexByClient.setCDPAccountManager(cell.getStringCellValue());
        					   break;
        				   case 11:
        					   relatorioOperationalRiskIndexByClient.setWeightedOperationalRiskCurrent(String.valueOf(cell.getNumericCellValue()));
        					   break;       
        				   case 12:
        					   relatorioOperationalRiskIndexByClient.setWeightedOperationalRiskCurrentColor(cell.getStringCellValue());
        					   break;
        				   case 13:
        					   relatorioOperationalRiskIndexByClient.setWeightedOperationalRiskMonthEnd(String.valueOf(cell.getNumericCellValue()));
        					   break;
        				   case 14:
        					   relatorioOperationalRiskIndexByClient.setWeightedOperationalRiskMonthEndColor(cell.getStringCellValue());
        					   break;
        				   case 15:
        					   relatorioOperationalRiskIndexByClient.setControls30DaysPastDue(String.valueOf(cell.getNumericCellValue()));
        					   break;
        				   case 16:
        					   relatorioOperationalRiskIndexByClient.setControls60DaysPastDue(String.valueOf(cell.getNumericCellValue()));
        					   break;
        				   case 17:
        					   relatorioOperationalRiskIndexByClient.setNonCompliantControls(String.valueOf(cell.getNumericCellValue()));
        					   break;       
        				   case 18:
        					   relatorioOperationalRiskIndexByClient.setPastDueControls(String.valueOf(cell.getNumericCellValue()));
        					   break;
        				   case 19:
        					   relatorioOperationalRiskIndexByClient.setControlwHighestOperationalRisk(cell.getStringCellValue());
        					   break;
        				   case 20:
        					   relatorioOperationalRiskIndexByClient.setOperationalRiskScoreOfWorstControl(String.valueOf(cell.getNumericCellValue()));
        					   break;
        				   case 21:
        					   relatorioOperationalRiskIndexByClient.setIsImplemented(cell.getStringCellValue());
        					   break;
        					   
        				   }

        			   }
        			   
		        	   // Data Extracao
		        	    relatorioOperationalRiskIndexByClient.setDataExtracao(dataAtualPlanilhaFinal);
		        	    
		        	    listaRelatorioOperationalRiskIndexByClient.add(relatorioOperationalRiskIndexByClient);
            		   
            	   }
            	   arquivo.close();
            	   
               }

		} catch (Exception e) {
     	   contadorErroslerRelatorioExcel ++;
            Thread.sleep(3000);
            System.out.println("Arquivo Excel nao encontrado! Tentando resolver, tentativa de numero: " + contadorErroslerRelatorioExcel);
            
            // Tento ler o arquivo por ate 5 vezes
            if (contadorErroslerRelatorioExcel <= 3) {
            	
				System.out.println("Deu erro no metodo lerRelatorioOperationalRiskIndexByClient, tentativa de acerto: " + contadorErroslerRelatorioExcel);
				lerRelatorioOperationalRiskIndexByClient(relatorio, subdiretorioRelatoriosBaixados);
            
            } else {
         	   throw new Exception("Arquivo Excel nao encontrado! : " + e);
            }
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

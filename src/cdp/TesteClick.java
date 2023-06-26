package cdp;

import java.awt.Robot;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.interactions.PointerInput;
import org.openqa.selenium.interactions.Sequence;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.WebDriverWait;

import bean.RelatorioControlsDueThisMonth;
import bean.RelatorioOperationalRiskIndexReport;
import cdp.dao.ClientDataProtectionDao;

public class TesteClick {
	
	private static List<RelatorioControlsDueThisMonth> listaRelatorioControlsDueThisMonth;
	private static List<RelatorioOperationalRiskIndexReport> listaRelatorioOperationalRiskIndexByClient;

	public static void main(String[] args) throws Exception {
		
   		listaRelatorioControlsDueThisMonth = new ArrayList<RelatorioControlsDueThisMonth>();
		listaRelatorioOperationalRiskIndexByClient = new ArrayList<RelatorioOperationalRiskIndexReport>();
		
		ClientDataProtectionDao testeClientDataProtectionDao = new ClientDataProtectionDao();
		
		testeClientDataProtectionDao.inserirLeftJoin_ControlDue_OperacionalRisc();
		
		StringBuilder s =(new StringBuilder()).append("Sachin").append("Tendulkar");  

		System.out.println(s);
		
		AutomacaoClientDataProtection obj = new AutomacaoClientDataProtection();
		
		String subdiretorioRelatoriosBaixados2 = "C:/Viegas/desenvolvimento/Selenium/ClientDataProtection/relatorios/2022_11_07 15_30_09";
		
		String subdiretorioRelatoriosBaixados = "C:\\Viegas\\desenvolvimento\\Selenium\\ClientDataProtection\\relatorios\\2022_11_07 15_30_09";
		
		obj.lerRelatorioControlsDueThisMonth(subdiretorioRelatoriosBaixados2 + "/" + "ControlsDueThisMonth.xlsx", subdiretorioRelatoriosBaixados, listaRelatorioControlsDueThisMonth);
		
		obj.lerRelatorioOperationalRiskIndexReport(subdiretorioRelatoriosBaixados2 + "/" + "Operational Risk Index By Client.xlsx", subdiretorioRelatoriosBaixados, listaRelatorioOperationalRiskIndexByClient);
		
		Robot robo= new Robot();
		RoboCapturaTeclas rob= new RoboCapturaTeclas(robo);
		Runtime.getRuntime().exec("notepad" );
		robo.delay(1500);
		
		//rob.datilografar("O cachorro é português, é sim ão ão ão üüÜ, !@#$%¨&*()_+{}`^?;.,<>:ããõ");
		
		
	}

}

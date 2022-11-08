package cdp;

import java.awt.Robot;
import java.util.Collections;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.interactions.PointerInput;
import org.openqa.selenium.interactions.Sequence;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.WebDriverWait;

public class TesteClick {

	public static void main(String[] args) throws Exception {
		
		
		
		AutomacaoClientDataProtection obj = new AutomacaoClientDataProtection();
		
		String subdiretorioRelatoriosBaixados2 = "C:/Viegas/desenvolvimento/Selenium/ClientDataProtection/relatorios/2022_11_05 16_02_07";
		
		String subdiretorioRelatoriosBaixados = "C:\\Viegas\\desenvolvimento\\Selenium\\ClientDataProtection\\relatorios\\2022_11_05 16_02_07";
		
		//obj.lerRelatorioOperationalRiskIndexByClient(subdiretorioRelatoriosBaixados2 + "/" + "Operational Risk Index By Client.xlsx", subdiretorioRelatoriosBaixados);
		
		Robot robo= new Robot();
		RoboCapturaTeclas rob= new RoboCapturaTeclas(robo);
		Runtime.getRuntime().exec("notepad" );
		robo.delay(1500);
		
		rob.datilografar("O cachorro é português, é sim ão ão ão üüÜ, !@#$%¨&*()_+{}`^?;.,<>:ããõ");
		
		
	}

}

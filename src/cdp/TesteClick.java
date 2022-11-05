package cdp;

import java.util.Collections;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.interactions.PointerInput;
import org.openqa.selenium.interactions.Sequence;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.WebDriverWait;

public class TesteClick {

	public static void main(String[] args) {

		
        // Viegas
    	WebDriver driver = getWebDriver();
    	JavascriptExecutor js = (JavascriptExecutor) driver;
    	WebDriverWait wait = new WebDriverWait(driver, 20);
    	
    	//((JavascriptExecutor) getDriver()).executeScript("arguments[0].scrollIntoView();", webElement);
        
		// Abrindo a URl
		driver.manage().window().maximize();
		//driver.get(getValor("url.painel.rtc"));
		// Não funciona chamar a tela do painel direto, então tenho que chamar a tela do projeto modelo
		// e depois chamar a tela do painel
		driver.get(getValor("url.projeto.artefato.rtc") + idProjeto);
        
        // Preenchendo os dados de login
        WebElement username = driver.findElement(By.id("jazz_app_internal_LoginWidget_0_userId"));
        username.sendKeys(getValor("username"));
        
        // Preenchendo os dados da senha
        WebElement senha = driver.findElement(By.id("jazz_app_internal_LoginWidget_0_password"));
        senha.sendKeys(getValor("senha"));

        // Fazendo login
        WebElement botaoLogin = driver.findElement(By.xpath("//button[@type='submit']"));
        //js.executeScript("arguments[0].click()", botaoLogin);
        
        
        //Used points class to get x and y coordinates of element.
        WebElement botaoLoginTeste = driver.findElement(By.xpath("//button[@type='submit']"));
        Point classname = botaoLoginTeste.getLocation();
        int xcordi = classname.getX();
        //System.out.println("Element's Position from left side "+xcordi +" pixels.");
        int ycordi = classname.getY();
        //System.out.println("Element's Position from top "+ycordi +" pixels.");
        //clicking on the logo based on x coordinate and y coordinate 
        //you will be redirecting to the home page of softwaretestingmaterial.com
        //action.moveByOffset(xOffset, yOffset);
        
        
        Actions builder = new Actions(driver);
        Action moveM = builder.moveByOffset(xcordi, ycordi).build();
        moveM.perform();

        Action click = builder.click().build();
        click.perform();
        
        
        // Chamo a tela do painel depois de ter chamado a tela de login
        driver.get(getValor("url.painel.rtc"));
        
        // Espera para aparecer a página.
        //Para isso espero aparecer um elemento da home, que é um texto no cabeçalho: Operation Solution
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@id='jazz_ui_internal_TruncatedTextNode_0']/div")));
        
        String projetoTeste = "Teste Viegas";
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span [text()='"+projetoTeste+"']"))).click();
        
        Thread.sleep(5000);
        
        // Descer toda a tela
        js.executeScript("window.scrollBy(0,1000)");
        Thread.sleep(500); 
        
        
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("jazz_ui_MenuPopup_22")));
        //wait.until(ExpectedConditions.elementToBeClickable(By.id("jazz_ui_MenuPopup_19"))).click();
        
        WebElement abrirMenuReunioes = driver.findElement(By.id("jazz_ui_MenuPopup_22"));
        //((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", abrirMenuReunioes);
        //js.executeScript("window.scrollBy(0,1000)");
        Thread.sleep(500); 
        
        
        
        //System.out.println("getWidth "+abrirMenuReunioes.getSize().getWidth() +" pixels.");
        //System.out.println("getHeight "+abrirMenuReunioes.getSize().getHeight() +" pixels.");
        
        
        Point classnameAbrirMenuReunioes = abrirMenuReunioes.getLocation();
        int xAbrirMenuReunioes = classnameAbrirMenuReunioes.getX();
        System.out.println("Element's Position from left side "+xAbrirMenuReunioes +" pixels.");
        int yAbrirMenuReunioes = classnameAbrirMenuReunioes.getY();
        System.out.println("Element's Position from top side "+yAbrirMenuReunioes +" pixels.");
        
        Actions builder2 = new Actions(driver);
        //js.executeScript("window.scrollBy(0,1000)");
        //Thread.sleep(500); 

        // Projeto Modelo - Reuniões Status
        //Element's Position from left side 464 pixels.
        //Element's Position from top side 405 pixels.
    
        
        
        // Projeto Modelo - Áreas Envolvidas
        // Element's Position from left side 1454 pixels.
        // Element's Position from top side 405 pixels.
        
        
        
        // Projeto Modelo - Artefatos Status
        //Element's Position from left side 1454 pixels.
        //Element's Position from top side 484 pixels.
        //x=1454-667 = 787
        //y=484-49   = 435
        
        
        
        // Projeto Modelo - Projeto Modelo - Áreas Envolvidas
        // Action moveMyAbrirMenuReunioes = builder2.moveByOffset(667, -30).build();   
        
        // Projeto Modelo - Artefatos Status
        // Action moveMyAbrirMenuReunioes = builder2.moveByOffset(667, 49).build();
        
        // Projeto Modelo - Requisitos Solicitados - Indicador Tipo
        // Action moveMyAbrirMenuReunioes = builder2.moveByOffset(667, 170).build();
        
        // Projeto Modelo - Requisitos Aceitos - Indicador Tipo
        // Action moveMyAbrirMenuReunioes = builder2.moveByOffset(667, 235).build();
  
        Action moveMyAbrirMenuReunioes2 = builder2.moveByOffset(667, -30).click().build();
        moveMyAbrirMenuReunioes2.perform();
        Thread.sleep(500); 
        
        //Action moveMyAbrirMenuReunioes3 = builder2.moveByOffset(0, 0).click().build();
        //moveMyAbrirMenuReunioes3.perform();
        Thread.sleep(500);
        
        Action moveMyAbrirMenuReunioes = builder2.moveByOffset(667, 235).click().build();
        moveMyAbrirMenuReunioes.perform();

        //Action click2 = builder2.contextClick().build();
        //Action click2 = builder2.click().build();
        //click2.perform();
        
        
        WebElement element = driver.findElement(By.id("jazz_ui_MenuPopup_19"));
        //((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", element);
        //js.executeScript("window.scrollBy(0,1000)");
        Thread.sleep(500); 
        Actions act = new Actions(driver);
        act.moveByOffset(667, 49).click().build().perform();
        //act.moveToElement(element).perform();
        
       // act.moveByOffset(471, 365).perform();
        
       // Action click3 = act.click().build();
       // click3.perform();
        
        
        
        // Viegas
      
		
		
	}

}

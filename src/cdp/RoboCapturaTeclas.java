package cdp;

import java.awt.Dimension;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.event.KeyEvent;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

import javax.imageio.ImageIO;

/*
 * Peguei esse programa da internet: https://www.guj.com.br/t/robot-darth-vader-quero-dominar-o-mundo/131842
 * Está pegando todas as acentuções,
 *	Só está parado no Ç, ç , " , ’ e o ? está saindo como /
 * 
 */
   
public final class RoboCapturaTeclas {
	private Robot robot;
	private BufferedImage imagemArquivo, imagemPedaco,imagemInteira;
   
    public RoboCapturaTeclas(Robot robot) throws NullPointerException {  
         if (robot == null) {  
             throw new NullPointerException();  
         }  
         this.robot = robot;  
     }     
     @SuppressWarnings("deprecation")
	public void clique(int x, int y) {
         robot.mouseMove(x, y);  
         robot.mousePress(KeyEvent.BUTTON1_MASK);  
         robot.mouseRelease(KeyEvent.BUTTON1_MASK);           
     } 
    
    
     @SuppressWarnings("deprecation")
	public void cliqueDireito(int x, int y) {  
         robot.mouseMove(x, y);  
         robot.mousePress(KeyEvent.BUTTON2_MASK);
         robot.mouseRelease(KeyEvent.BUTTON2_MASK);         
     } 
     
        
     public void datilografar(int key) {  
         robot.keyPress(key);  
         robot.keyRelease(key);  
     }   
     public void datilografarSequencia(int... keys) {  
         for (int i = 0; i < keys.length; i++) {  
             robot.keyPress(keys[i]);  
         }  
         for (int i = keys.length - 1; i >= 0; i--) {  
             robot.keyRelease(keys[i]);  
         }  
     } 
     public void rolarMouse(int rotacao){
    	 robot.mouseWheel(rotacao);
     }
    public void datilografar(String string) {
    	char acento=' ';
         for (char c : string.toCharArray()) {
        	 acento=precisaAcentuar(c); 
        	 if(acento!=' '){
        		 escreve(acento);
        		 escreve(c);
        	 }else{
        		 escreve(c);
        	 }
             
         }  
     }  
     private void escreve(char c){
    	 if (Character.isUpperCase(c) || precisaPrecionarShift(c)) {  
             robot.keyPress(KeyEvent.VK_SHIFT);  
         }  
         Integer i = MapaCaracteres.KEYS.get(Character.toUpperCase(c));  
         if (i == null) {  
             i = KeyEvent.VK_NUMBER_SIGN;  
         }  
         datilografar(i);  
         if (Character.isUpperCase(c) || precisaPrecionarShift(c)) {  
             robot.keyRelease(KeyEvent.VK_SHIFT);  
         }  
    	 
     }
     private boolean precisaPrecionarShift(char c) {  
         return c == '!' || c == '@' || c == '#' || c == '$' || c == '%'|| c == '¨'  
                 || c == '&' || c == '*' || c == '(' || c == ')' || c == '`'|| c == '_'  
                 || c == '{'|| c == '+' || c == '}' || c == '^' || c == '?' || c == ':'  
                 || c == '>' || c == '<' || c == '|' || c == '\"' || c == '\'';  
     }
     private char precisaAcentuar(char c){
    	  if(c=='Á'||c=='É'||c=='Í'||c=='Ó'||c=='Ú'||
    			  c=='á'||c=='é'||c=='í'||c=='ó'||c=='ú'){
    		  return '´';
    	  }else if(c=='Â'||c=='Ê'||c=='Î'||c=='Ô'||c=='Û'||c=='â'||
    			  c=='ê'||c=='î'||c=='ô'||c=='û'){
    		  return '^';
    	  }else if(c=='Ã'||c=='Õ'||c=='ã'||c=='õ'){
    		  return '~';
    	  }else if(c=='À'||c=='È'||c=='Ì'||c=='Ò'||c=='Ù'||c=='à'||
    			  c=='è'||c=='ì'||c=='ò'||c=='ù'){
    		  return '`';
    	  }else if(c=='Ä'||c=='Ë'||c=='Ï'||c=='Ö'||c=='Ü'||c=='ä'||
    			  c=='ë'||c=='ï'||c=='ö'||c=='ü'){
    		  return '¨';
    	  }else{
    		  return ' ';
    	  }
     }
     private boolean compararImagens(BufferedImage image1, BufferedImage image2) {
 		if(image1.getWidth() != image2.getWidth() || image1.getHeight() 
 				!= image2.getHeight())
 			return false;

 		for(int x = 0; x < image1.getWidth(); x++) {
 			for(int y = 0; y < image1.getHeight(); y++) {
 				if(image1.getRGB(x, y) != image2.getRGB(x, y))
 					return false;
                 }
             }
             return true;
 	}
    public void capturarImagemPedaco(int inicioH,int inicioV,int altura,int largura){
    	imagemPedaco = robot.createScreenCapture(new Rectangle(inicioH, inicioV,altura,
    			largura));    	
    }
    public void capturarImagem(){
    	Toolkit toolkit = Toolkit.getDefaultToolkit();
    	final Dimension dimension = toolkit.getScreenSize();
		while (true) {
			// Pegando a imagem e gravando
			imagemInteira = robot.createScreenCapture(new Rectangle(0, 0,
					dimension.width, dimension.height));
			try {
				ImageIO.write(imagemInteira, "JPEG", new File("C:\\LogalSavamento.JPEG"));
			} catch (IOException e) {
				System.out.println("Não foi possivel Guardar a imagem");
				e.printStackTrace();
			}
		}   	
    	
    }
    public void carregarImagem(String arquivo){
    	try {
			imagemArquivo = ImageIO.read(new File(arquivo));
		} catch (IOException e) {
			System.out.println("Não foi possível carregar o arquivo");
			e.printStackTrace();
		}
    } 
    
    
   
 }  
package cdp;

import java.awt.event.KeyEvent;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

public final class MapaCaracteres {  
  
    public static final Map<Character, Integer> KEYS;  
  
    static {  
        Map<Character, Integer> map = new HashMap<Character, Integer>();  
        for (char c = 'A'; c <= 'Z'; c++) {  
            map.put(c, (int) c);  
        }  
        for (char c = '0'; c <= '9'; c++) {  
            map.put(c, (int) c);  
        }  
        for (char c = '\b'; c <= '\n'; c++) {  
            map.put(c, (int) c);  
        }
        for(char c='À';c<='Ä';c++){
        	map.put(c, 65);
        }
        for(char c='È';c<='Ë';c++){
        	map.put(c, 69);
        }
        for(char c='Ì';c<='Ï';c++){
        	map.put(c, 73);
        }
        for(char c='Ò';c<='Ö';c++){
        	map.put(c, 79);
        }
        for(char c='Ù';c<='Ü';c++){
        	map.put(c, 85);
        }
        
        map.put('\'', KeyEvent.VK_QUOTEDBL);  
        map.put('\"', KeyEvent.VK_QUOTEDBL); 
        map.put('!', KeyEvent.VK_1);  
        map.put('@', KeyEvent.VK_2); 
        map.put('#', KeyEvent.VK_3);
        map.put('$', KeyEvent.VK_4); 
        map.put('%', KeyEvent.VK_5);
        map.put('¨', KeyEvent.VK_6);
        map.put('&', KeyEvent.VK_7);
        map.put('*', KeyEvent.VK_8); 
        map.put('(', KeyEvent.VK_9); 
        map.put(')', KeyEvent.VK_0);
        map.put('-', KeyEvent.VK_MINUS);
        map.put('_', KeyEvent.VK_MINUS);
        map.put('+', KeyEvent.VK_EQUALS); 
        map.put('=', KeyEvent.VK_EQUALS);
        
        map.put('`', KeyEvent.VK_DEAD_ACUTE);
        map.put('´', KeyEvent.VK_DEAD_ACUTE);
        map.put('~', KeyEvent.VK_DEAD_TILDE);  
        map.put('^', KeyEvent.VK_DEAD_TILDE);
        map.put('[', KeyEvent.VK_OPEN_BRACKET);  
        map.put(']', KeyEvent.VK_CLOSE_BRACKET);  
        map.put('{', KeyEvent.VK_OPEN_BRACKET);  
        map.put('}', KeyEvent.VK_CLOSE_BRACKET);
        map.put('Ç', KeyEvent.VK_DEAD_CEDILLA);
        
        map.put('\\',KeyEvent.VK_BACK_SLASH);  
        map.put('|', KeyEvent.VK_BACK_SLASH); 
        map.put(',', KeyEvent.VK_COMMA); 
        map.put('<', KeyEvent.VK_COMMA); 
        map.put('.', KeyEvent.VK_PERIOD);
        map.put('>', KeyEvent.VK_PERIOD);
        map.put(';', KeyEvent.VK_SEMICOLON);  
        map.put(':', KeyEvent.VK_SEMICOLON); 
        map.put('/', KeyEvent.VK_DIVIDE);  
        map.put('?', KeyEvent.VK_DIVIDE);
        map.put(' ', KeyEvent.VK_SPACE); 
            
        KEYS = Collections.unmodifiableMap(map);  
    }
    private MapaCaracteres() {  
            // can not instantiate  
    } 
  }

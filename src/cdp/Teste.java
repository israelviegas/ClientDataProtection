package cdp;


import java.text.SimpleDateFormat;
import java.time.ZonedDateTime;
import java.util.Date;

public class Teste {

	public static void main(String[] args) {

		
        ZonedDateTime actualDate =  ZonedDateTime.now();

        actualDate = actualDate.minusDays(3);
        
        
		
        String filename = String.format("%s/reconciliation_[%s]/branch_[%s]_equipment.csv",
                new SimpleDateFormat("yyyyMMdd").format(new Date(0)),
                123,
                456);
        
        System.out.println(filename);
		
		
	}

}

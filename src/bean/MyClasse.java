package bean;

import java.lang.reflect.Field;
import java.lang.reflect.Method;

public class MyClasse {
    public static void main(String[] args) {
        Class<RelatorioControlsDueThisMonth> myClass = RelatorioControlsDueThisMonth.class;

        // Obter o nome dos campos
        Field[] fields = myClass.getDeclaredFields();
        for (Field field : fields) {
            System.out.println("Nome do campo: " + field.getName());
        }

        // Obter o nome dos métodos get
        Method[] methods = myClass.getDeclaredMethods();
        for (Method method : methods) {
            String methodName = method.getName();
            if (methodName.startsWith("get")) {
                System.out.println("Nome do método get: " + methodName);
            }
        }
    }
}

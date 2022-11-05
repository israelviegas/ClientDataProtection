package adquira;

import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

import adquira.util.ConnectionFactory;
import adquira.util.Util;

public class AdquiraDao {
	
	private Connection connection;
	
	 public AdquiraDao() throws IOException {
		 
		 String sid = Util.getValor("sid");
	
		 this.connection = new ConnectionFactory().getConnection(sid);
	 
	 }
	 
	 public void deletarPedidos() throws SQLException, IOException {
		 
		 try {
			 
			 String sql = "DELETE FROM Tb_Faturamento";
			 
			 PreparedStatement statement = this.connection.prepareStatement(sql);
			 
			 statement.executeUpdate();
			 
			 if (statement != null) {
				 statement.close();
			 }
			 
		 } catch (SQLException e) {
			 throw new RuntimeException(e);
		 } finally {
			 this.connection.close();
		 }
		 
	 }
	 
	 public List<Pedido> recuperaPedidos(Pedido pedido) throws SQLException, IOException {
		 
		 List<Pedido> listaPedidosNaoFaturados = new ArrayList<Pedido>();
		 
		 try {
			 
			 String sql = "SELECT Numero_Pedido, Contrato_SAP, Erro_No_Pedido FROM Tb_Faturamento WHERE Numero_Pedido = ?";
			 
			 PreparedStatement statement = this.connection.prepareStatement(sql);
			 statement.setString(1, pedido.getNumero().trim());
			 
			 ResultSet rs = statement.executeQuery();
			 
            while ( rs.next() ) {
                
            	String numeroPedido = rs.getString("Numero_Pedido");
            	String numeroContractNumber = rs.getString("Contrato_SAP");
                String mensagemDeErroNoPedido = rs.getString("Erro_No_Pedido");
                
                Pedido pedidoNaoFaturado = new Pedido();
                pedidoNaoFaturado.setNumero(numeroPedido);
                pedidoNaoFaturado.setNumeroContractNumber(numeroContractNumber);
                pedidoNaoFaturado.setMensagemDeErroNoPedido(mensagemDeErroNoPedido);
                
                listaPedidosNaoFaturados.add(pedidoNaoFaturado);
                
            }

			 if (statement != null) {
				 statement.close();
			 }
			 
		 } catch (SQLException e) {
			 throw new RuntimeException(e);
		 } finally {
			 this.connection.close();
		 }
		
		 return listaPedidosNaoFaturados;
		 
	 }

	 public void inserirPedido(Pedido pedido) throws SQLException, IOException {
		 
		 try {
			 
			 String sql = "INSERT INTO Tb_Faturamento (               ";
			 sql += "Contrato,                                        ";
			 sql += "Frente,                                          ";
			 sql += "Contrato_SAP,                                    ";
			 sql += "WBS,                                             ";
			 sql += "Numero_Pedido,                                   ";
			 sql += "Data_Pedido,                                     ";
			 sql += "Valor_Pedido,                                    ";
			 sql += "Cnpj_Cliente,                                    ";
			 sql += "Comprador,                                       ";
			 sql += "Faturado,                                        ";
			 sql += "Salvo_No_Sharepoint,                             ";
			 sql += "Campo_Observacao_No_Sharepoint,                  ";
			 sql += "Data_Extração,                                   ";
			 sql += "Erro_No_Pedido		                  			  ";
			 sql +=  ") VALUES (                                      ";
			 sql += "?,                                               ";
			 sql += "?,                                               ";
			 sql += "?,                                               ";
			 sql += "?,                                               ";
			 sql += "?,                                               ";
			 sql += "?,                                               ";
			 sql += "?,                                               ";
			 sql += "?,                                               ";
			 sql += "?,                                               ";
			 sql += "?,                                               ";
			 sql += "?,                                               ";
			 sql += "?,                                               ";
			 sql += "?,                                               ";
			 sql += "?                                                ";			 
			 sql += ")";
			 
			 PreparedStatement statement = this.connection.prepareStatement(sql);
			 statement.setString(1, pedido.getContractNumber().getContrato());
			 statement.setString(2, pedido.getContractNumber().getFrente());
			 statement.setString(3, pedido.getContractNumber().getNumero());
			 statement.setString(4, pedido.getContractNumber().getWbs());
			 statement.setString(5, pedido.getNumero());
			 statement.setString(6, pedido.getData());
			 statement.setString(7, String.valueOf(pedido.getValor()));
			 statement.setString(8, pedido.getCnpjCliente());
			 statement.setString(9, pedido.getComprador());
			 statement.setString(10, pedido.isFaturado() ? "Sim" : "Não" );
			 
	         // Se o pedido não está faturado, mostro a informação se foi salvo no sharepoint
	         if (!pedido.isFaturado()) {
	        	 statement.setString(11,  pedido.isSalvoNoSharepoint() ? "Sim" : "Não" );
	         // Se o pedido está faturado, subentende-se que está salvo no sharepoint
	         } else {
	        	 statement.setString(11, "Sim");
	         }
	         statement.setString(12, pedido.getObservacaoSharepoint());
	         statement.setString(13, pedido.getDataExtracao());
	         statement.setString(14, pedido.getMensagemDeErroNoPedido());
			 
			 statement.executeUpdate();
			 
			 if (statement != null) {
				 statement.close();
			 }
			 
		 } catch (SQLException e) {
			 
			 String mensagem1 = "Houve algum problema no momento da inserção deste pedido: " + "\n";
			 
			 mensagem1 += pedido.getContractNumber().getContrato() + ", " + pedido.getContractNumber().getNumero() + ", " + pedido.getNumero() + "\n";
			 
			 mensagem1 += "Esta é a mensagem de erro: "  + "\n";
			 
			 mensagem1 += e.getMessage() + "\n";
			 
			 if (mensagem1.contains("String or binary data would be truncated")) {
				 
				 mensagem1 += "A mensagem acima - String or binary data would be truncated - significa que algum campo do registro estourou o limite de tamanho permitido para o seu respectivo campo na tabela do banco." ;
			 }
			 
			 throw new RuntimeException(mensagem1);
		 } finally {
			 this.connection.close();
		 }
		 
	 }
	 
}

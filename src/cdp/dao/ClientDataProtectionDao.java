package cdp.dao;

import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

import bean.RelatorioControlsDueThisMonth;
import bean.RelatorioOperationalRiskIndexByClient;
import cdp.util.ConnectionFactory;
import cdp.util.Util;

import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;


public class ClientDataProtectionDao {
	
	private Connection connection;
	
	 public ClientDataProtectionDao() throws IOException {
		 
		 String sid = Util.getValor("sid");
	
		 this.connection = new ConnectionFactory().getConnection(sid);
	 
	 }
	 
	 public void deletarRelatorioControlsDueThisMonth() throws SQLException, IOException {
		 
		 try {
			 
			 String sql = "DELETE FROM CDP_Controls_Due";
			 
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
	 
	 public void deletarRelatorioOperationalRiskIndexByClient() throws SQLException, IOException {
		 
		 try {
			 
			 String sql = "DELETE FROM CDP_Operational_Risc";
			 
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


	 public void inserirRelatorioControlsDueThisMonth(RelatorioControlsDueThisMonth relatorioControlsDueThisMonth) throws SQLException, IOException {
		 
		 try {
			 
			 String sql = "INSERT INTO CDP_Controls_Due (         ";
			 sql += "Client_Name,                                 ";
			 sql += "Control_ID,                                  ";
			 sql += "Indicator,                                   ";
			 sql += "Verified,                                    ";
			 sql += "Service_Name,                                ";
			 sql += "Client_Data_Protection_Requirement,		  ";
			 sql += "Related_Contractual_Requirement,             ";
			 sql += "Regulatory_Requirement,                      ";
			 sql += "Engagement_Level_Procedure,                  ";
			 sql += "Compliance,                                  ";
			 sql += "Ongoing_Frequency,                           ";
			 sql += "Control_Name,                                ";
			 sql += "Control_Owner,                               ";
			 sql += "Next_DueDate,                                ";
			 sql += "Delivery_Location_Source,                    ";
			 sql += "Data_Extracao                                ";
			 sql +=  ") VALUES (                                  ";
			 sql += "?,                                           ";
			 sql += "?,                                           ";
			 sql += "?,                                           ";
			 sql += "?,                                           ";
			 sql += "?,                                           ";
			 sql += "?,                                           ";
			 sql += "?,                                           ";
			 sql += "?,                                           ";
			 sql += "?,                                           ";
			 sql += "?,                                           ";
			 sql += "?,                                           ";
			 sql += "?,                                           ";
			 sql += "?,                                           ";
			 sql += "?,                                           ";
			 sql += "?,                                           ";
			 sql += "?                                            ";			 
			 sql += ")";
			 
			 PreparedStatement statement = this.connection.prepareStatement(sql);
			 statement.setString(1, relatorioControlsDueThisMonth.getClientName());
			 statement.setString(2, relatorioControlsDueThisMonth.getControlID());
			 statement.setString(3, relatorioControlsDueThisMonth.getIndicator());
			 statement.setString(4, relatorioControlsDueThisMonth.getVerified());
			 statement.setString(5, relatorioControlsDueThisMonth.getServiceName());
			 statement.setString(6, relatorioControlsDueThisMonth.getClientDataProtectionRequirement());
			 statement.setString(7, relatorioControlsDueThisMonth.getRelatedContractualRequirement());
			 statement.setString(8, relatorioControlsDueThisMonth.getRegulatoryRequirement());
			 statement.setString(9, relatorioControlsDueThisMonth.getEngagementLevelProcedure());
			 statement.setString(10, relatorioControlsDueThisMonth.getCompliance());
			 statement.setString(11, relatorioControlsDueThisMonth.getOngoingFrequency());
			 statement.setString(12, relatorioControlsDueThisMonth.getControlName());
			 statement.setString(13, relatorioControlsDueThisMonth.getControlOwner());
			 statement.setString(14, relatorioControlsDueThisMonth.getNextDueDate());
			 statement.setString(15, relatorioControlsDueThisMonth.getDeliveryLocationSource());
			 statement.setString(16, relatorioControlsDueThisMonth.getDataExtracao());
			 
			 statement.executeUpdate();
			 
			 if (statement != null) {
				 statement.close();
			 }
			 
		 } catch (SQLException e) {
			 
			 String mensagem = e.getMessage() + "\n";
			 
			 if (mensagem.contains("String or binary data would be truncated")) {
				 
				 mensagem += "A mensagem acima - String or binary data would be truncated - significa que algum campo do registro estourou o limite de tamanho permitido para o seu respectivo campo na tabela do banco." ;
			 }
			 
			 throw new RuntimeException(mensagem);
		 } finally {
			 this.connection.close();
		 }
		 
	 }
	 
	 public void inserirRelatorioOperationalRiskIndexByClient(RelatorioOperationalRiskIndexByClient relatorioOperationalRiskIndexByClient) throws SQLException, IOException {
		 
		 try {
			 
			 String sql = "INSERT INTO CDP_Operational_Risc (         ";
			 sql += "CDPTracker_ID,                                   ";
			 sql += "Client_Name,                                     ";
			 sql += "Tier,                                            ";
			 sql += "Master_Customer_Nbr,                             ";
			 sql += "Market,                                          ";
			 sql += "Market_Unit,                                     ";
			 sql += "Owning_Organization,                             ";
			 sql += "Accountable_MD,                                  ";
			 sql += "Account_ISL,                                     ";
			 sql += "Account_Contract_Manager,                        ";
			 sql += "CDP_Account_Manager,                             ";
			 sql += "CDPControl_ID,                               	  ";
			 sql += "Control_Name,                                    ";
			 sql += "Control_Category,                                ";
			 sql += "Delivery_Location_Source,                        ";
			 sql += "Control_Owner,                                   ";
			 sql += "Compliance_Status,                               ";
			 sql += "Last_Completed_Date,                             ";
			 sql += "Next_Due_Date,                                   ";
			 sql += "Nbr_Days_Non_Compliant,                          ";
			 sql += "Nbr_Days_Past_Due,                               ";
			 sql += "Weighted_Operational_Risk,                       ";
			 sql += "Month_End_Weighted_Operational_Risk,			  ";
			 sql += "Data_Extracao                                    ";
			 sql +=  ") VALUES (                            	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?,                                     	      ";
			 sql += "?                                      	      ";			 
			 sql += ")";
			 
			 PreparedStatement statement = this.connection.prepareStatement(sql);
			 statement.setString(1, relatorioOperationalRiskIndexByClient.getCDPTrackerID());
			 statement.setString(2, relatorioOperationalRiskIndexByClient.getClientName());
			 statement.setString(3, relatorioOperationalRiskIndexByClient.getTier());
			 statement.setString(4, relatorioOperationalRiskIndexByClient.getMasterCustomerNbr());
			 statement.setString(5, relatorioOperationalRiskIndexByClient.getMarket());
			 statement.setString(6, relatorioOperationalRiskIndexByClient.getMarketUnit());
			 statement.setString(7, relatorioOperationalRiskIndexByClient.getOwningOrganization());
			 statement.setString(8, relatorioOperationalRiskIndexByClient.getAccountableMD());
			 statement.setString(9, relatorioOperationalRiskIndexByClient.getAccountISL());
			 statement.setString(10, relatorioOperationalRiskIndexByClient.getAccountContractManager());
			 statement.setString(11, relatorioOperationalRiskIndexByClient.getCDPAccountManager());
			 statement.setString(12, relatorioOperationalRiskIndexByClient.getCDPControlID());
			 statement.setString(13, relatorioOperationalRiskIndexByClient.getControlName());
			 statement.setString(14, relatorioOperationalRiskIndexByClient.getControlCategory());
			 statement.setString(15, relatorioOperationalRiskIndexByClient.getDeliveryLocationSource());
			 statement.setString(16, relatorioOperationalRiskIndexByClient.getControlOwner());
			 statement.setString(17, relatorioOperationalRiskIndexByClient.getComplianceStatus());
			 statement.setString(18, relatorioOperationalRiskIndexByClient.getLastCompletedDate());
			 statement.setString(19, relatorioOperationalRiskIndexByClient.getNextDueDate());
			 statement.setString(20, relatorioOperationalRiskIndexByClient.getNbrDaysNonCompliant());
			 statement.setString(21, relatorioOperationalRiskIndexByClient.getNbrDaysPastDue());
			 statement.setString(22, relatorioOperationalRiskIndexByClient.getWeightedOperationalRisk());
			 statement.setString(23, relatorioOperationalRiskIndexByClient.getMonthEndWeightedOperationalRisk());
			 statement.setString(24, relatorioOperationalRiskIndexByClient.getDataExtracao());

			 statement.executeUpdate();
			 
			 if (statement != null) {
				 statement.close();
			 }
			 
		 } catch (SQLException e) {
			 
			 String mensagem = e.getMessage() + "\n";
			 
			 if (mensagem.contains("String or binary data would be truncated")) {
				 
				 mensagem += "A mensagem acima - String or binary data would be truncated - significa que algum campo do registro estourou o limite de tamanho permitido para o seu respectivo campo na tabela do banco." ;
			 }
			 
			 throw new RuntimeException(mensagem);
		 } finally {
			 this.connection.close();
		 }
		 
	 }

	 
}

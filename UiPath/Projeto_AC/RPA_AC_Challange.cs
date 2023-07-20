using System;
using System.Collections.Generic;
using System.Data;
using UiPath.CodedWorkflows;
using UiPath.Core;
using UiPath.Core.Activities.Storage;
using UiPath.Orchestrator.Client.Models;
using UiPath.Testing;
using UiPath.Testing.Activities.TestData;
using UiPath.Testing.Activities.TestDataQueues.Enums;
using UiPath.Testing.Enums;
using UiPath.UIAutomationNext.API.Contracts;
using UiPath.UIAutomationNext.API.Models;
using UiPath.UIAutomationNext.Enums;
using System.Data.OleDb;

namespace CA_RPAChallenge
{
    public class RPAChallenge : Projeto_AC.CodedWorkflow
    {
        [Workflow]
        public void Execute()
        {
			string InputfilePath = @"E:\LSR Tecnologias\Clientes\Web\Sicredi\UiPath\challenge.xlsx";
			string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={InputfilePath};Extended Properties='Excel 12.0;HDR=YES;IMEX=1;'";
			DataTable dtPlanilhaExcel = new DataTable();
			
			using (OleDbConnection connection = new OleDbConnection(connectionString))
			{
				connection.Open();
				string sheetName = "Sheet1";
				string query = $"SELECT * FROM [{sheetName}$]";
				
				OleDbCommand command = new OleDbCommand(query,connection);
				
				using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
				{
					adapter.Fill(dtPlanilhaExcel);
				}				
			}

			foreach (DataRow row in dtPlanilhaExcel.Rows)
			{
				Log(row["First Name"] + "\t");
			}
			
			var rpaACChallenge = uiAutomation.Open("Chrome: Rpa Challenge");
			rpaACChallenge.Click("btnStart");
			
			foreach (DataRow row in dtPlanilhaExcel.Rows)
			{
				rpaACChallenge.TypeInto("First Name", row["First Name"].ToString());
				rpaACChallenge.TypeInto("Last Name", row["Last Name"].ToString());
				rpaACChallenge.TypeInto("Email", row["Email"].ToString());
				rpaACChallenge.TypeInto("Address", row["Address"].ToString());
				rpaACChallenge.TypeInto("Phone Number", row["Phone Number"].ToString());
				rpaACChallenge.TypeInto("Company Name", row["Company Name"].ToString());
				rpaACChallenge.TypeInto("Role in Company", row["Role in Company"].ToString());
				rpaACChallenge.Click("btnSubmit");
			}			
        }
    }
}
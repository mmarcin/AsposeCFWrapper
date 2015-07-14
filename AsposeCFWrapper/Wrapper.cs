using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace Aspose.Words
{

    class AsposeCFWrapper
    {
        public string TemplateDir;          // c:/projects/eProcurement/www_root/prototypes/         
        public string TemplateName;         // Template.docx

        public string OutputDir;            // c:/temp/
        public string OutputDocumentName;   // OutputFile.docx

        public string SelectString;         // SELECT * FROM pro_tProcurement
        public string TableNameString;      // Procutements

        public AsposeCFWrapper(string TemplateDir, string TemplateName, string OutputDir, string OutputDocumentName, string SelectString, string TableNameString)
        {
            this.TemplateDir = TemplateDir;
            this.TemplateName = TemplateName;
            this.OutputDir = OutputDir;
            this.OutputDocumentName = OutputDocumentName;
            this.SelectString = SelectString;
            this.TableNameString = TableNameString;
        }

        public void ExecuteRegions()
        {
            Document doc = new Document(this.TemplateDir + this.TemplateName);                      // Create Aspose document object

            DataTable TableWithData = GetDatabaseResults(this.SelectString, this.TableNameString);  // Get Data

            doc.MailMerge.ExecuteWithRegions(TableWithData);                                        // Make MailMerge With regions

            doc.Save(this.OutputDir + this.OutputDocumentName);                                     // Save the result
        }

        private static DataTable GetDatabaseResults(string SelectString, string TableNameString)
        {
            DataTable table = ExecuteDataTable(SelectString);
            table.TableName = TableNameString ;
            return table;
        }


        // ------------------------------------------------------------------------------------------------------------------------------------
        // Toto je povodny priklad z Aspose. Ked to pobezi, tak to zmazeme
        // ------------------------------------------------------------------------------------------------------------------------------------
        public void ExecuteWithRegionsDataTable()
        {
            Document doc = new Document(TemplateDir + TemplateName);

            int orderId = 10444;

            // Perform several mail merge operations populating only part of the document each time.

            // Use DataTable as a data source.
            DataTable orderTable = GetTestOrder(orderId);
            doc.MailMerge.ExecuteWithRegions(orderTable);

            // Instead of using DataTable you can create a DataView for custom sort or filter and then mail merge.
            DataView orderDetailsView = new DataView(GetTestOrderDetails(orderId));
            orderDetailsView.Sort = "ExtendedPrice DESC";
            doc.MailMerge.ExecuteWithRegions(orderDetailsView);

            doc.Save(OutputDir + OutputDocumentName);
        }

        private static DataTable GetTestOrder(int orderId)
        {
            DataTable table = ExecuteDataTable(string.Format(
                "SELECT * FROM AsposeWordOrders WHERE OrderId = {0}", orderId));
            table.TableName = "Orders";
            return table;
        }

        private static DataTable GetTestOrderDetails(int orderId)
        {
            DataTable table = ExecuteDataTable(string.Format(
                "SELECT * FROM AsposeWordOrderDetails WHERE OrderId = {0} ORDER BY ProductID", orderId));
            table.TableName = "OrderDetails";
            return table;
        }

        /// <summary>
        /// Utility function that creates a connection, command, 
        /// executes the command and return the result in a DataTable.
        /// </summary>
        private static DataTable ExecuteDataTable(string commandText)
        {
            // Open the database connection.
            string connString = "Server=127.0.0.1;Database=eProcurement;User Id=sa;Password=Lomtec2000;";
            SqlConnection  conn = new SqlConnection (connString);
            conn.Open();

            // Create and execute a command.
            SqlCommand cmd = new SqlCommand(commandText, conn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable table = new DataTable();
            da.Fill(table);

            // Close the database.
            conn.Close();

            return table;
        }
    }
}

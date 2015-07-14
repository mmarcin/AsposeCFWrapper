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
        string MyDir = "c:/projetcs/eProcurement/www_root/prototypes/";

        public void ExecuteWithRegionsDataTable()
        {
            Document doc = new Document(MyDir + "MailMerge.ExecuteWithRegions.doc");

            int orderId = 10444;

            // Perform several mail merge operations populating only part of the document each time.

            // Use DataTable as a data source.
            DataTable orderTable = GetTestOrder(orderId);
            doc.MailMerge.ExecuteWithRegions(orderTable);

            // Instead of using DataTable you can create a DataView for custom sort or filter and then mail merge.
            DataView orderDetailsView = new DataView(GetTestOrderDetails(orderId));
            orderDetailsView.Sort = "ExtendedPrice DESC";
            doc.MailMerge.ExecuteWithRegions(orderDetailsView);

            doc.Save(MyDir + "MailMerge.ExecuteWithRegionsDataTable Out.doc");
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

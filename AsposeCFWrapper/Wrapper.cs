using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Aspose.Words;

namespace AsposeCF
{

    public class Wrapper
    {
        public string TemplateDir;          // Dir, where the template file is located      
        public string TemplateName;         // Template.docx

        public string OutputDir;            // Dir, where the output file will be written - c:/temp/
        public string OutputDocumentName;   // OutputFile.docx

        public Document doc = new Document();
        
        public Wrapper(string TemplateDir, string TemplateName, string OutputDir, string OutputDocumentName)
        {
            this.TemplateDir = TemplateDir;
            this.TemplateName = TemplateName;
            this.OutputDir = OutputDir;
            this.OutputDocumentName = OutputDocumentName;
            doc = new Document(this.TemplateDir + this.TemplateName);  // Create Aspose document object
        }

        public void Execute(string[] names, object[] values)
        {
            doc.MailMerge.Execute(names, values);
        }

        public void ExecuteRegions(string SelectString, string TableName)
        {            
            DataTable TableWithData = GetDatabaseResults(SelectString, TableName);  // Get Data
            doc.MailMerge.ExecuteWithRegions(TableWithData);                         // Make MailMerge With regions            
        }

        public void Save()
        {
            doc.Save(this.OutputDir + this.OutputDocumentName); // Save the result
        }


        private static DataTable GetDatabaseResults(string SelectString, string TableNameString)
        {
            DataTable table = ExecuteDataTable(SelectString);
            table.TableName = TableNameString ;
            return table;
        }

        private static DataTable ExecuteDataTable(string commandText)
        {
            // Open the database connection.
            string connString = "Server=local.ebiz.sk;Database=eProcurement;User Id=sa;Password=Lomtec2000;";
            SqlConnection conn = new SqlConnection (connString);
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

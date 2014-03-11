using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data.OleDb;
using Aspose.Cells;
using System.Data;
using FISCA.UDT;
using FISCA.Data;
using System.Windows.Forms;
using EMBA.Export.Forms;

namespace EMBA.Export
{
    public static class Extensions
    {
        public static List<string> ResolveField(this string SQL)
        {
            List<string> fields = new List<string>();
            try
            {
                QueryHelper Access = new QueryHelper();

                DataTable dataTable = Access.Select(SQL + " limit 0;");
                foreach (DataColumn column in dataTable.Columns)
                {
                    if (!fields.Contains(column.ColumnName))
                        fields.Add(column.ColumnName);
                }
            }
            catch { }

            return fields;
        }

        public static DataTable ToDataTable(this DBF_Table Table)
        {
            DataTable dataTable = new DataTable();
            dataTable.TableName = Table.TableName;

            foreach (IDBF_Field_Schema field in Table.Fields)
            {
                dataTable.Columns.Add(new DataColumn(field.Name));
            }

            foreach (DBF_Row row in Table.Rows)
            {
                DataRow newRow = dataTable.NewRow();

                foreach (IDBF_Field_Schema field in Table.Fields)
                {
                    newRow[field.Name] = (row[field.Name] + "");
                }

                dataTable.Rows.Add(newRow);
            }

            return dataTable;
        }

        public static bool ToDBF(this DBF_Table Table, string FileName, bool OpenDirectory)
        {
            DataTable dataTable = new DataTable();
            dataTable.TableName = Table.TableName;

            foreach (IDBF_Field_Schema field in Table.Fields)
            {
                dataTable.Columns.Add(new DataColumn(field.Name));
            }

            foreach (DBF_Row row in Table.Rows)
            {
                DataRow newRow = dataTable.NewRow();

                foreach (IDBF_Field_Schema field in Table.Fields)
                {
                    newRow[field.Name] = (row[field.Name] + "");
                }

                dataTable.Rows.Add(newRow);
            }

            FileInfo fileInfo = new FileInfo(FileName);
            dataTable.TableName = fileInfo.Name;
            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(dataTable);

            try
            {
                bool result = Extensions.ExportToDBF(dataSet, fileInfo.Name, fileInfo.DirectoryName);
                if (result && OpenDirectory)
                {
                    System.Diagnostics.Process.Start(fileInfo.DirectoryName);
                }
                return result;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private static bool ExportToDBF(DataSet dsExport, string tableName, string FilePath)
        {
            string connectionString = String.Format(@"Provider=VFPOLEDB; Data Source={0}; Extended Properties=dBase III;", FilePath);
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                using (OleDbCommand command = connection.CreateCommand())
                {
                    try
                    {
                        connection.Open();
                    }
                    catch 
                    {
                        ShowHelpLinkForm frm = new ShowHelpLinkForm();
                        frm.Text = "錯誤";
                        frm.SetMessage("產生「DBF」檔案需要「VFPOLEDB」驅動程式，請點選以下連結下載並安裝。");
                        frm.SetLinkURL("http://download.microsoft.com/download/b/f/b/bfbfa4b8-7f91-4649-8dab-9a6476360365/VFPOLEDBSetup.msi");
                        frm.ShowDialog();
                        return false;
                    }
                    StringBuilder createStatement = new StringBuilder("Create Table " + tableName + " ( ");
                    for (int iCol = 0; iCol < dsExport.Tables[0].Columns.Count; iCol++)
                    {
                        createStatement.Append("[" + dsExport.Tables[0].Columns[iCol].ColumnName.ToString() + "]");
                        if (iCol == dsExport.Tables[0].Columns.Count - 1)
                        {
                            createStatement.Append(" varchar(250) )");
                        }
                        else
                        {
                            createStatement.Append(" varchar(250), ");
                        }
                    }
                    command.CommandText = createStatement.ToString();
                    command.ExecuteNonQuery();
                }
            }
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                OleDbCommand commandInsert = new OleDbCommand();
                OleDbTransaction transaction = null;
                try
                {
                    commandInsert.Connection = connection;
                    transaction = connection.BeginTransaction();
                    commandInsert.Transaction = transaction;

                    //string insertStatement = "Insert Into " + tableName + " Values ( ";
                    //string insertTemp = string.Empty; 

                    //  SqlCommand sqlcmd = new SqlCommand("INSERT INTO myTable (c1, c2, c3, c4) VALUES (@c1, @c2, @c3, @c4)", sqlconn);
                    //  sqlcmd.Parameters.AddWithValue("@c1", 1); // 設定參數 @c1 的值。
                    for (int row = 0; row < dsExport.Tables[0].Rows.Count; row++)
                    {
                        commandInsert.Parameters.Clear();
                        StringBuilder insertTemp = new StringBuilder("Insert Into " + tableName + " Values ( ");
                        for (int col = 0; col < dsExport.Tables[0].Columns.Count; col++)
                        {                           
                            // Create Insert Statement
                            if (col == dsExport.Tables[0].Columns.Count - 1)
                                insertTemp.Append("? ) ;");
                            else
                                insertTemp.Append("? , ");
                            commandInsert.Parameters.Add(dsExport.Tables[0].Columns[col].ColumnName, OleDbType.Char).Value = dsExport.Tables[0].Rows[row][col]; 
                        }
                        commandInsert.CommandText = insertTemp.ToString();
                        commandInsert.ExecuteNonQuery();
                    } 
                    // Commit the transaction.
                    transaction.Commit();
                }
                catch (Exception ex)
                {
                    try
                    {
                        // Attempt to roll back the transaction.
                        transaction.Rollback();
                    }
                    catch   {   }
                    finally
                    {
                        throw new Exception(ex.Message);
                    }
                }
            }
            return true;
        }

        public static bool ValidateBySchema(this DBF_Table table)
        {
            foreach (DBF_Field field in table.Fields)
            {
                //if (!field.AllowNull && String.IsNullOrEmpty(field.Value))
                //    return false;

                //if (field.Size < field.Value.Length)
                //    return false;
            }

            return false;
        }

        public static Workbook ToWorkbook(this DataTable dataTable, bool autoFitColumns)
        {
            return ToWorkbook(dataTable, autoFitColumns, null);
        }

        public static Workbook ToWorkbook(this DataTable dataTable, bool autoFitColumns, List<string> SelectedFields)
        {
            Workbook wb = new Workbook();

            int i = -1;
            for (int columnIndex = 0; columnIndex < dataTable.Columns.Count; columnIndex++)
            {
                if (SelectedFields != null && !SelectedFields.Contains(dataTable.Columns[columnIndex].ColumnName))
                    continue;

                i++;
                wb.Worksheets[0].Cells[0, i].PutValue(dataTable.Columns[columnIndex].ColumnName);
            }

            if (dataTable.Rows.Count == 0)
                return wb;

            for (int rowIndex = 0; rowIndex < dataTable.Rows.Count; rowIndex++)
            {
                i = -1;
                for (int columnIndex = 0; columnIndex < dataTable.Columns.Count; columnIndex++)
                {
                    if (SelectedFields != null && !SelectedFields.Contains(dataTable.Columns[columnIndex].ColumnName))
                        continue;

                    i++;
                    wb.Worksheets[0].Cells[rowIndex + 1, i].PutValue(dataTable.Rows[rowIndex][columnIndex] + "");
                }
            }
            if (autoFitColumns)
                wb.Worksheets[0].AutoFitColumns();

            return wb;
        }

        public static string Save(this Workbook wb, bool open)
        {
            return Extensions.Save(wb, open, wb.Worksheets[0].Name);
        }

        public static string Save(this Workbook wb, bool open, string fileName)
        {
            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = fileName + ".xls";
            sd.Filter = "Excel 2003 相容檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    wb.Save(sd.FileName, FileFormatType.Excel2003);
                }
                catch
                {
                    MessageBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return string.Empty;
                }
            }

            if (open && System.IO.File.Exists(sd.FileName))
                System.Diagnostics.Process.Start(sd.FileName);

            return sd.FileName;
        }

        public static string Save(this Workbook wb, bool open, string fileName, bool showSaveFileDialog)
        {
            if (showSaveFileDialog)
                return Extensions.Save(wb, open, fileName);

            try
            {
                wb.Save(fileName, FileFormatType.Excel2003);
            }
            catch
            {
                MessageBox.Show("指定路徑無法存取。", "建立檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return string.Empty;
            }

            if (open && System.IO.File.Exists(fileName))
                System.Diagnostics.Process.Start(fileName);

            return fileName;
        }
    }
}

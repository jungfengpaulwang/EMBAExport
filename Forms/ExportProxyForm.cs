using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using FISCA.Data;
using FISCA.LogAgent;
using FISCA.Presentation.Controls;

namespace EMBA.Export
{
    /// <summary>
    /// 使用者選擇欄位
    /// </summary>
    public partial class ExportProxyForm : BaseForm
    {
        private Dictionary<string, ListViewItem> export_Fields;
        private string query_SQL;
        private BackgroundWorker _BGWLoadData;
        private QueryHelper queryHelper;
        private string keyField;
        private IEnumerable<string> invisibleFields;
        private Dictionary<string, List<KeyValuePair<string, string>>> replaceFields;
        private bool autoSaveLog;
        private bool autoSaveFile;
        private List<string> selectedFields;
        private List<string> resolvedFields;
        private bool FieldsResolved;
        public bool HideMeWhenProcessStart { get; set; }

        public ExportProxyForm() : this(true)
        {
        }

        public ExportProxyForm(bool HideMeWhenProcessStart = true) 
        {
            InitializeComponent();
            this.HideMeWhenProcessStart = HideMeWhenProcessStart;
            this.Load += new System.EventHandler(this.ExportProxyForm_Load);
            this.selectedFields = new List<string>();
        }

        public ExportProxyForm(string QuerySQL, bool AutoSaveLog, bool AutoSaveFile, string KeyField, IEnumerable<string> InvisibleFields, Dictionary<string, List<KeyValuePair<string, string>>> ReplaceFields, bool HideMeWhenProcessStart = true)
            : this(HideMeWhenProcessStart)
        {
            this.HideMeWhenProcessStart = HideMeWhenProcessStart;
            this.keyField = KeyField;

            this.invisibleFields = InvisibleFields;
            this.replaceFields = ReplaceFields;

            this.QuerySQL = QuerySQL;

            this.autoSaveFile = AutoSaveFile;
            this.autoSaveLog = AutoSaveLog;
        }

        public BackgroundWorker SalvageOperation
        {
            get
            {
                return _BGWLoadData;
            }
        }

        public string QuerySQL
        {
            get { return this.query_SQL; }
            set 
            { 
                this.query_SQL = value;

                if (!this.FieldsResolved)
                {
                    this.ResolvedFields = query_SQL.ResolveField();
                    this.ResolveField();
                    this.FieldsResolved = true;
                }
            }
        }

        public List<string> SelectedFields
        {
            get { return this.selectedFields; }
        }

        public Dictionary<string, ListViewItem> Fields
        {
            get
            {
                return export_Fields;
            }
        }

        public bool AutoSaveFile
        {
            get { return this.autoSaveFile; }
            set { this.autoSaveFile = value; }
        }

        public bool AutoSaveLog
        {
            get { return this.autoSaveLog; }
            set { this.autoSaveLog = value; }
        }

        public string KeyField
        {
            get { return this.keyField; }
            set { this.keyField = value; }
        }

        public IEnumerable<string> InvisibleFields
        {
            get { return this.invisibleFields; }
            set { this.invisibleFields = value; }
        }

        public Dictionary<string, List<KeyValuePair<string, string>>> ReplaceFields
        {
            get { return this.replaceFields; }
            set { this.replaceFields = value; }
        }

        public List<string> ResolvedFields
        {
            get { return this.resolvedFields; }
            set { this.resolvedFields = value; }
        }

        //public void SetResolvedFields(List<string> fields)
        //{
        //    this.resolvedFields = fields;
        //}

        private void _BGWLoadData_DoWork(object sender, DoWorkEventArgs e)
        {
            e.Result = queryHelper.Select(query_SQL);
        }

        private void _BGWLoadData_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            DataTable dataTable = e.Result as DataTable;

            this.selectedFields = new List<string>();
            dataTable.Columns.Cast<DataColumn>().ToList().ForEach((x) =>
            {
                if (this.export_Fields.ContainsKey(x.ColumnName))
                {
                    if (!this.selectedFields.Contains(x.ColumnName) && this.export_Fields[x.ColumnName].Checked)
                        this.selectedFields.Add(x.ColumnName);
                }
            });

            if (this.autoSaveFile)
                SaveFile(dataTable);

            if (this.autoSaveLog)
                SaveLog(dataTable);
        }

        private void SaveFile(DataTable dataTable)
        {
            if (dataTable == null || dataTable.Rows.Count == 0)
            {
                FISCA.Presentation.Controls.MsgBox.Show("無資料可匯出。", "欄位空白", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (this.replaceFields != null && this.replaceFields.Count > 0)
            {
                foreach(DataRow row in dataTable.Rows)
                {
                    foreach(string field in replaceFields.Keys)
                    {
                        if (!dataTable.Columns.Contains(field))
                            continue;

                        IEnumerable<KeyValuePair<string, string>> filterKVs = replaceFields[field].Where(x=>x.Key.ToUpper() == (row[field] + "").Trim().ToUpper());
                        if (filterKVs.Count() > 0)
                            row[field] = filterKVs.ElementAt(0).Value;
                    }
                }
            }
            dataTable.ToWorkbook(true, this.selectedFields).Save(true, this.Tag as string);
        }

        private void SaveLog(DataTable dataTable)
        {
            LogSaver logBatch = ApplicationLog.CreateLogSaverInstance();
            List<string> system_IDs = new List<string>();
            List<string> Fields = new List<string>();

            if (this.selectedFields == null || this.selectedFields.Count == 0)
                return;

            dataTable.Columns.Cast<DataColumn>().ToList().ForEach((x) =>
            {
                if (this.selectedFields.Contains(x.ColumnName) && !Fields.Contains(x.ColumnName)) 
                    Fields.Add(x.ColumnName);
            });
            dataTable.Rows.Cast<DataRow>().ToList().ForEach(x => system_IDs.Add(x[this.keyField] + ""));
            string strSystemIDs = String.Join(",", system_IDs);
            string strFields = String.Join(",", Fields);
            string strCategory = string.Empty;

            if (this.Text.IndexOf("學生") > 0) strCategory = "student";
            if (this.Text.IndexOf("班級") > 0) strCategory = "class";
            if (this.Text.IndexOf("教師") > 0) strCategory = "teacher";
            if (this.Text.IndexOf("課程") > 0) strCategory = "course";

            logBatch.AddBatch(this.Text, strCategory, this.keyField + "：" + strSystemIDs, "匯出欄位：" + strFields);
            try
            {
                logBatch.LogBatch(true);
            }
            catch
            {

            }
        }

        public virtual void SetQueryString(string SQL)
        {
            this.query_SQL = SQL; 
        }

        public void ResolveField()
        {
            List<string> resolve_Fields = this.ResolvedFields;
            export_Fields = new Dictionary<string, ListViewItem>();
            this.FieldContainer.Clear();
            this.selectedFields.Clear();
            for (int i = 0; i < resolve_Fields.Count; i++)
            {
                if (this.invisibleFields != null && this.invisibleFields.Count() > 0 && this.invisibleFields.Contains(resolve_Fields[i]))
                    continue;

                ListViewItem item = FieldContainer.Items.Add(resolve_Fields[i]);
                item.Checked = true;
                if (!this.selectedFields.Contains(item.Text))
                    this.selectedFields.Add(item.Text);
                if (!export_Fields.ContainsKey(resolve_Fields[i]))
                    export_Fields.Add(resolve_Fields[i], item);
            }
        }

        private void ExportProxyForm_Load(object sender, EventArgs e)
        {
            if (this.DesignMode)
                return;

            this.btnExport.Click += new System.EventHandler(this.OnExportButtonClick);
            this.btnExit.Click += new System.EventHandler(this.OnExitButtonClick);
            this.chkSelectAll.CheckedChanged += new System.EventHandler(this.chkSelectAll_CheckedChanged);

            queryHelper = new QueryHelper();

            _BGWLoadData = new BackgroundWorker();
            _BGWLoadData.DoWork += new DoWorkEventHandler(_BGWLoadData_DoWork);
            _BGWLoadData.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGWLoadData_RunWorkerCompleted);
        }

        protected virtual void OnExportButtonClick(object sender, EventArgs e)
        {
            if (this.HideMeWhenProcessStart)
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
            else
                this.DialogResult = System.Windows.Forms.DialogResult.None;

            _BGWLoadData.RunWorkerAsync();
            //this.Close();
        }

        private void OnExitButtonClick(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            //this.Close();
        }

        private void chkSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            this.selectedFields.Clear();
            foreach (ListViewItem item in this.FieldContainer.Items)
            {
                item.Checked = chkSelectAll.Checked;
                if (item.Checked)
                    if (!this.selectedFields.Contains(item.Text))
                        this.selectedFields.Add(item.Text);
            }
        }

        private void chkSelectAll_CheckedChanged_1(object sender, EventArgs e)
        {

        }

        private void FieldContainer_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            if (e.Item.Checked)
                if (!this.selectedFields.Contains(e.Item.Text))
                    this.selectedFields.Add(e.Item.Text);
            else
                this.selectedFields.Remove(e.Item.Text);
        }
    }    
}
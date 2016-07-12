using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Resources;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace PivotAutoUpdate
{
    public partial class ThisAddIn
    {
        private Excel.Workbook _activeWorkbook;
        private Excel.Worksheet _activeWorksheet;
        private const string PropertyNamePivotAutoRefresh = "PivotAutoRefresh";

        public Excel.Worksheet ActiveWorksheet
        {
            get { return _activeWorksheet; }
            private set
            {
                if (_activeWorksheet != null)
                    _activeWorksheet.Change -= OnChangeWorksheet;
                _activeWorksheet = value;
                if (_activeWorksheet != null)
                    _activeWorksheet.Change += OnChangeWorksheet;
            }
        }

        public Excel.Workbook ActiveWorkbook
        {
            get { return _activeWorkbook; }
            private set
            {
                if (_activeWorkbook != null)
                    _activeWorkbook.SheetActivate -= OnWorksheetActivate;

                _activeWorkbook = value;

                if (_activeWorkbook != null)
                {
                    ActiveWorksheet = _activeWorkbook.ActiveSheet as Excel.Worksheet;
                    _activeWorkbook.SheetActivate += OnWorksheetActivate;

                    var properties = _activeWorkbook.CustomDocumentProperties as Office.DocumentProperties;
                    if (properties != null)
                    {
                        try
                        {
                            var autoupdate = properties[PropertyNamePivotAutoRefresh];
                            if (autoupdate != null)
                            {
                                _ribbon.AutoUpdateEnabled = (bool) autoupdate.Value;
                            }
                        }
                        catch { }
                    }
                    else
                    {
                        _ribbon.AutoUpdateEnabled = false;
                    }
                }
            }
        }

        private void OnChangeWorksheet(Excel.Range target)
        {
            if (_ribbon.AutoUpdateEnabled == false)
                return;

            //disable change
            ActiveWorksheet.Change -= OnChangeWorksheet;

            ActiveWorkbook.RefreshAll();

            //// find all pivot
            //Excel.PivotTables tables = ActiveWorkbook.PivotTables;
            //Debug.WriteLine(tables.Count);
            //foreach (Excel.PivotTable pivotTable in tables)
            //    pivotTable.RefreshTable(); //RefreshTable();

            //foreach (Excel.Worksheet worksheet in ActiveWorkbook.Sheets)
            //{
            //    var pivot = worksheet.PivotTables();
            //    if (pivot is Excel.PivotTables)
            //    {
            //        foreach (Excel.PivotTable pivotTable in (Excel.PivotTables) pivot)
            //        {
            //            pivotTable.RefreshTable();
            //        }
            //    }

            //}

            // enable change
            ActiveWorksheet.Change += OnChangeWorksheet;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // attach event
            ActiveWorkbook = Application.ActiveWorkbook;
            Application.WorkbookActivate += OnWorkbookActivate;
        }

        private void OnWorkbookActivate(Excel.Workbook wb)
        {
            ActiveWorkbook = wb;
        }

        private void OnWorksheetActivate(object worksheet)
        {
            ActiveWorksheet = worksheet as Excel.Worksheet;

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private Ribbon _ribbon;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new Ribbon();
            _ribbon.AutoUpdateEnabledChanged += (sender, b) =>
            {
                var properties = _activeWorkbook.CustomDocumentProperties as Office.DocumentProperties;
                if (properties != null)
                {
                    try
                    {
                        var autoupdate = properties[PropertyNamePivotAutoRefresh];
                        autoupdate.Value = b;
                    }
                    catch
                    {
                        properties.Add(PropertyNamePivotAutoRefresh, false, Office.MsoDocProperties.msoPropertyTypeBoolean, b);
                    }
                }
            };
            return _ribbon;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}

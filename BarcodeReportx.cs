using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using DevExpress.XtraReports.UI;
using System.Data;

namespace BarcodePrinting
{
    public partial class BarcodeReportx : DevExpress.XtraReports.UI.XtraReport
    {
        public BarcodeReportx()
        {
            InitializeComponent();
        }

        private void SCM010101TRPT_BeforePrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
        }
    }
}

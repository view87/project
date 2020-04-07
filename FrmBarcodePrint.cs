using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace BarcodePrinting
{
    public partial class FrmBarcodePrint : Form
    {
        string[] sheetNames;
        public DataTable f_kqyfInfo;
        public int isCancel
        {
            get;
            set;
        }

        string strFileName = string.Empty;
        DataTable dtBarcode = new DataTable();
        DataSet ds = new DataSet();

        public FrmBarcodePrint()
        {
            InitializeComponent();
            dtBarcode.TableName = "BarcodeTable";
            dtBarcode.Columns.Add(new DataColumn("barcode", typeof(string)));
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Microsoft Excel File (*.xls)|*.xls";
            open.RestoreDirectory = true;
            if (open.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = open.FileName;
                showSheetNames();
            }
        }

        private void showSheetNames()
        {
            cboSheetName.Items.Clear();
            sheetNames = UHelper.ExcelHelper.GetWorkSheetNames(textBox1.Text.Trim());
            for (int i = 0; i < sheetNames.Length; i++)
            {
                cboSheetName.Items.Add(sheetNames[i]);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text.Trim()))
            {
                MessageBox.Show("请选择有效的Excel文件!!!");

                return;
            }
            if (string.IsNullOrEmpty(cboSheetName.Text.Trim()))
            {
                MessageBox.Show("请选择工作表!!!");
                return;
            }

            f_kqyfInfo = UHelper.ExcelHelper.ExecuteDataTable(textBox1.Text.Trim(), cboSheetName.Text.Trim());

            gridControl1.DataSource = null;
            gridControl1.DataSource = f_kqyfInfo;
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                //黎马敦大码
                if (f_kqyfInfo != null && f_kqyfInfo.Rows.Count > 0)
                {
                    ds.Tables.Clear();
                    DataTable _dt = f_kqyfInfo.Copy();
                    for (int i = 0; i < _dt.Columns.Count; i++)
                    {
                        switch (_dt.Columns[i].ColumnName)
                        {
                            case "材料标准名称":
                                _dt.Columns[i].ColumnName = "changpinname";
                                break;
                            case "数量":
                                _dt.Columns[i].ColumnName = "num";
                                break;
                            case "规格":
                                _dt.Columns[i].ColumnName = "spec";
                                break;
                            case "净重":
                                _dt.Columns[i].ColumnName = "weigth";
                                break;
                            case "托盘号":
                                _dt.Columns[i].ColumnName = "tuopanhao";
                                break;
                            case "板号":
                                _dt.Columns[i].ColumnName = "banhao";
                                break;
                            case "包装日期":
                                _dt.Columns[i].ColumnName = "baozhuangriqi";
                                break;
                            case "内含质量批识别码号":
                                _dt.Columns[i].ColumnName = "shibiehao";
                                break;
                            case "内含质量批":
                                _dt.Columns[i].ColumnName = "zhiliangpi";
                                break;
                            case "供应商名称":
                                _dt.Columns[i].ColumnName = "gongyishang";
                                break;
                            case "执行标准号":
                                _dt.Columns[i].ColumnName = "biaozhunhao";
                                break;
                            case "条码号":
                                _dt.Columns[i].ColumnName = "barcode";
                                break;
                        }
                    }
                    _dt.TableName = "BarcodeTable";
                    ds.Tables.Add(_dt);

                    BarcodeReport rpt = new BarcodeReport();
                    rpt.DataMember = "BarcodeTable";
                    rpt.DataSource = ds;
                    this.Visible = false;
                    rpt.ShowPreviewDialog();
                    this.Visible = true;
                }
            }
            else
            {
                //虹之彩大码 
                ds.Tables.Clear();
                DataTable _dt = f_kqyfInfo.Copy();
                for (int i = 0; i < _dt.Columns.Count; i++)
                {
                    switch (_dt.Columns[i].ColumnName)
                    {
                        case "产品号":
                            _dt.Columns[i].ColumnName = "changpinhao";
                            break;
                        case "产品标识":
                            _dt.Columns[i].ColumnName = "changpinbiaoshi";
                            break;
                        case "产品编码":
                            _dt.Columns[i].ColumnName = "changpinbianma";
                            break;
                        case "计划单号":
                            _dt.Columns[i].ColumnName = "jihuadanhao";
                            break;
                        case "生产批":
                            _dt.Columns[i].ColumnName = "jihuapi";
                            break;
                        case "包装日期":
                            _dt.Columns[i].ColumnName = "date";
                            break;
                        case "材料标准名称":
                            _dt.Columns[i].ColumnName = "fname";
                            break;
                        case "内含质量批":
                            _dt.Columns[i].ColumnName = "zhiliangpi";
                            break;
                        case "执行标准":
                            _dt.Columns[i].ColumnName = "zhixingbiaozhun";
                            break;
                        case "规格":
                            _dt.Columns[i].ColumnName = "spec";
                            break;
                        case "数量":
                            _dt.Columns[i].ColumnName = "fnum";
                            break;
                        case "加工代码":
                            _dt.Columns[i].ColumnName = "jiagongdaima";
                            break;
                        case "净重":
                            _dt.Columns[i].ColumnName = "jingzhong";
                            break;
                        case "台板号":
                            _dt.Columns[i].ColumnName = "taibanhao";
                            break;
                        case "托盘号":
                            _dt.Columns[i].ColumnName = "tuopanhao";
                            break;
                        case "执行标准号":
                            _dt.Columns[i].ColumnName = "biaozhunhao";
                            break;
                        case "码包号":
                            _dt.Columns[i].ColumnName = "mabaohao";
                            break;
                        case "供应商地址":
                            _dt.Columns[i].ColumnName = "addr";
                            break;
                        case "一维码":
                            _dt.Columns[i].ColumnName = "barcode";
                            break;
                        case "警语":
                            _dt.Columns[i].ColumnName = "jingyu";
                            break;
                    }
                }
                _dt.TableName = "BarcodeTable";
                ds.Tables.Add(_dt);

                BarcodeReportd rpt = new BarcodeReportd();
                rpt.DataMember = "BarcodeTable";
                rpt.DataSource = ds;
                this.Visible = false;
                rpt.ShowPreviewDialog();
                this.Visible = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                //黎马敦小码
                if (f_kqyfInfo != null && f_kqyfInfo.Rows.Count > 0)
                {
                    ds.Tables.Clear();
                    DataTable _dt = f_kqyfInfo.Copy();
                    for (int i = 0; i < _dt.Columns.Count; i++)
                    {
                        switch (_dt.Columns[i].ColumnName)
                        {
                            case "材料标准名称":
                                _dt.Columns[i].ColumnName = "changpinname";
                                break;
                            case "数量":
                                _dt.Columns[i].ColumnName = "num";
                                break;
                            case "规格":
                                _dt.Columns[i].ColumnName = "spec";
                                break;
                            case "板号":
                                _dt.Columns[i].ColumnName = "banhao";
                                break;
                            case "印刷日期":
                                _dt.Columns[i].ColumnName = "yingshuariqi";
                                break;
                            case "质量批":
                                _dt.Columns[i].ColumnName = "zhiliangpi";
                                break;
                            case "质量批识别号":
                                _dt.Columns[i].ColumnName = "shibiehao";
                                break;
                            case "供应商名称":
                                _dt.Columns[i].ColumnName = "gongyishang";
                                break;
                            case "条码号":
                                _dt.Columns[i].ColumnName = "barcode";
                                break;
                        }
                    }
                    _dt.TableName = "BarcodeTable";
                    ds.Tables.Add(_dt);

                    BarcodeReport1 rpt = new BarcodeReport1();
                    rpt.DataMember = "BarcodeTable";
                    rpt.DataSource = ds;
                    this.Visible = false;
                    rpt.ShowPreviewDialog();
                    this.Visible = true;
                }
            }
            else
            {
                //虹之彩小码
                if (f_kqyfInfo != null && f_kqyfInfo.Rows.Count > 0)
                {
                    ds.Tables.Clear();
                    DataTable _dt = f_kqyfInfo.Copy();
                    for (int i = 0; i < _dt.Columns.Count; i++)
                    {
                        switch (_dt.Columns[i].ColumnName)
                        {
                            case "材料标准名称":
                                _dt.Columns[i].ColumnName = "changpinname";
                                break;
                            case "产品编码":
                                _dt.Columns[i].ColumnName = "wuliaocode";
                                break;
                            case "生产批":
                                _dt.Columns[i].ColumnName = "spec";
                                break;
                            case "数量":
                                _dt.Columns[i].ColumnName = "fnum";
                                break;
                            case "台板号":
                                _dt.Columns[i].ColumnName = "taibanhao";
                                break;
                            case "质量批":
                                _dt.Columns[i].ColumnName = "zhiliangpi";
                                break;
                            case "加工代码":
                                _dt.Columns[i].ColumnName = "jiagongdaima";
                                break;
                            case "成品检验":
                                _dt.Columns[i].ColumnName = "chengpinjianyan";
                                break;
                            case "挑选":
                                _dt.Columns[i].ColumnName = "tiaoxuan";
                                break;
                            case "生产单号":
                                _dt.Columns[i].ColumnName = "shengchandanhao";
                                break;
                            case "打包":
                                _dt.Columns[i].ColumnName = "dabao";
                                break;
                            case "日期":
                                _dt.Columns[i].ColumnName = "date";
                                break;
                            case "供应商名称":
                                _dt.Columns[i].ColumnName = "gongyishang";
                                break;
                            case "码包号":
                                _dt.Columns[i].ColumnName = "mabaohao";
                                break;
                            case "一维码":
                                _dt.Columns[i].ColumnName = "barcode";
                                break;
                        }
                    }
                    _dt.TableName = "BarcodeTable";
                    ds.Tables.Add(_dt);

                    BarcodeReportx rpt = new BarcodeReportx();
                    rpt.DataMember = "BarcodeTable";
                    rpt.DataSource = ds;
                    this.Visible = false;
                    rpt.ShowPreviewDialog();
                    this.Visible = true;
                }
            }
        }
    }
}

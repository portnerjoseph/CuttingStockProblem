using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CuttingProblemRevised
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            //this.dataGridView1.DataBindingComplete +=new DataGridViewBindingCompleteEventHandler(this.DataBindingComplete);
        }
        List<double> lengths;
        //List<CutList> solutions =new List<CutList>();
        //List<string> solutions = new List<string>();
        List<double> numbers;
        String contain = "";
        String textfile = "";
        String textfilenowaste = "";
        double totsize;
        private void btn_back_Click(object sender, EventArgs e)
        {
            switchpage();
        }
        private void switchpage()
        {
            panel1.Visible = !panel1.Visible;
            panel2.Visible = !panel2.Visible;
        }
        int numRows=0;
        List<string>[] lg4Output ;
        private object progressBar;
        private ProgressBar progressBar1;
        private int max;
        private delegate void SetProgressDelegate(int value);
        private void UpdateProgress(int value)
        {
            if (value < progressBar1.Minimum || value > progressBar1.Maximum)
            {
                return;
            }

            if (progressBar1.InvokeRequired)
            {
                // Execute the update on the main UI thread
                progressBar1.Invoke(new SetProgressDelegate(UpdateProgress), new object[] { value });
            }
            else
            {
                // Update the progress bar value
                progressBar1.Value = value;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            textfile = "";
            textfilenowaste = "";
            contain = "";
            numbers = new List<double>();
            lengths = new List<double>();
            try
            {
                totsize = Convert.ToDouble(textBox1.Text);
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    lengths.Add(Convert.ToDouble(row.Cells[0].Value));
                }
            }
            catch (FormatException c)
            {
                return;
            }
            lengths.RemoveAll(i => i == 0);
            lengths.Sort();
            lg4Output = new List<string>[lengths.Count];
            for (int x = 0; x < lg4Output.Length; x++)
                lg4Output[x] = new List<string>();
            recursive("", lengths, 0, -1);
            outputToTable(contain);
            exportToolStripMenuItem.Visible = true;
            lbl_numRows.Text = "There are " + numRows + "  cutting patterns.";
            switchpage();
        }
        private void recursive(String output, List<double> lengths, double sum, int count)
        {

            String temp = "";
            int x = 0;
            count++;
            if (count < lengths.Count)
            {
                temp = x + ",";
                recursive(output + temp, lengths, sum, count);
                while (sum <= totsize - lengths[count])
                {
                    x++;
                    sum += lengths[count];
                    temp = x + ",";
                    recursive(output + temp, lengths, sum, count);
                }
            }
            if (!contain.Contains(output))
            {
                contain += output + "\n";
                //.Add(output);
                numbers.Add(sum);
            }

        }
        private void outputToTable(string input)
        {
            int index = 0;
            numRows = 0;
            dataGridViewOutput.Rows.Clear();
            dataGridViewOutput.Columns.Clear();
            dataGridViewOutput.Refresh();
            double min = lengths.Min();
            List<string> line = new List<string>();
            String[] temp= input.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            for (int x = 0; x < lengths.Count; x++)
                dataGridViewOutput.Columns.Add("len" + lengths[x], "cut: " + lengths[x]);
            dataGridViewOutput.Columns.Add("waste", "remnant");
           // Array.Sort(numbers.ToArray(), temp);
            for (int x=0;x<temp.Length;x++)
            {
                if(totsize- numbers[x]<min)
                {
                    line = new List<string>(temp[x].Split(','));
                    line.RemoveAll(s => string.IsNullOrWhiteSpace(s));
                    addToLg4(line);
                    textfilenowaste += string.Join(", ", line) + "\r\n";
                    line.Add("" + (totsize - numbers[x]));
                    textfile += string.Join(", ", line) + "\r\n";
                   // lbl_output.Text += string.Join(",", line)+"\n";
                    dataGridViewOutput.Rows.Add(line.ToArray());
                    numRows++;
                }
            }
            //dataGridViewOutput.Sort(dataGridViewOutput.Columns["waste"], ListSortDirection.Ascending);
        }
        private void addToLg4 (List<string> temp)
        {
            for(int x=0;x<lg4Output.Length;x++)
                lg4Output[x].Add(temp[x]);
        }
        private double[] convertArr(List<string> templist)
        {
            double[] ans = new double[templist.Count];
            for(int x=0;x<templist.Count;x++)
            {
                ans[x] = Convert.ToDouble(templist[x]);
            }
            return ans;
        }

        private void DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            // Loops through each row in the DataGridView, and adds the
            // row number to the header
            foreach (DataGridViewRow dGVRow in this.dataGridView1.Rows)
            {
                dGVRow.HeaderCell.Value = String.Format("{0}", dGVRow.Index + 1);
            }

            // This resizes the width of the row headers to fit the numbers
            this.dataGridView1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        private void dataGridViewOutput_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var grid = sender as DataGridView;
            var rowIdx = "X"+(e.RowIndex + 1).ToString();
            var centerFormat = new StringFormat()
            {
                // right alignment might actually make more sense for numbers
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            var headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
            e.Graphics.DrawString(rowIdx, this.Font, SystemBrushes.ControlText, headerBounds, centerFormat);
        }

        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var grid = sender as DataGridView;
            var rowIdx = (e.RowIndex + 1).ToString();

            var centerFormat = new StringFormat()
            {
                // right alignment might actually make more sense for numbers
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            var headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
            e.Graphics.DrawString(rowIdx, this.Font, SystemBrushes.ControlText, headerBounds, centerFormat);
        }

        private void dataGridViewOutput_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
        {
            double a = double.Parse(e.CellValue1.ToString()), b = double.Parse(e.CellValue2.ToString());

            // If the cell value is already an integer, just cast it instead of parsing

            e.SortResult = a.CompareTo(b);

            e.Handled = true;
        }

        private void dataGridView1_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
        {
            double a = double.Parse(e.CellValue1.ToString()), b = double.Parse(e.CellValue2.ToString());

            // If the cell value is already an integer, just cast it instead of parsing

            e.SortResult = a.CompareTo(b);

            e.Handled = true;
        }


        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
                dataGridViewOutput.CurrentCell = null;

                dataGridViewOutput.Columns["waste"].Visible = !dataGridViewOutput.Columns["waste"].Visible;
        }

        private void asGridToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                sfd.FilterIndex = 1;
                sfd.DefaultExt = ".txt";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(sfd.FileName, textfile);
                }
            }

        }

        private void asGridNoWasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                sfd.FilterIndex = 1;
                sfd.DefaultExt = ".txt";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(sfd.FileName, textfilenowaste);
                }
            }
        }

        private void dataGridView2_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var grid = sender as DataGridView;
            var rowIdx = ""+lengths[e.RowIndex];

            var centerFormat = new StringFormat()
            {
                // right alignment might actually make more sense for numbers
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };
            var headerBounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
            e.Graphics.DrawString(rowIdx, this.Font, SystemBrushes.ControlText, headerBounds, centerFormat);
        }
        private void formatDataGridView()
        {
            dataGridView2.Rows.Clear();
            dataGridView2.Refresh();
            for(int x=0;x<lengths.Count;x++)
            {
                dataGridView2.Rows.Add();
            }
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel2.Visible = true;
            panel3.Visible = false;
        }

        private void asLg4ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
            panel2.Visible = false;
            panel3.Visible = true;
            formatDataGridView();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            String strLg4="";
            string[] xnums = new string[lg4Output[0].Count];
            List<string> inputList = new List<string>();
            foreach (DataGridViewRow dGVRow in this.dataGridView2.Rows)
                inputList.Add(dGVRow.Cells[0].Value.ToString().Trim());

            strLg4 += "MODEL: \r\nMin = (";
            for (int x = 0; x < xnums.Length; x++)
                xnums[x]= "X" + (x + 1);

            strLg4 += string.Join(" + ", xnums)+");\r\n";
            for(int x=0;x<lg4Output.Length; x++)
            {
                for(int y=0;y<lg4Output[x].Count-1;y++)
                    strLg4 += lg4Output[x][y] + "*" + xnums[y]+" + ";

                strLg4 += lg4Output[x].Last() + "*" + xnums.Last() + " > " + inputList[x]+";\r\n";
            }
            for(int x=0;x<xnums.Length;x++)
                strLg4 += "@GIN (" + xnums[x] + ");\r\n";

            strLg4 += "END";

            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "lg4 files (*.lg4)|*.lg4|All files (*.*)|*.*";
                sfd.FilterIndex = 1;
                sfd.DefaultExt = ".lg4";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(sfd.FileName, strLg4);
                }
            }

        }




        private void asExcelDocumentToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //makes sure remant is shown
            checkBox1.Checked = true;


            SaveFileDialog sfd = new SaveFileDialog();
             sfd.Filter = "Excel Documents (*.xls)|*.xls";
             sfd.FileName = "Cutting_Stock_Solution";
             if (sfd.ShowDialog() == DialogResult.OK)
             {
                 // Copy DataGridView results to clipboard
                 copyAlltoClipboard();

                 object misValue = System.Reflection.Missing.Value;
                 Excel.Application xlexcel = new Excel.Application();

                 xlexcel.DisplayAlerts = false; // Without this you will get two confirm overwrite prompts
                 Excel.Workbook xlWorkBook = xlexcel.Workbooks.Add(misValue);
                 Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


                 // Paste clipboard results to worksheet range
                 Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[2, 1];
                 CR.Select();
                 xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                 // For some reason column A is always blank in the worksheet. ¯\_(ツ)_/¯
                 // Delete blank column A and select cell A1
                 Excel.Range delRng = xlWorkSheet.get_Range("A:A").Cells;
                 delRng.Delete(Type.Missing);
                 xlWorkSheet.get_Range("A1").Select();


                for (int x = 0; x < lengths.Count; x++)
                    xlWorkSheet.Cells[1, (x + 1)] = "Cut: " + lengths[x];
                xlWorkSheet.Cells[1, lengths.Count+1] = "remnant";
                xlWorkSheet.Cells[1, 1].EntireRow.Font.Bold = true;




                // Save the excel file under the captured location from the SaveFileDialog
                xlWorkBook.SaveAs(sfd.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                 xlexcel.DisplayAlerts = true;
                 xlWorkBook.Close(true, misValue, misValue);
                 xlexcel.Quit();

                 releaseObject(xlWorkSheet);
                 releaseObject(xlWorkBook);
                 releaseObject(xlexcel);

                 // Clear Clipboard and DataGridView selection
                 Clipboard.Clear();
                 dataGridViewOutput.ClearSelection();

                 // Open the newly saved excel file
                 if (File.Exists(sfd.FileName))
                     System.Diagnostics.Process.Start(sfd.FileName);
             }
        }
        
        private void copyAlltoClipboard()
        {
            dataGridViewOutput.SelectAll();
            DataObject dataObj = dataGridViewOutput.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occurred while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void asLg4WithToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
    
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Merge
{
    public partial class Form1 : Form
    {
        private string workdir = @"C:\Users\King\Desktop\something";
        DataTable allClassMessage = new DataTable();
        DataTable allTeachersMessage = new DataTable();
        string pattern = @"[\u4E00-\u9FA5]{2,4}";
        string savePath;

        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog.Title = "查找输入文件";
            openFileDialog.InitialDirectory = workdir;   //@是取消转义字符的意思
            openFileDialog.Filter = "Excel(*.xls)|*.xls|Excelx(*.xlsx)|*.xlsx|All Files(*.*)|*.*";
            openFileDialog.FilterIndex = 2;
            openFileDialog.Multiselect = true;
            openFileDialog.RestoreDirectory = false;

            if (openFileDialog.ShowDialog() != DialogResult.OK)
                return;

            var files = openFileDialog.FileNames;

            files.ToList().ForEach(item =>
            {
                try
                {
                    IWorkbook wb = WorkbookFactory.Create(item);
                    ExcelToDataTable(wb, allClassMessage);
                    listBox1.Items.Add(System.IO.Path.GetFileNameWithoutExtension(item));
                    listBox1.Refresh();
                }
                catch
                {
                    var warning = String.Format("这个文件：{0}，以被占用", item);
                    MessageBox.Show(warning);
                }
            });

            //savePath = openFileDialog.
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog.FilterIndex = 2;
            openFileDialog.Title = "查找教研室文件";
            openFileDialog.InitialDirectory = workdir;   //@是取消转义字符的意思
            openFileDialog.Filter = "Excel(*.xls)|*.xls|Excelx(*.xlsx)|*.xlsx|All Files(*.*)|*.*";
            
            openFileDialog.RestoreDirectory = false;

            if (openFileDialog.ShowDialog() != DialogResult.OK)
                return;

            var files = openFileDialog.FileNames;

            files.ToList().ForEach(item =>
            {
                try
                {
                    IWorkbook wb = WorkbookFactory.Create(item);
                    ExcelToDataTable(wb, allTeachersMessage);
                    
                }
                catch
                {
                    var warning = String.Format("这个文件：{0}，以被占用", item);
                    MessageBox.Show(warning);
                }
            });
        }

        public void ExcelToDataTable(IWorkbook workbook, DataTable data)
        {
            ISheet sheet = workbook.GetSheetAt(0);
            if (sheet == null) return;
            //开始读取的行号
            int StartReadRow = 1;
            //获取sheet文件中的行数
            int RowLength = sheet.LastRowNum;
            //获取该行的列数(即该行的长度)
            int CellLength = sheet.GetRow(0).LastCellNum;
            if (data.Columns.Count < 2)
            {
                IRow columnNameRow = sheet.GetRow(0);
                //遍历读取
                for (int columnNameIndex = 0; columnNameIndex < CellLength; columnNameIndex++)
                {
                    //不为空，则读入
                    if (columnNameRow.GetCell(columnNameIndex) != null)
                    {
                        //获取该单元格的值
                        string cellValue = columnNameRow.GetCell(columnNameIndex).StringCellValue;
                        if (cellValue != null)
                        {
                            //为DataTable添加列名
                            data.Columns.Add(new DataColumn(cellValue));
                        }
                    }
                }
            }

            for (int RowIndex = StartReadRow; RowIndex < RowLength; RowIndex++)
            {
                //获取sheet表中对应下标的一行数据
                IRow currentRow = sheet.GetRow(RowIndex);   //RowIndex代表第RowIndex+1行

                if (currentRow == null) continue;  //表示当前行没有数据，则继续
                                                   //获取第Row行中的列数，即Row行中的长度
                                                   //int currentColumnLength = currentRow.LastCellNum;

                //创建DataTable的数据行
                DataRow dataRow = data.NewRow();
                //遍历读取数据
                for (int columnIndex = 0; columnIndex < CellLength; columnIndex++)
                {
                    //没有数据的单元格默认为空
                    if (currentRow.GetCell(columnIndex) != null)
                    {
                        var cur = currentRow.GetCell(columnIndex);
                        if (cur.CellType == CellType.Formula)
                        {
                            dataRow[columnIndex] = currentRow.GetCell(columnIndex).NumericCellValue;
                        }
                        else
                        {
                            dataRow[columnIndex] = currentRow.GetCell(columnIndex);
                        }

                    }
                }

                //去掉名字中的空格，进一步合并
                string temp = (dataRow["姓名"] == DBNull.Value) ? string.Empty : dataRow["姓名"].ToString().Trim();

                foreach (Match match in Regex.Matches(temp, pattern))
                {
                    dataRow["姓名"] = match.Value;
                    data.Rows.Add(dataRow.ItemArray);
                }

                //把DataTable的数据行添加到DataTable中
                
            }
            workbook.Close();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(allClassMessage.Columns.Count == 0 ||allTeachersMessage.Columns.Count ==0)
            {
                return;
            }
            var acm = from c in allClassMessage.AsEnumerable()
                      select c;
            var atm = from t in allTeachersMessage.AsEnumerable()
                      select t;
            var all = from c in acm
                      join t in atm on c.Field<string>("姓名") equals t.Field<string>("姓名")
                      select new { t, c };
            var DistinguishByLaboratory = from tc in all
                                               group tc by tc.t.Field<string>("所在系（或教研室）");

            var currentDirectory = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

            foreach (var Department in DistinguishByLaboratory)
            {
                string savePath = CreateDirectory(currentDirectory, Department.Key);
                var DistinguishByName = from tc in Department
                                        group tc by tc.t.Field<string>("姓名");
                foreach (var person in DistinguishByName)
                {
                    var dt = new DataTable();
                    dt = allClassMessage.Clone();
                    foreach (var tc in person)
                        dt.Rows.Add(tc.c.ItemArray);
                    DataTableToExcel(savePath + @"/" + person.Key + ".xlsx", dt, person.Key, true);
                }

            }

            var allc = from tc in all
                       select tc.c;

            var other = acm.Except(allc).GroupBy(item=>item.Field<string>("姓名"));
            savePath = CreateDirectory(currentDirectory, "其他");
            foreach (var person in other)
            {
                var dt = new DataTable();
                dt = allClassMessage.Clone();
                foreach (var tc in person)
                    dt.Rows.Add(tc.ItemArray);
                DataTableToExcel(savePath + @"/" + person.Key + ".xlsx", dt, person.Key, true);
            }



        }

        private string CreateDirectory(string currentDirectory,string key)
        {
            string path = currentDirectory + "/输出/" + key;
            DirectoryInfo di = new DirectoryInfo(path);
            if (di.Exists)
                di.Delete(true);
            di.Create();
            return path;
        }

        /// <summary>
        /// 把DataTable的数据写入到指定的excel文件中
        /// </summary>
        /// <param name="TargetFileNamePath">目标文件excel的路径</param>
        /// <param name="sourceData">要写入的数据</param>
        /// <param name="sheetName">excel表中的sheet的名称，可以根据情况自己起</param>
        /// <param name="IsWriteColumnName">是否写入DataTable的列名称</param>
        /// <returns>返回写入的行数</returns>
        public static int DataTableToExcel(string TargetFileNamePath, DataTable sourceData, string sheetName, bool IsWriteColumnName)
        {
            FileStream fs = null; ;
            //数据验证
            if (!File.Exists(TargetFileNamePath))
            {
                //excel文件的路径不存在
                fs = new FileStream(TargetFileNamePath, FileMode.Create, FileAccess.ReadWrite);
            }
            if (sourceData == null)
            {
                throw new ArgumentException("要写入的DataTable不能为空");
            }

            if (sheetName == null && sheetName.Length == 0)
            {
                throw new ArgumentException("excel中的sheet名称不能为空或者不能为空字符串");
            }



            //根据Excel文件的后缀名创建对应的workbook
            IWorkbook workbook = null;
            if (TargetFileNamePath.IndexOf(".xlsx") > 0)
            {  //2007版本的excel
                workbook = new XSSFWorkbook();
            }
            else if (TargetFileNamePath.IndexOf(".xls") > 0) //2003版本的excel
            {
                workbook = new HSSFWorkbook();
            }
            else
            {
                return -1;    //都不匹配或者传入的文件根本就不是excel文件，直接返回
            }



            //excel表的sheet名
            ISheet sheet = workbook.CreateSheet(sheetName);
            if (sheet == null) return -1;   //无法创建sheet，则直接返回


            //写入Excel的行数
            int WriteRowCount = 0;



            //指明需要写入列名，则写入DataTable的列名,第一行写入列名
            if (IsWriteColumnName)
            {
                //sheet表创建新的一行,即第一行
                IRow ColumnNameRow = sheet.CreateRow(0); //0下标代表第一行
                //进行写入DataTable的列名
                for (int colunmNameIndex = 0; colunmNameIndex < sourceData.Columns.Count; colunmNameIndex++)
                {
                    ColumnNameRow.CreateCell(colunmNameIndex).SetCellValue(sourceData.Columns[colunmNameIndex].ColumnName);
                }
                WriteRowCount++;
            }

            if (sourceData.Rows.Count >= 1)
            {
                //写入数据
                for (int row = 0; row < sourceData.Rows.Count; row++)
                {
                    //sheet表创建新的一行
                    IRow newRow = sheet.CreateRow(WriteRowCount);
                    for (int column = 0; column < sourceData.Columns.Count; column++)
                    {

                        newRow.CreateCell(column).SetCellValue(sourceData.Rows[row][column].ToString());

                    }

                    WriteRowCount++;  //写入下一行
                }
            }

            //写入到excel中
            //FileStream fs = new FileStream(TargetFileNamePath, FileMode.Open, FileAccess.Write);
            workbook.Write(fs);

            //fs.Flush();
            fs.Close();

            workbook.Close();
            return WriteRowCount;
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}

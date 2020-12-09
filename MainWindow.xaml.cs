using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using System.Windows.Controls.Primitives;
using NPOI.SS.Util;
using System.Diagnostics;
using Microsoft.CSharp;
using System.CodeDom.Compiler;
using System.Reflection;
using System.Text.RegularExpressions;
using NPOI.XSSF.UserModel;

namespace LabAutomationPhysicochemical
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
        /// <summary>
        /// 检测化合物名称合计
        /// </summary>
        List<KeyValuePair<string,string>> compoundsNameList = new List<KeyValuePair<string,string>>();
        /// <summary>
        /// 横表添加了平行样之后的分样
        /// </summary>
        List<List<string>> verticalSampleNameList = new List<List<string>>();
        /// <summary>
        /// 竖表添加了平行样之后的分样
        /// </summary>
        List<List<string>> horizontalSampleNameList = new List<List<string>>();

        /// <summary>
        /// 委托单号
        /// </summary>
        string ReportNo = string.Empty;

        //调整一个横表格的总列数
        int verticalSheetColumnCount = 12;
        //调整一个横表格的总列数
        int horizontalSheetColumnCount = 7;

        /// <summary>
        /// 生活饮用水标准的datatable
        /// </summary>
        DataSet FirstStanderDataSet = new DataSet();

        /// <summary>
        /// HJ84-2016的datatable
        /// </summary>
        DataSet SecondStanderDataSet = new DataSet();

        /// <summary>
        /// 初始化每个元素的datatset
        /// </summary>
        DataSet AutoLoadDataSet = new DataSet();

        public MainWindow()
		{
			InitializeComponent();
		}

		private void Window_Loaded(object sender,RoutedEventArgs e)
		{
			topScrollViewer.DragEnter += scDragEnter;
			topScrollViewer.Drop += scDrop;
			mainScrollViewer.DragEnter += scDragEnter;
			mainScrollViewer.Drop += scDrop;
            AutoLoad();
        }

        /// <summary>
        /// 初始化中文名和检出限
        /// </summary>
        private void AutoLoad()
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;
            string ExcelName = "AutoLoad.xlsx";
            string fullpath = System.IO.Path.Combine(path,ExcelName);
            IWorkbook workbook = null;
            if (File.Exists(fullpath))
            {
                using (FileStream fs = File.OpenRead(fullpath))
                {
                    // 2007版本
                    if (fullpath.Contains(".xlsx"))
                    {
                        workbook = new XSSFWorkbook(fs);
                    }
                    // 2003版本
                    else if (fullpath.Contains(".xls"))
                    {
                        workbook = new HSSFWorkbook(fs);
                    }
                    if (workbook != null)
                    {
                        for (int i = 0; i < workbook.NumberOfSheets; i++)
                        {
                            ISheet sheet = workbook.GetSheetAt(i);
                            DataTable dataTable = new DataTable();
                            dataTable.TableName = sheet.SheetName;
                            for (int j = 0; j <= sheet.LastRowNum; j++)
                            {
                                IRow row = sheet.GetRow(j);
                                if (j == 0)
                                {
                                    for (int k = 0; k < row.LastCellNum; k++)
                                    {
                                        ICell cell = row.GetCell(k);
                                        string value = cell.StringCellValue;
                                        dataTable.Columns.Add(value);
                                    }

                                }
                                else
                                {
                                    DataRow dr = dataTable.NewRow();
                                    dr.ItemArray = row.ToArray();
                                    dataTable.Rows.Add(dr);
                                }
                                
                            }
                            AutoLoadDataSet.Tables.Add(dataTable);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 拖动进入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void scDragEnter(object sender,DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Link;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        /// <summary>
        /// 拖动放下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void scDrop(object sender,DragEventArgs e)
        {
            //foreach(string str in e.Data.GetFormats())
            //{
            //	MessageBox.Show(str);
            //}
            ScrollViewer scrollViewer = sender as ScrollViewer;

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Link;

                string[] paths = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (int.Parse(scrollViewer.Tag.ToString()) == 0)
                {
                    //导入模板
                    //CreateTemplate(paths[0]);
                }
                else if (int.Parse(scrollViewer.Tag.ToString()) == 1)
                {
                    //创建数据结构
                    CreateExcel(paths[0]);
                }
            }
            e.Handled = true;
        }

        /// <summary>
        /// 全部清空,重置
        /// </summary>
        private void AllClear()
        {
            compoundsNameList.Clear();
            ReportNo = string.Empty;
            ReportNoLabel.Content = ReportNo;
            FirstStanderDataSet.Tables.Clear();
            SecondStanderDataSet.Tables.Clear();
            maingrid.Children.Clear();
        }

        /// <summary>
        /// 通过文本创造核心内容
        /// </summary>
        /// <param name="path"></param>
        private void CreateExcel(string path)
        {
            AllClear();
            IWorkbook workbook = null;
            TabControl tabControl = new TabControl();
            tabControl.Name = "tabControl";
            if (File.Exists(path))
            {
                using (FileStream fs = File.OpenRead(path))
                {
                    // 2007版本
                    if (path.Contains(".xlsx"))
                    {
                        workbook = new XSSFWorkbook(fs);
                    }
                    // 2003版本
                    else if (path.Contains(".xls"))
                    {
                        workbook = new HSSFWorkbook(fs);
                    }
                    if (workbook != null)
                    {
                        ISheet sheet = workbook.GetSheetAt(0);//读取第一个sheet，当然也可以循环读取每个sheet
                        IRow firstRow = sheet.GetRow(0);
                        int Num = 0;
                        for (int i = 0; i < firstRow.LastCellNum; i++)
                        {
                            ICell cell = firstRow.GetCell(i);
                            if (cell != null)
                            {
                                if (i == 0)
                                {
                                    ReportNo = cell.StringCellValue;
                                }
                                else
                                {
                                    if (cell.StringCellValue != string.Empty && cell.StringCellValue != "")
                                    {
                                        string compoundName = cell.StringCellValue.Trim();
                                        DataTable dataTable = new DataTable();
                                        dataTable.TableName = compoundName;
                                        CreateDataTable(tabControl,sheet,dataTable,cell.ColumnIndex,Num);
                                        Num++;
                                    }
                                }
                            }
                        }
                    }

                    //AddParallelSamplesToList();

                    maingrid.Children.Add(tabControl);
                    ReportNoLabel.Content = ReportNo;
                }
            }
        }

        /// <summary>
        /// 生成检出datatable
        /// </summary>
        /// <param name="tabControl"></param>
        /// <param name="sheet"></param>
        /// <param name="dataTable"></param>
        /// <param name="compoundsNum"></param>
        /// <param name="num"></param>
        private void CreateDataTable(TabControl tabControl,ISheet sheet,DataTable dataTable,int compoundsNum,int num)
        {
            int rowCount = sheet.LastRowNum;//总行数
            //要转中文名
            string modelC = string.Empty;
            if (testZDRadioButton.IsChecked == true)
            {
                DataTable dt = AutoLoadDataSet.Tables[0];
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string value = dt.Rows[i][0].ToString();
                    if (dataTable.TableName == value)
                    {
                        dataTable.TableName = dt.Rows[i][1].ToString();
                        modelC = dt.Rows[i][2].ToString();
                    }
                }
                FirstStanderDataSet.Tables.Add(dataTable);
            }
            else if (testJCRadioButton.IsChecked == true)
            {
                DataTable dt = AutoLoadDataSet.Tables[1];
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string value = dt.Rows[i][0].ToString();
                    if (dataTable.TableName == value)
                    {
                        dataTable.TableName = dt.Rows[i][1].ToString();
                        modelC = dt.Rows[i][2].ToString();
                    }
                }
                SecondStanderDataSet.Tables.Add(dataTable);
            }
            for (int i = 1; i <= rowCount; i++)
            {
                IRow row = sheet.GetRow(i);
                //由于Excel在非数据区进行了格式设置，那么sheet.LastRowNum 得到的值就会与实际得到的值不符。从而因有非空验证，造成导入失败。
                //所以直接先判断第一个单元格是否为空，在进行后面的操作
                ICell firstCell = row.GetCell(0);
                if (firstCell != null)
                {
                    //第二行都是表头，要组成datatable
                    if (i == 1)
                    {
                        for (int j = 0; j < 2; j++)
                        {
                            ICell secondCell = row.GetCell(j);
                            if (secondCell != null)
                            {
                                if (secondCell.StringCellValue != string.Empty && secondCell.StringCellValue != "")
                                {
                                    dataTable.Columns.Add(secondCell.StringCellValue);
                                }
                            }
                        }
                        dataTable.Columns.Add("样品量");
                    }
                    //第三行开始是数据
                    else
                    {
                        DataRow dataRow = dataTable.NewRow();
                        for (int k = 0; k < 2; k++)
                        {
                            ICell cell = row.GetCell(k);
                            if (cell != null)
                            {
                                if (cell.CellType == CellType.Numeric)
                                {
                                    dataRow[k] = cell.NumericCellValue;
                                }
                                else
                                {
                                    dataRow[k] = cell.StringCellValue.Trim();
                                }
                            }
                        }
                        ICell newCell = row.GetCell(compoundsNum);
                        if (newCell.CellType == CellType.Numeric)
                        {
                            dataRow[dataRow.ItemArray.Length - 1] = newCell.NumericCellValue;
                        }
                        else
                        {
                            dataRow[dataRow.ItemArray.Length - 1] = newCell.StringCellValue;
                        }
                        dataTable.Rows.Add(dataRow);
                    }
                }
            }
            dataTable.Columns.Add("C");

            for (int i = dataTable.Rows.Count - 1; i >= 0; i--)
            {
                string value = dataTable.Rows[i][0].ToString();
                if (value.Contains("BQLH"))
                {
                    dataTable.Rows.RemoveAt(i);
                }
            }

            AddParallelSamplesToList(dataTable);

            TabItem tabItem = new TabItem();
            //tabItem.Header = name[1] + " | " + name[2];
            StackPanel stackPanel = CreateStackPanel(dataTable.TableName,modelC,num);
            tabItem.Header = stackPanel;
            DataGrid dg = new DataGrid();
            dg.Name = "dataGrid";
            dg.ItemsSource = dataTable.DefaultView;
            dg.CanUserSortColumns = true;
            dg.CanUserReorderColumns = true;
            dg.IsTabStop = true;
            tabItem.Content = dg;
            tabControl.Items.Add(tabItem);
        }

        /// <summary>
        /// 创建tabheader用的stackpanel
        /// </summary>
        /// <param name="compoundsName"></param>
        /// <param name="modelC"></param>
        /// <param name="num"></param>
        /// <returns></returns>
        private StackPanel CreateStackPanel(string compoundsName,string modelC,int num)
        {
            StackPanel stackPanel = new StackPanel();
            stackPanel.Orientation = Orientation.Horizontal;
            stackPanel.VerticalAlignment = System.Windows.VerticalAlignment.Center;
            stackPanel.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;

            Label numLabel = new Label();
            numLabel.Content = (num + 1).ToString() + ".";
            numLabel.HorizontalContentAlignment = System.Windows.HorizontalAlignment.Center;
            numLabel.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;

            Label compoundslabel = new Label();
            compoundslabel.Content = compoundsName;
            compoundslabel.HorizontalContentAlignment = System.Windows.HorizontalAlignment.Center;
            compoundslabel.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
            TextBox textBox = new TextBox();
            textBox.Width = 50;
            textBox.HorizontalAlignment = System.Windows.HorizontalAlignment.Center;
            textBox.VerticalAlignment = System.Windows.VerticalAlignment.Center;
            textBox.TabIndex = num;
            textBox.KeyUp += Tab_TextBox_KeyUp;
            textBox.Text = modelC;


            stackPanel.Children.Add(numLabel);
            stackPanel.Children.Add(compoundslabel);
            stackPanel.Children.Add(textBox);

            return stackPanel;
        }

        /// <summary>
        /// enter切换检出限
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Tab_TextBox_KeyUp(object sender,KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                TextBox textbox = sender as TextBox;
                StackPanel stackPanel = textbox.Parent as StackPanel;
                TabItem tabItem = stackPanel.Parent as TabItem;
                TabControl tabControl = tabItem.Parent as TabControl;
                int tabNum = tabControl.Items.IndexOf(tabItem);
                //到达最大值
                TabItem nextTabItem;
                if (tabNum == tabControl.Items.Count - 1)
                {
                    nextTabItem = tabControl.Items[0] as TabItem;
                }
                else
                {
                    nextTabItem = tabControl.Items[tabNum + 1] as TabItem;
                }
                StackPanel nextStackPanel = nextTabItem.Header as StackPanel;
                foreach (var item in nextStackPanel.Children)
                {
                    if (item.GetType() == typeof(TextBox))
                    {
                        TextBox nextTextBox = item as TextBox;
                        Keyboard.Focus(nextTextBox);
                        nextTextBox.Focus();
                    }
                }

            }
        }

        /// <summary>
        /// 添加平行样
        /// </summary>
        /// <param name="dataTable"></param>
        private void AddParallelSamplesToList(DataTable dataTable)
        {
            List<string> sampleNameListV = new List<string>();
            List<string> sampleNameListH = new List<string>();
            string Ebanlance = "Dup";
            string Cbanlance = "平均值";
            int importTakeNumV = 17;
            int importTakeNumH = 23;
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                string name = dataTable.Rows[i]["目标化合物"].ToString();
                string value = dataTable.Rows[i]["样品量"].ToString();
                if (name.Contains("Dup"))
                {
                    DataRow dataRow = dataTable.NewRow();
                    dataRow[0] = name.Replace(Ebanlance,Cbanlance);
                    dataRow[1] = "/";
                    dataRow[2] = "/";
                    dataTable.Rows.InsertAt(dataRow,i + 1);
                }
                if (value.Contains("n.a."))
                {
                    dataTable.Rows[i]["样品量"] = 0;
                }
                sampleNameListV.Add(name);
            }
            sampleNameListV.Add("以下空白");
            sampleNameListH = sampleNameListV.ToList();
            //横表
            if (verticalSampleNameList.Count == 0)
            {
                int Count = sampleNameListV.Count % importTakeNumV > 0 ? sampleNameListV.Count / importTakeNumV + 1 : sampleNameListV.Count / importTakeNumV;
                for (int i = 0; i < Count; i++)
                {
                    if (i == Count - 1)
                    {
                        List<string> cellList = sampleNameListV.ToList();
                        verticalSampleNameList.Add(cellList);
                    }
                    else
                    {
                        List<string> cellList = sampleNameListV.Take(importTakeNumV).ToList();
                        sampleNameListV.RemoveRange(0,importTakeNumV);
                        verticalSampleNameList.Add(cellList);
                    }
                }
            }
            //竖表
            if (horizontalSampleNameList.Count == 0)
            {
                int Count = sampleNameListH.Count % importTakeNumH > 0 ? sampleNameListH.Count / importTakeNumH + 1 : sampleNameListH.Count / importTakeNumH;
                for (int i = 0; i < Count; i++)
                {
                    if (i == Count - 1)
                    {
                        List<string> cellList = sampleNameListH.ToList();
                        horizontalSampleNameList.Add(cellList);
                    }
                    else
                    {
                        List<string> cellList = sampleNameListH.Take(importTakeNumH).ToList();
                        sampleNameListH.RemoveRange(0,importTakeNumH);
                        horizontalSampleNameList.Add(cellList);
                    }
                }
            }
        }


        /// <summary>
        /// 导出生成Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void importExcel_Click(object sender,RoutedEventArgs e)
        {
            if (FirstStanderDataSet.Tables.Count == 0 && SecondStanderDataSet.Tables.Count == 0)
            {
                return;
            }
            //判断化合物是否大于2，从而分割成横表或者竖表
            if (compoundsNameList.Count > 4)
            {
                CreateMutiVerticalExcel();
            }
            //有机组横表
            else if (compoundsNameList.Count > 2)
            {
                CreateVerticalExcel();
            }
            else
            {
                CreateHorizontalExcel();
            }
        }

        /// <summary>
        /// 生成compoundsNameList
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void importExcel_MouseMove(object sender,MouseEventArgs e)
        {
            if (FirstStanderDataSet.Tables.Count == 0 && SecondStanderDataSet.Tables.Count == 0)
            {
                return;
            }
            compoundsNameList.Clear();
            TabControl tabControl = maingrid.Children[0] as TabControl;
            foreach (TabItem tabItem in tabControl.Items)
            {
                string compoundsName = string.Empty;
                string modelC = string.Empty;
                StackPanel stackPanel = tabItem.Header as StackPanel;
                foreach (var item in stackPanel.Children)
                {
                    if (item.GetType() == typeof(Label))
                    {
                        compoundsName = (item as Label).Content.ToString();
                    }
                    else if (item.GetType() == typeof(TextBox))
                    {
                        if ((item as TextBox).Text != null && (item as TextBox).Text != "" && (item as TextBox).Text != string.Empty)
                        {

                            modelC = (item as TextBox).Text;
                        }
                    }
                }
                KeyValuePair<string,string> keyValuePair = new KeyValuePair<string,string>(compoundsName,modelC);
                compoundsNameList.Add(keyValuePair);
            }
        }

        /// <summary>
        /// 创建原版横表
        /// </summary>
        private void CreateMutiVerticalExcel()
        {
            var workbook = new HSSFWorkbook();
            //切割化合物
            List<List<KeyValuePair<string,string>>> renameList = new List<List<KeyValuePair<string,string>>>();
            List<KeyValuePair<string,string>> cnList = compoundsNameList.ToList();
            int importTakeNum = 4;
            int importCount = compoundsNameList.Count % importTakeNum > 0 ? compoundsNameList.Count / importTakeNum + 1 : compoundsNameList.Count / importTakeNum;
            for (int i = 0; i < importCount; i++)
            {
                if (i == importCount - 1)
                {
                    List<KeyValuePair<string,string>> cellList = cnList.ToList();
                    renameList.Add(cellList);
                }
                else
                {
                    List<KeyValuePair<string,string>> cellList = cnList.Take(importTakeNum).ToList();
                    cnList.RemoveRange(0,importTakeNum);
                    renameList.Add(cellList);
                }
            }
            for (int j = 0; j < renameList.Count; j++)
            {
                var sheet = workbook.CreateSheet(j.ToString() + "-横-水");
                sheet.ForceFormulaRecalculation = true;
                int Count = 0;
                foreach (List<string> list in verticalSampleNameList)
                {
                    CreateVerticalSheet(sheet,list,Count,renameList[j]);
                    Count++;
                }

                // 自动调整列距
                for (int i = 0; i < Count * verticalSheetColumnCount; i++)
                {
                    if (i % verticalSheetColumnCount == 0)
                    {
                        sheet.SetColumnWidth(i,40 * 256);
                    }
                    else if ((i + 1) % verticalSheetColumnCount == 0)
                    {
                        sheet.SetColumnWidth(i,20 * 256);
                    }
                    else
                    {
                        sheet.AutoSizeColumn(i);
                    }

                    if (sheet.GetColumnWidth(i) < 10 * 256)
                    {
                        sheet.SetColumnWidth(i,10 * 256);
                    }
                }
            }

            ExportToExcel(workbook);
        }

        /// <summary>
        /// 创建竖表Excel
        /// </summary>
        private void CreateHorizontalExcel()
        {
            var workbook = new HSSFWorkbook();
            var sheet = workbook.CreateSheet("竖-水");
            sheet.ForceFormulaRecalculation = true;
            int Count = 0;
            foreach (List<string> list in horizontalSampleNameList)
            {
                CreateHorizontalSheet(sheet,list,Count);
                Count++;
            }

            // 自动调整列距
            for (int i = 0; i < Count * horizontalSheetColumnCount; i++)
            {
                if (i % horizontalSheetColumnCount == 0)
                {
                    sheet.SetColumnWidth(i,40 * 256);
                }
                else if ((i + 1) % horizontalSheetColumnCount == 0)
                {
                    sheet.SetColumnWidth(i,20 * 256);
                }
                else
                {
                    sheet.AutoSizeColumn(i);
                }
                if (sheet.GetColumnWidth(i) < 10 * 256)
                {
                    sheet.SetColumnWidth(i,10 * 256);
                }
            }

            ExportToExcel(workbook);
        }

        /// <summary>
        /// 创建横表Excel
        /// </summary>
        private void CreateVerticalExcel()
        {
            var workbook = new HSSFWorkbook();
            var sheet = workbook.CreateSheet("横-水");
            sheet.ForceFormulaRecalculation = true;
            int Count = 0;
            foreach (List<string> list in verticalSampleNameList)
            {
                CreateVerticalSheet(sheet,list,Count,compoundsNameList);
                Count++;
            }

            // 自动调整列距
            for (int i = 0; i < Count * verticalSheetColumnCount; i++)
            {
                if (i % verticalSheetColumnCount == 0)
                {
                    sheet.SetColumnWidth(i,40 * 256);
                }
                else if ((i + 1) % verticalSheetColumnCount == 0)
                {
                    sheet.SetColumnWidth(i,20 * 256);
                }
                else
                {
                    sheet.AutoSizeColumn(i);
                }
                if (sheet.GetColumnWidth(i) < 10 * 256)
                {
                    sheet.SetColumnWidth(i,10 * 256);
                }
            }

            ExportToExcel(workbook);
        }

        /// <summary>
        /// 创建竖表Excel
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="cellList"></param>
        /// <param name="Count"></param>
        private void CreateHorizontalSheet(ISheet sheet,List<string> cellList,int Count)
        {
            DataSet dataSet = new DataSet();
            if (testZDRadioButton.IsChecked == true)
            {
                dataSet = FirstStanderDataSet;
            }
            else if (testJCRadioButton.IsChecked == true)
            {
                dataSet = SecondStanderDataSet;
            }
            HSSFWorkbook workbook = sheet.Workbook as HSSFWorkbook;
            //设置顶部大标题样式
            HSSFCellStyle cellStyle = CreateStyle(workbook);
            HSSFCellStyle bordercellStyle = CreateStyle(workbook);
            //HSSFCellStyle cellGreyStyle = CreateGreyStyle(workbook);
            //HSSFCellStyle cellGreenStyle = CreateGreenStyle(workbook);
            //HSSFCellStyle cellRedStyle = CreateRedStyle(workbook);
            //HSSFCellStyle cellYellowStyle = CreateYellowStyle(workbook);
            bordercellStyle.BorderLeft = BorderStyle.Thin;
            bordercellStyle.BorderTop = BorderStyle.Thin;
            bordercellStyle.BorderLeft = BorderStyle.Thin;
            bordercellStyle.BorderRight = BorderStyle.Thin;
            //前五行 大表头
            for (int i = 0; i < 5; i++)
            {
                //第一行最右显示委托单号
                HSSFRow row = (Count == 0) ? (HSSFRow)sheet.CreateRow(i) : (HSSFRow)sheet.GetRow(i); //创建行或者获取行
                row.HeightInPoints = 30;
                switch (i)
                {
                    case 0:
                        {
                            for (int j = horizontalSheetColumnCount * Count; j < horizontalSheetColumnCount * Count + horizontalSheetColumnCount; j++)
                            {
                                var cell = row.CreateCell(j);
                                cell.CellStyle = bordercellStyle;
                                if (j == horizontalSheetColumnCount * Count)
                                {
                                    cell.SetCellValue("计算公式：");
                                    CellRangeAddress region = new CellRangeAddress(i,i,horizontalSheetColumnCount * Count,horizontalSheetColumnCount * Count + 1);
                                    sheet.AddMergedRegion(region);
                                }
                                else if (j == horizontalSheetColumnCount * Count + 2)
                                {
                                    cell.SetCellValue("目标化合物");
                                    CellRangeAddress nameregion = new CellRangeAddress(i,i,horizontalSheetColumnCount * Count + 2,horizontalSheetColumnCount * Count + 5);
                                    sheet.AddMergedRegion(nameregion);
                                }
                                else if (j == horizontalSheetColumnCount * Count + 6)
                                {
                                    cell.SetCellValue("备注");
                                    CellRangeAddress namesregion = new CellRangeAddress(i,i + 4,horizontalSheetColumnCount * Count + 6,horizontalSheetColumnCount * Count + 6);
                                    sheet.AddMergedRegion(namesregion);
                                }
                            }
                            break;
                        }
                    case 1:
                        {

                            for (int j = horizontalSheetColumnCount * Count; j < horizontalSheetColumnCount * Count + horizontalSheetColumnCount; j++)
                            {
                                var cell = row.CreateCell(j);
                                cell.CellStyle = bordercellStyle;

                                if (j == horizontalSheetColumnCount * Count + 6)
                                {
                                    cell.SetCellValue(string.Empty);
                                }
                                else if (j > horizontalSheetColumnCount * Count + 3 && j < horizontalSheetColumnCount * Count + 4 + compoundsNameList.Count)
                                {
                                    cell.SetCellValue(compoundsNameList[j - horizontalSheetColumnCount * Count - 4].Key);
                                }
                                else if (j > horizontalSheetColumnCount * Count + 1 && j < horizontalSheetColumnCount * Count + 2 + compoundsNameList.Count)
                                {
                                    cell.SetCellValue(compoundsNameList[j - horizontalSheetColumnCount * Count - 2].Key);
                                }
                                else if (j == horizontalSheetColumnCount * Count)
                                {
                                    CellRangeAddress region = new CellRangeAddress(i,i + 2,j,j + 1);
                                    sheet.AddMergedRegion(region);
                                    //要和公式那一块绑定在一起
                                    string wdm = FormulaComboBox.Text + "\n"
                                    + "C——样品中待测离子浓度，" + ZDJCCompanyComboBox.Text + "\n"
                                    + "Ci——查得样品中待测离子的浓度，" + TargetCompanyComboBox.Text + "\n"
                                    + "f——样品稀释倍数。";
                                    HSSFRichTextString rtsWdm = new HSSFRichTextString(wdm);
                                    cell.SetCellValue(rtsWdm);
                                }
                            }
                            break;
                        }
                    case 2:
                        {
                            for (int j = horizontalSheetColumnCount * Count; j < horizontalSheetColumnCount * Count + horizontalSheetColumnCount; j++)
                            {
                                var cell = row.CreateCell(j);
                                cell.CellStyle = bordercellStyle;
                                if (j == horizontalSheetColumnCount * Count + 2)
                                {
                                    cell.SetCellValue("/");
                                    CellRangeAddress firstregion = new CellRangeAddress(i,i,horizontalSheetColumnCount * Count + 2,horizontalSheetColumnCount * Count + 3);
                                    sheet.AddMergedRegion(firstregion);
                                }
                                else if (j == horizontalSheetColumnCount * Count + 4)
                                {
                                    string value = string.Empty;
                                    if (testJCRadioButton.IsChecked == true)
                                    {
                                        value = "√\t" + testJCRadioButton.Content + "\t□\t" + testZDRadioButton.Content + "(" + ZDJCCompanyComboBox.Text + ")";
                                    }
                                    else if (testZDRadioButton.IsChecked == true)
                                    {
                                        value = "□\t" + testJCRadioButton.Content + "\t√\t" + testZDRadioButton.Content + "(" + ZDJCCompanyComboBox.Text + ")";
                                    }
                                    cell.SetCellValue(value);
                                    CellRangeAddress secondregion = new CellRangeAddress(i,i,horizontalSheetColumnCount * Count + 4,horizontalSheetColumnCount * Count + 5);
                                    sheet.AddMergedRegion(secondregion);
                                }
                            }
                            break;
                        }
                    case 3:
                        {
                            for (int j = horizontalSheetColumnCount * Count; j < horizontalSheetColumnCount * Count + horizontalSheetColumnCount; j++)
                            {
                                var cell = row.CreateCell(j);
                                cell.CellStyle = bordercellStyle;
                                if (j == horizontalSheetColumnCount * Count + 2)
                                {
                                    cell.SetCellValue("/");
                                    CellRangeAddress firstregion = new CellRangeAddress(i,i,horizontalSheetColumnCount * Count + 2,horizontalSheetColumnCount * Count + 3);
                                    sheet.AddMergedRegion(firstregion);
                                }
                                else if (j > horizontalSheetColumnCount * Count + 3 && j < horizontalSheetColumnCount * Count + 4 + compoundsNameList.Count)
                                {
                                    cell.SetCellValue(compoundsNameList[j - horizontalSheetColumnCount * Count - 4].Value);
                                }
                            }
                            break;
                        }
                    case 4:
                        {
                            for (int j = horizontalSheetColumnCount * Count; j < horizontalSheetColumnCount * Count + horizontalSheetColumnCount; j++)
                            {
                                var cell = row.CreateCell(j);
                                cell.CellStyle = bordercellStyle;
                                if (j == horizontalSheetColumnCount * Count)
                                {
                                    cell.SetCellValue("样品编号");
                                }
                                else if (j == horizontalSheetColumnCount * Count + 1)
                                {
                                    cell.SetCellValue("稀释倍数f");
                                }
                                else if (j == horizontalSheetColumnCount * Count + 2)
                                {
                                    cell.SetCellValue("目标化合物测定值 M  (mg/L)");
                                    CellRangeAddress thirdregion = new CellRangeAddress(i,i,horizontalSheetColumnCount * Count + 2,horizontalSheetColumnCount * Count + 3);
                                    sheet.AddMergedRegion(thirdregion);
                                }
                                else if (j == horizontalSheetColumnCount * Count + 4)
                                {
                                    cell.SetCellValue("目标化合物浓度 C (mg/L)");
                                    CellRangeAddress fourthregion = new CellRangeAddress(i,i,horizontalSheetColumnCount * Count + 4,horizontalSheetColumnCount * Count + 5);
                                    sheet.AddMergedRegion(fourthregion);
                                }
                            }
                            break;
                        }
                }
            }

            //正式数据录入
            for (int k = 0; k < cellList.Count; k++)
            {
                HSSFRow sampleRow = (Count == 0) ? (HSSFRow)sheet.CreateRow(5 + k) : (HSSFRow)sheet.GetRow(5 + k); //创建行或者获取行
                sampleRow.HeightInPoints = 20;
                string sampleName = cellList[k];
                string sampleDilutionratio = string.Empty;
                List<string> sampleSize = new List<string>();
                for (int p = 0; p < dataSet.Tables.Count; p++)
                {
                    DataTable dataTable = dataSet.Tables[p];
                    for (int r = 0; r < compoundsNameList.Count; r++)
                    {
                        if (dataTable.TableName == compoundsNameList[r].Key)
                        {
                            for (int o = 0; o < dataTable.Rows.Count; o++)
                            {
                                if (dataTable.Rows[o][0].ToString() == sampleName)
                                {
                                    if (sampleDilutionratio == string.Empty)
                                    {
                                        sampleDilutionratio = dataTable.Rows[o][1].ToString();
                                    }
                                    sampleSize.Add(dataTable.Rows[o][2].ToString());
                                    break;
                                }
                            }
                        }
                    }
                }

                for (int l = horizontalSheetColumnCount * Count; l < horizontalSheetColumnCount * Count + horizontalSheetColumnCount; l++)
                {
                    var sampleCell = sampleRow.CreateCell(l);
                    sampleCell.CellStyle = bordercellStyle;
                    if (l == horizontalSheetColumnCount * Count)
                    {
                        sampleCell.SetCellValue(sampleName);
                    }
                    else if (l == horizontalSheetColumnCount * Count + 1)
                    {
                        sampleCell.SetCellValue(sampleDilutionratio);
                    }
                    else if (l > horizontalSheetColumnCount * Count + 1 && l < horizontalSheetColumnCount * Count + 2 + compoundsNameList.Count)
                    {
                        if (sampleName.Contains("以下空白"))
                        {
                            sampleCell.SetCellValue(string.Empty);
                        }
                        else
                        {
                            string value = sampleSize[l - horizontalSheetColumnCount * Count - 2];
                            if (!value.Contains("/"))
                            {
                                decimal C1 = decimal.Parse(value);
                                C1 = Math.Round(C1,4,MidpointRounding.ToEven);
                                value = CalculateAccuracyCX(C1.ToString(),4);
                            }
                            sampleCell.SetCellValue(value);
                        }
                    }
                    else if (l > horizontalSheetColumnCount * Count + 3 && l < horizontalSheetColumnCount * Count + 4 + compoundsNameList.Count)
                    {

                        string value = string.Empty;
                        DataTable dataTable = dataSet.Tables[l - horizontalSheetColumnCount * Count - 4];
                        if (sampleName.Contains("平均值"))
                        {
                            for (int o = 0; o < dataTable.Rows.Count; o++)
                            {
                                if (dataTable.Rows[o][0].ToString() == sampleName)
                                {
                                    decimal C1 = decimal.Parse(dataTable.Rows[o - 1][3].ToString());
                                    decimal C2 = decimal.Parse(dataTable.Rows[o - 2][3].ToString());
                                    value = CompareCompoundWithFormulaAverage(compoundsNameList[l - horizontalSheetColumnCount * Count - 4].Key,C1,C2);
                                    break;
                                }
                            }
                        }
                        else if (!sampleName.Contains("以下空白"))
                        {
                            value = CompareCompoundWithFormula(dataSet,compoundsNameList[l - horizontalSheetColumnCount * Count - 4].Key,sampleName,sampleDilutionratio,sampleSize[l - horizontalSheetColumnCount * Count - 4]);
                        }
                        sampleCell.SetCellValue(value);
                    }
                    else if (l == horizontalSheetColumnCount * Count + horizontalSheetColumnCount - 1)
                    {
                        sampleCell.SetCellValue("/");
                    }
                }
            }
        }

        /// <summary>
        /// 创建横表Excel
        /// </summary>
        private void CreateVerticalSheet(ISheet sheet,List<string> cellList,int Count,List<KeyValuePair<string,string>> CompoundsNameList)
        {
            DataSet dataSet = new DataSet();
            if (testZDRadioButton.IsChecked == true)
            {
                dataSet = FirstStanderDataSet;
            }
            else if (testJCRadioButton.IsChecked == true)
            {
                dataSet = SecondStanderDataSet;
            }
            HSSFWorkbook workbook = sheet.Workbook as HSSFWorkbook;
            //设置顶部大标题样式
            HSSFCellStyle cellStyle = CreateStyle(workbook);
            HSSFCellStyle bordercellStyle = CreateStyle(workbook);
            //HSSFCellStyle cellGreyStyle = CreateGreyStyle(workbook);
            //HSSFCellStyle cellGreenStyle = CreateGreenStyle(workbook);
            //HSSFCellStyle cellRedStyle = CreateRedStyle(workbook);
            //HSSFCellStyle cellYellowStyle = CreateYellowStyle(workbook);
            bordercellStyle.BorderLeft = BorderStyle.Thin;
            bordercellStyle.BorderTop = BorderStyle.Thin;
            bordercellStyle.BorderLeft = BorderStyle.Thin;
            bordercellStyle.BorderRight = BorderStyle.Thin;
            //前五行 大表头
            for (int i = 0; i < 5; i++)
            {
                //第一行最右显示委托单号
                HSSFRow row = (Count == 0) ? (HSSFRow)sheet.CreateRow(i) : (HSSFRow)sheet.GetRow(i); //创建行或者获取行
                row.HeightInPoints = 30;
                switch (i)
                {
                    case 0:
                        {
                            for (int j = verticalSheetColumnCount * Count; j < verticalSheetColumnCount * Count + verticalSheetColumnCount; j++)
                            {
                                var cell = row.CreateCell(j);
                                cell.CellStyle = bordercellStyle;
                                if (j == verticalSheetColumnCount * Count)
                                {
                                    cell.SetCellValue("计算公式：");
                                    CellRangeAddress region = new CellRangeAddress(i,i,verticalSheetColumnCount * Count,verticalSheetColumnCount * Count + 2);
                                    sheet.AddMergedRegion(region);
                                }
                                else if (j == verticalSheetColumnCount * Count + 3)
                                {
                                    cell.SetCellValue("目标化合物");
                                    CellRangeAddress nameregion = new CellRangeAddress(i,i,verticalSheetColumnCount * Count + 3,verticalSheetColumnCount * Count + 10);
                                    sheet.AddMergedRegion(nameregion);
                                }
                                else if (j == verticalSheetColumnCount * Count + 11)
                                {
                                    cell.SetCellValue("备注");
                                    CellRangeAddress namesregion = new CellRangeAddress(i,i + 4,verticalSheetColumnCount * Count + 11,verticalSheetColumnCount * Count + 11);
                                    sheet.AddMergedRegion(namesregion);
                                }
                            }
                            break;
                        }
                    case 1:
                        {

                            for (int j = verticalSheetColumnCount * Count; j < verticalSheetColumnCount * Count + verticalSheetColumnCount; j++)
                            {
                                var cell = row.CreateCell(j);
                                cell.CellStyle = bordercellStyle;

                                if (j == verticalSheetColumnCount * Count + 11)
                                {
                                    cell.SetCellValue(string.Empty);
                                }
                                else if (j > verticalSheetColumnCount * Count + 6 && j < verticalSheetColumnCount * Count + 7 + CompoundsNameList.Count)
                                {
                                    cell.SetCellValue(CompoundsNameList[j - verticalSheetColumnCount * Count - 7].Key);
                                }
                                else if (j > verticalSheetColumnCount * Count + 2 && j < verticalSheetColumnCount * Count + 3 + CompoundsNameList.Count)
                                {
                                    cell.SetCellValue(CompoundsNameList[j - verticalSheetColumnCount * Count - 3].Key);
                                }
                                else if (j == verticalSheetColumnCount * Count)
                                {
                                    CellRangeAddress region = new CellRangeAddress(i,i + 2,j,j + 2);
                                    sheet.AddMergedRegion(region);
                                    //要和公式那一块绑定在一起
                                    string wdm = FormulaComboBox.Text + "\n"
                                    + "C——样品中待测离子浓度，" + ZDJCCompanyComboBox.Text + "\n"
                                    + "Ci——查得样品中待测离子的浓度，" + TargetCompanyComboBox.Text + "\n"
                                    + "f——样品稀释倍数。";
                                    HSSFRichTextString rtsWdm = new HSSFRichTextString(wdm);
                                    cell.SetCellValue(rtsWdm);
                                }
                            }
                            break;
                        }
                    case 2:
                        {
                            for (int j = verticalSheetColumnCount * Count; j < verticalSheetColumnCount * Count + verticalSheetColumnCount; j++)
                            {
                                var cell = row.CreateCell(j);
                                cell.CellStyle = bordercellStyle;
                                if (j == verticalSheetColumnCount * Count + 3)
                                {
                                    cell.SetCellValue("/");
                                    CellRangeAddress firstregion = new CellRangeAddress(i,i,verticalSheetColumnCount * Count + 3,verticalSheetColumnCount * Count + 6);
                                    sheet.AddMergedRegion(firstregion);
                                }
                                else if (j == verticalSheetColumnCount * Count + 7)
                                {
                                    string value = string.Empty;
                                    if (testJCRadioButton.IsChecked == true)
                                    {
                                        value = "√\t" + testJCRadioButton.Content + "\t□\t" + testZDRadioButton.Content + "(" + ZDJCCompanyComboBox.Text + ")";
                                    }
                                    else if (testZDRadioButton.IsChecked == true)
                                    {
                                        value = "□\t" + testJCRadioButton.Content + "\t√\t" + testZDRadioButton.Content + "(" + ZDJCCompanyComboBox.Text + ")";
                                    }
                                    cell.SetCellValue(value);
                                    CellRangeAddress secondregion = new CellRangeAddress(i,i,verticalSheetColumnCount * Count + 7,verticalSheetColumnCount * Count + 10);
                                    sheet.AddMergedRegion(secondregion);
                                }
                            }
                            break;
                        }
                    case 3:
                        {
                            for (int j = verticalSheetColumnCount * Count; j < verticalSheetColumnCount * Count + verticalSheetColumnCount; j++)
                            {
                                var cell = row.CreateCell(j);
                                cell.CellStyle = bordercellStyle;
                                if (j == verticalSheetColumnCount * Count + 3)
                                {
                                    cell.SetCellValue("/");
                                    CellRangeAddress firstregion = new CellRangeAddress(i,i,verticalSheetColumnCount * Count + 3,verticalSheetColumnCount * Count + 6);
                                    sheet.AddMergedRegion(firstregion);
                                }
                                else if (j > verticalSheetColumnCount * Count + 6 && j < verticalSheetColumnCount * Count + 7 + CompoundsNameList.Count)
                                {
                                    cell.SetCellValue(CompoundsNameList[j - verticalSheetColumnCount * Count - 7].Value);
                                }
                            }
                            break;
                        }
                    case 4:
                        {
                            for (int j = verticalSheetColumnCount * Count; j < verticalSheetColumnCount * Count + verticalSheetColumnCount; j++)
                            {
                                var cell = row.CreateCell(j);
                                cell.CellStyle = bordercellStyle;
                                if (j == verticalSheetColumnCount * Count)
                                {
                                    cell.SetCellValue("样品编号");
                                    CellRangeAddress firstregion = new CellRangeAddress(i,i,verticalSheetColumnCount * Count,verticalSheetColumnCount * Count + 1);
                                    sheet.AddMergedRegion(firstregion);
                                }
                                else if (j == verticalSheetColumnCount * Count + 2)
                                {
                                    cell.SetCellValue("稀释倍数f");
                                }
                                else if (j == verticalSheetColumnCount * Count + 3)
                                {
                                    cell.SetCellValue("目标化合物测定值 M  (mg/L)");
                                    CellRangeAddress thirdregion = new CellRangeAddress(i,i,verticalSheetColumnCount * Count + 3,verticalSheetColumnCount * Count + 6);
                                    sheet.AddMergedRegion(thirdregion);
                                }
                                else if (j == verticalSheetColumnCount * Count + 7)
                                {
                                    cell.SetCellValue("目标化合物浓度 C (mg/L)");
                                    CellRangeAddress fourthregion = new CellRangeAddress(i,i,verticalSheetColumnCount * Count + 7,verticalSheetColumnCount * Count + 10);
                                    sheet.AddMergedRegion(fourthregion);
                                }
                            }
                            break;
                        }
                }
            }
            //正式数据录入
            for (int k = 0; k < cellList.Count; k++)
            {
                HSSFRow sampleRow = (Count == 0) ? (HSSFRow)sheet.CreateRow(5 + k) : (HSSFRow)sheet.GetRow(5 + k); //创建行或者获取行
                sampleRow.HeightInPoints = 20;
                string sampleName = cellList[k];
                string sampleDilutionratio = string.Empty;
                List<string> sampleSize = new List<string>();
                for (int p = 0; p < dataSet.Tables.Count; p++)
                {
                    DataTable dataTable = dataSet.Tables[p];
                    for (int r = 0; r < CompoundsNameList.Count; r++)
                    {
                        if (dataTable.TableName == CompoundsNameList[r].Key)
                        {
                            for (int o = 0; o < dataTable.Rows.Count; o++)
                            {
                                if (dataTable.Rows[o][0].ToString() == sampleName)
                                {
                                    if (sampleDilutionratio == string.Empty)
                                    {
                                        sampleDilutionratio = dataTable.Rows[o][1].ToString();
                                    }
                                    sampleSize.Add(dataTable.Rows[o][2].ToString());
                                    break;
                                }
                            }
                        }
                    }
                }

                for (int l = verticalSheetColumnCount * Count; l < verticalSheetColumnCount * Count + verticalSheetColumnCount; l++)
                {
                    var sampleCell = sampleRow.CreateCell(l);
                    sampleCell.CellStyle = bordercellStyle;
                    if (l == verticalSheetColumnCount * Count)
                    {
                        sampleCell.SetCellValue(sampleName);
                        CellRangeAddress firstregion = new CellRangeAddress(k + 5,k + 5,verticalSheetColumnCount * Count,verticalSheetColumnCount * Count + 1);
                        sheet.AddMergedRegion(firstregion);
                    }
                    else if (l == verticalSheetColumnCount * Count + 2)
                    {
                        sampleCell.SetCellValue(sampleDilutionratio);
                    }
                    else if (l > verticalSheetColumnCount * Count + 2 && l < verticalSheetColumnCount * Count + 3 + CompoundsNameList.Count)
                    {
                        if (sampleName.Contains("以下空白"))
                        {
                            sampleCell.SetCellValue(string.Empty);
                        }
                        else
                        {
                            string value = sampleSize[l - verticalSheetColumnCount * Count - 3];
                            if (!value.Contains("/"))
                            {
                                decimal C1 = decimal.Parse(value);
                                C1 = Math.Round(C1,4,MidpointRounding.ToEven);
                                value = CalculateAccuracyCX(C1.ToString(),4);
                            }
                            sampleCell.SetCellValue(value);
                        }
                    }
                    else if (l > verticalSheetColumnCount * Count + 6 && l < verticalSheetColumnCount * Count + 7 + CompoundsNameList.Count)
                    {

                        string value = string.Empty;
                        DataTable dataTable = dataSet.Tables[l - verticalSheetColumnCount * Count - 7];
                        if (sampleName.Contains("平均值"))
                        {
                            for (int o = 0; o < dataTable.Rows.Count; o++)
                            {
                                if (dataTable.Rows[o][0].ToString() == sampleName)
                                {
                                    decimal C1 = decimal.Parse(dataTable.Rows[o - 1][3].ToString());
                                    decimal C2 = decimal.Parse(dataTable.Rows[o - 2][3].ToString());
                                    value = CompareCompoundWithFormulaAverage(CompoundsNameList[l - verticalSheetColumnCount * Count - 7].Key,C1,C2);
                                    break;
                                }
                            }
                        }
                        else if (!sampleName.Contains("以下空白"))
                        {
                            value = CompareCompoundWithFormula(dataSet,CompoundsNameList[l - verticalSheetColumnCount * Count - 7].Key,sampleName,sampleDilutionratio,sampleSize[l - verticalSheetColumnCount * Count - 7]);
                        }
                        sampleCell.SetCellValue(value);
                    }
                    else if (l == verticalSheetColumnCount * Count + verticalSheetColumnCount - 1)
                    {
                        sampleCell.SetCellValue("/");
                    }
                }
            }
        }

        private string CompareCompoundWithFormulaAverage(string compoundName,decimal c1,decimal c2)
        {
            //检出限
            decimal modelC = decimal.Zero;
            decimal C = (c1 + c2) / 2;
            foreach (KeyValuePair<string,string> keyValuePair in compoundsNameList)
            {
                if (keyValuePair.Key == compoundName)
                {
                    modelC = decimal.Parse(keyValuePair.Value);
                }
            }
            if (C > modelC)
            {
                string realC = string.Empty;
                //大于等于1,取3位有效数字
                if (C >= 1 && C < 1000)
                {
                    realC = CalculateAccuracyThree(C);
                }
                //小于1,与检出限一致
                else if (C < 1)
                {
                    realC = CalculateAccuracyC(C,modelC);
                }
                else
                {
                    realC = ScientificCounting(C);
                }
                return realC;
            }
            //ND和<检出限的不一样
            if (testZDRadioButton.IsChecked == true)
            {
                return "<" + modelC;
            }
            else
            {
                return "ND";
            }
        }

        private string CompareCompoundWithFormula(DataSet dataSet,string compoundName,string sampleName,string sampleDilutionratio,string sampleSize)
        {
            //计算公式C = M×f
            //检出限
            decimal modelC = decimal.Zero;
            foreach (KeyValuePair<string,string> keyValuePair in compoundsNameList)
            {
                if (keyValuePair.Key == compoundName)
                {
                    modelC = decimal.Parse(keyValuePair.Value);
                }
            }
            //稀释倍数
            decimal f;
            //取样量
            decimal M;
            if (sampleDilutionratio.Contains("/") || sampleSize.Contains("/"))
            {
                //ND和<检出限的不一样
                if (testZDRadioButton.IsChecked == true)
                {
                    return "<" + modelC;
                }
                else
                {
                    return "ND";
                }
            }
            else
            {
                f = decimal.Parse(sampleDilutionratio);
                M = decimal.Parse(sampleSize);
                M = Math.Round(M,4,MidpointRounding.ToEven);
            }

            decimal taggetC = decimal.Parse((TargetCompanyComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
            decimal ZDJCCi = decimal.Parse((ZDJCCompanyComboBox.SelectedItem as ComboBoxItem).Tag.ToString());
            //单位换算
            decimal k = taggetC / ZDJCCi;
            //结果
            decimal C = M * f * k;

            if (C > modelC)
            {
                string realC = string.Empty;
                //大于等于1,取3位有效数字
                if (C >= 1 && C < 1000)
                {
                    realC = CalculateAccuracyThree(C);
                }
                //小于1,与检出限一致
                else if (C < 1)
                {
                    realC = CalculateAccuracyC(C,modelC);
                }
                else
                {
                    realC = ScientificCounting(C);
                }
                foreach (DataTable dataTable in dataSet.Tables)
                {
                    if (dataTable.TableName == compoundName)
                    {
                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            if (dataTable.Rows[i][0].ToString() == sampleName)
                            {
                                dataTable.Rows[i][3] = realC;
                                break;
                            }
                        }
                    }
                }
                return realC;
            }
            //ND和<检出限的不一样
            if (testZDRadioButton.IsChecked == true)
            {
                return "<" + modelC;
            }
            else
            {
                return "ND";
            }
        }

        private string CalculateAccuracyC(decimal C,decimal modelC)
        {
            string[] numC = modelC.ToString().Split(".");
            int numModelC = numC[1].Length;
            C = Math.Round(C,numModelC,MidpointRounding.ToEven);
            string realC = C.ToString();
            return realC;
        }



        /// <summary>
        /// 导出到Excel
        /// </summary>
        /// <param name="workbook"></param>
        private void ExportToExcel(HSSFWorkbook workbook)
        {
            //自己选位置
            /*System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
			fbd.ShowDialog();
			if (fbd.SelectedPath != string.Empty)
			{
				string filename = sheet.SheetName + ".xls";
				string path = System.IO.Path.Combine(fbd.SelectedPath,filename);
				using (FileStream stream = new FileStream(path,FileMode.OpenOrCreate,FileAccess.ReadWrite))
				{
					workbook.Write(stream);
					stream.Flush();
				}
			}*/
            //特定位置
            try
            {
                string path = @"E:\CreateExcel\" + ReportNo + @"\";
                //创建用户临时图片文件夹或者清空临时文件夹所有文件
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                string filename = ReportNo + "-离子色谱分析结果汇总表.xls";
                string fullpath = System.IO.Path.Combine(path,filename);
                if (File.Exists(fullpath))
                {
                    File.Delete(fullpath);
                }
                using (FileStream stream = new FileStream(fullpath,FileMode.OpenOrCreate,FileAccess.ReadWrite))
                {
                    workbook.Write(stream);
                    stream.Flush();
                }
                Process process = new Process();
                ProcessStartInfo processStartInfo = new ProcessStartInfo(fullpath);
                processStartInfo.UseShellExecute = true;
                process.StartInfo = processStartInfo;
                process.Start();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        private HSSFCellStyle CreateStyle(HSSFWorkbook workbook)
        {
            HSSFCellStyle cellStyle = (HSSFCellStyle)workbook.CreateCellStyle(); //创建列头单元格实例样式
            cellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center; //水平居中
            cellStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center; //垂直居中
            cellStyle.WrapText = true;//自动换行
                                      //cellStyle.BorderBottom = BorderStyle.Thin;
                                      //cellStyle.BorderRight = BorderStyle.Thin;
                                      //cellStyle.BorderTop = BorderStyle.Thin;
                                      //cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.TopBorderColor = HSSFColor.Black.Index;//DarkGreen(黑绿色)
            cellStyle.RightBorderColor = HSSFColor.Black.Index;
            cellStyle.BottomBorderColor = HSSFColor.Black.Index;
            cellStyle.LeftBorderColor = HSSFColor.Black.Index;

            return cellStyle;
        }

        /// <summary>
        /// 科学计数法
        /// </summary>
        /// <param name="testNum"></param>5
        /// <returns></returns>
        private string ScientificCounting(decimal testNum)
        {
            string returnnum = string.Empty;
            string oneNum = "1";
            if (testNum.ToString().Length >= 4)
            {
                for (int i = 0; i < testNum.ToString().Length - 1; i++)
                {
                    oneNum += "0";
                }

                decimal onenum = decimal.Parse(oneNum);
                decimal finalnum = Math.Round(testNum / onenum,2,MidpointRounding.ToEven);
                string finalvalue = CalculateAccuracyCX(finalnum.ToString(),2);
                returnnum = finalvalue + "×" + "10" + (testNum.ToString().Length - 1).ToString();
            }
            return returnnum;
        }

        /// <summary>
        /// 取三位有效数字
        /// </summary>
        /// <param name="C"></param>
        /// <returns></returns>
        private string CalculateAccuracyThree(decimal C)
        {
            string realC = C.ToString();
            if (C < 10)
            {
                C = Math.Round(C,2,MidpointRounding.ToEven);
                realC = CalculateAccuracyCX(C.ToString(),2);
            }
            else if (C < 100)
            {
                C = Math.Round(C,1,MidpointRounding.ToEven);
                realC = CalculateAccuracyCX(C.ToString(),1);
            }
            else if (C < 1000)
            {
                C = Math.Round(C,0,MidpointRounding.ToEven);
                realC = C.ToString();
            }
            return realC;
        }

        /// <summary>
        /// 补齐X位数的零
        /// </summary>
        /// <param name="compoundName"></param>
        /// <param name="v"></param>
        /// <returns></returns>
        private string CalculateAccuracyCX(string value,int X)
        {
            string[] beforeValue = value.Split(".");
            int num;
            //没有小数点的
            if (beforeValue.Length < 2)
            {
                num = X;
            }
            else
            {
                num = X - beforeValue[beforeValue.Length - 1].Length;
            }
            //计算后补零
            if (num != 0)
            {
                if (value.ToString().Contains("."))
                {
                    string answer = value.ToString();
                    for (int i = 0; i < num; i++)
                    {
                        answer += "0";
                    }
                    return answer;
                }
                else
                {
                    string answer = value.ToString() + ".";
                    for (int i = 0; i < num; i++)
                    {
                        answer += "0";
                    }
                    return answer;
                }
            }

            return value;
        }

        /// <summary>
        /// 搜索
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void searchTextBox_TextChanged(object sender,RoutedEventArgs e)
        {
            string searchText = searchTextBox.Text;
            TabControl tabControl = GetVisualChild<TabControl>(maingrid);
            if (tabControl != null)
            {
                foreach (TabItem tabItem in tabControl.Items)
                {
                    if (tabItem.IsSelected)
                    {
                        string header = tabItem.Header.ToString();
                        DataGrid dataGrid = tabItem.Content as DataGrid;
                        if (searchText != null && searchText != "")
                        {
                            for (int i = 0; i < dataGrid.ItemContainerGenerator.Items.Count - 1; i++)
                            {
                                dataGrid.ScrollIntoView(dataGrid.Items[i]);
                                DataGridRow dgv = (DataGridRow)dataGrid.ItemContainerGenerator.ContainerFromIndex(i);
                                if (dgv == null)
                                {
                                    dataGrid.UpdateLayout();
                                    dataGrid.ScrollIntoView(dataGrid.Items[i]);
                                    dgv = (DataGridRow)dataGrid.ItemContainerGenerator.ContainerFromIndex(i);
                                }
                                bool showdgv = false;
                                DataRow dr = (dgv.Item as DataRowView).Row;
                                for (int j = 0; j < dr.ItemArray.Length; j++)
                                {
                                    dgv.UpdateLayout();
                                    DataGridCellsPresenter presenter = GetVisualChild<DataGridCellsPresenter>(dgv);
                                    DataGridCell cell = (DataGridCell)presenter.ItemContainerGenerator.ContainerFromIndex(j);
                                    string cellcontent = dr[j].ToString().Trim();
                                    if (cellcontent.ToLower().Contains(searchText.ToLower()))
                                    {
                                        cell.Background = new SolidColorBrush(Colors.Orange);
                                        showdgv = true;
                                    }
                                    else
                                    {
                                        cell.Background = null;
                                    }
                                }
                                if (showdgv)
                                {
                                    dgv.Visibility = Visibility.Visible;
                                }
                                else
                                {
                                    dgv.Visibility = Visibility.Collapsed;
                                }
                            }
                        }
                        else
                        {
                            for (int i = 0; i < dataGrid.ItemContainerGenerator.Items.Count - 1; i++)
                            {
                                DataGridRow dgv = (DataGridRow)dataGrid.ItemContainerGenerator.ContainerFromIndex(i);
                                if (dgv == null)
                                {
                                    dataGrid.UpdateLayout();
                                    dataGrid.ScrollIntoView(dataGrid.Items[i]);
                                    dgv = (DataGridRow)dataGrid.ItemContainerGenerator.ContainerFromIndex(i);
                                }
                                dgv.Visibility = Visibility.Visible;
                                DataRow dr = (dgv.Item as DataRowView).Row;
                                for (int j = 0; j < dr.ItemArray.Length; j++)
                                {
                                    dgv.UpdateLayout();
                                    DataGridCellsPresenter presenter = GetVisualChild<DataGridCellsPresenter>(dgv);
                                    DataGridCell cell = (DataGridCell)presenter.ItemContainerGenerator.ContainerFromIndex(j);
                                    cell.Background = null;
                                }
                            }
                        }
                    }
                }
            }
        }

        #region 辅助函数
        /// <summary>
        /// 获取父可视对象中第一个指定类型的子可视对象
        /// </summary>
        /// <typeparam name="T">可视对象类型</typeparam>
        /// <param name="parent">父可视对象</param>
        /// <returns>第一个指定类型的子可视对象</returns>
        public static T GetVisualChild<T>(Visual parent) where T : Visual
        {
            T child = default(T);
            int numVisuals = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < numVisuals; i++)
            {
                Visual v = (Visual)VisualTreeHelper.GetChild(parent,i);
                child = v as T;
                if (child == null)
                {
                    child = GetVisualChild<T>(v);
                }
                if (child != null)
                {
                    break;
                }
            }
            return child;
        }

        /// <summary>
        /// 父控件+控件名找到子控件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="obj"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public T GetChildObject<T>(DependencyObject obj,string name) where T : FrameworkElement
        {
            DependencyObject child = null;
            T grandChild = null;
            for (int i = 0; i <= VisualTreeHelper.GetChildrenCount(obj) - 1; i++)
            {
                child = VisualTreeHelper.GetChild(obj,i);
                if (child is T && (((T)child).Name == name || string.IsNullOrEmpty(name)))
                {
                    return (T)child;
                }
                else
                {
                    grandChild = GetChildObject<T>(child,name);
                    if (grandChild != null)
                        return grandChild;
                }
            }
            return null;
        }


        #endregion
    }
}

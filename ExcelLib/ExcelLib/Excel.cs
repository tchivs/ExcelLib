using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using static System.IO.File;

namespace ExcelLib
{
    /*
     2018-08-17 13:43:53
     luoc@zhiweicl.com
         */

    /// <summary>
    /// Excel抽象类，封装了常用的方法，只需要实现Hanlder方法即可。
    /// </summary>
    public abstract class Excel
    {
        private bool _debugMode;

        /// <summary>
        /// 实例化Excel对象
        /// </summary>
        /// <param name="debugMode">设置Debug模式（Excel可见性，屏幕刷新，不提示警告窗体）</param>
        /// <param name="isNewExcel">是否为新的Excel对象</param>
        protected Excel(bool debugMode = true, bool isNewExcel = true)
        {
            try
            {
                ExcelApp = GetExcelApplication(isNewExcel);
                DebugMode = debugMode;
            }
            catch (InvalidCastException)
            {
                throw new COMException("对不起,没有获取到本机安装的Excel对象，请尝试修复或者安装Office2016后使用本软件！");
            }
        }

        /// <summary>
        /// 设置DEBUG模式
        /// </summary>
        public bool DebugMode
        {
            get => _debugMode;
            set
            {
                _debugMode = value;
                //设置是否显示警告窗体
                DisplayAlerts = value;
                //设置是否显示Excel
                Visible = value;
                //禁止刷新屏幕
                ScreenUpdating = value;
            }
        }

        /// <summary>
        /// 是否显示警告窗体
        /// </summary>
        public bool DisplayAlerts
        {
            get => ExcelApp.DisplayAlerts;
            set
            {
                if (ExcelApp.DisplayAlerts == value) return;
                ExcelApp.DisplayAlerts = value;
            }
        }

        /// <summary>
        /// Excel实例对象
        /// </summary>
        public Application ExcelApp { get; }

        /// <summary>
        /// 开启或者关闭屏幕刷新
        /// </summary>
        public bool ScreenUpdating
        {
            get => ExcelApp.ScreenUpdating;
            set
            {
                if (ExcelApp.ScreenUpdating == value) return;
                ExcelApp.ScreenUpdating = value;
            }
        }

        /// <summary>
        /// Excel可见性
        /// </summary>
        public bool Visible
        {
            get => ExcelApp.Visible;
            set
            {
                if (ExcelApp.Visible == value) return;
                ExcelApp.Visible = value;
            }
        }

        /// <summary>
        /// 插入图片
        /// </summary>
        /// <param name="path">图片路径</param>
        /// <param name="worksheet">要插入的表</param>
        /// <param name="range">要插入的range</param>
        public void AddPic(string path, Worksheet worksheet, Range range)
        {
            this.AddPic(path, worksheet, range, range.Width, range.Height);
        }

        /// <summary>
        /// 插入图片
        /// </summary>
        /// <param name="path">图片路径</param>
        /// <param name="worksheet">要插入的表</param>
        /// <param name="range">要插入的range</param>
        /// <param name="width">图片的宽度</param>
        /// <param name="height">图片的高度</param>
        public void AddPic(string path, Worksheet worksheet, Range range, int width, int height)
        {
            worksheet.Shapes.AddPicture(path, Microsoft.Office.Core.MsoTriState.msoCTrue,
                    Microsoft.Office.Core.MsoTriState.msoCTrue, range.Left, range.Top, width, height).Placement =
                XlPlacement.xlMoveAndSize;
        }

        /// <summary>
        /// 添加工作簿
        /// </summary>
        /// <returns></returns>
        public Workbook AddWorkbook()
        {
            Workbook workbook = ExcelApp.Workbooks.Add();
            return workbook;
        }

        /// <summary>
        /// 关闭工作簿
        /// </summary>
        /// <param name="workbook"></param>
        public void CloseWorkbook(Workbook workbook)
        {
            workbook.Close(false, Type.Missing, Type.Missing);
        }

        /// <summary>
        /// 复制表头到另一个sheet中
        /// </summary>
        /// <param name="sourceWorksheet">表头所在的sheet</param>
        /// <param name="newWorksheet">要复制到的sheet</param>
        /// <param name="start">起始位置</param>
        public void CopyHeader(Worksheet sourceWorksheet, Worksheet newWorksheet, int start = 1)
        {
            if (sourceWorksheet.Rows != null)
                sourceWorksheet.Rows[start].Copy(newWorksheet.Cells[1, 1]); //把数据表的表头复制到新表中
        }

        /// <summary>
        /// 复制列到另一张表
        /// </summary>
        /// <param name="sourceWorksheet">源表</param>
        /// <param name="sourceRows">源列</param>
        /// <param name="sourceStart">起始位置</param>
        /// <param name="newWorksheet">目的表</param>
        /// <param name="newRows">目的列</param>
        /// <param name="newStart">目的位置</param>
        public void CopyRow2OtherSheet(Worksheet sourceWorksheet, string[] sourceRows, int sourceStart,
            Worksheet newWorksheet, string[] newRows, int newStart)
        {
            int intrngEnd = GetEndRow(sourceWorksheet);
            if (newRows != null && (sourceRows != null && sourceRows.Length == newRows.Length))
            {
                for (int i = 0; i < sourceRows.Length; i++)
                {
                    string rg = sourceRows[i] + sourceStart + ":" + sourceRows[i] + intrngEnd;
                    sourceWorksheet.Range[rg]
                        .Copy(newWorksheet.Range[newRows[i] + newStart]);
                    //  new_worksheet.Cells[65536, new_rows[i]].End[XlDirection.xlUp].Offset(1, 0).Resize(intrngEnd, 1).Value = source_worksheet.Cells[2, source_rows[i]].Resize(intrngEnd, new_rows[i]).Value;
                }
            }
            else
            {
                Console.WriteLine("Error source_rows length not is new_rows length!");
            }
        }

        /// <summary>
        /// 复制列到另一张表
        /// </summary>
        /// <param name="sourceWorksheet">源表</param>
        /// <param name="sourceRows">源列</param>
        /// <param name="sourceStart">起始位置</param>
        /// <param name="newWorksheet">目的表</param>
        /// <param name="newRows">目的列</param>
        /// <param name="newStart">目的位置</param>
        public void CopyRow2OtherSheet(Worksheet sourceWorksheet, int[] sourceRows, int sourceStart, Worksheet newWorksheet,
            int[] newRows, int newStart)
        {
            int intrngEnd = GetEndRow(sourceWorksheet);
            if (sourceRows.Length == newRows.Length)
            {
                for (int i = 0; i < sourceRows.Length; i++)
                {
                    newWorksheet.Cells[65536, newRows[i]].End[XlDirection.xlUp].Offset(sourceStart, 0).Resize(intrngEnd, sourceStart)
                        .Value = sourceWorksheet.Cells[newStart, sourceRows[i]].Resize(intrngEnd, newRows[i]).Value;
                }
            }
            else
            {
                Console.WriteLine("Error source_rows length not is new_rows length!");
            }
        }

        /// <summary>
        /// 创建一个Excel对象
        /// </summary>
        /// <param name="visible">是否显示Excel，默认为True</param>
        /// <param name="caption">标题栏</param>
        /// <returns>返回创建好的Excel对象</returns>
        public Application CreateExcelApplication(bool visible = true, string caption = "New Application")
        {
            Application application = new Application
            {
                Visible = visible,
                Caption = caption
            };
            return application;
        }

        /// <summary>
        /// 退出Excel
        /// </summary>
        public void Exit()
        {
            if (ExcelApp.Workbooks.Count > 0)
            {
                ExcelApp.DisplayAlerts = false;
                ExcelApp.Workbooks.Close(); //关闭所有工作簿
            }
            ExcelApp.Quit(); //退出Excel
            ExcelApp.DisplayAlerts = true;
        }

        public string[] GetAllSheetNames(Workbook workbook)
        {
            List<string> res = new List<string>();
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                res.Add(worksheet.Name);
            }

            return res.ToArray();
        }

        /// <summary>
        /// 取有效列的最后一列的长度
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public int GetEndRow(Worksheet worksheet)
        {
            int res = worksheet.UsedRange.Rows.Count;
            return res;
        }

        /// <summary>
        /// 获取Excel对象，如果不存在则打开
        /// </summary>
        /// <returns>返回一个Excel对象</returns>
        public Application GetExcelApplication(bool newExcel)
        {
            Application application;
            try
            {
                if (newExcel)
                {
                    application = CreateExcelApplication();//创建一个新的Excel;
                }
                else
                {
                    application = (Application)Marshal.GetActiveObject("Excel.Application");//尝试取得正在运行的Excel对象
                }
                Debug.WriteLine("Get Running Excel");
            }
            //没有打开Excel则会报错
            catch (COMException)
            {
                application = CreateExcelApplication();//打开Excel
                Debug.WriteLine("Create new Excel");
            }
            Debug.WriteLine(application.Version);//打印Excel版本
            return application;
        }

        /// <summary>
        /// 获取workbook对象
        /// </summary>
        /// <param name="name">工作簿全名</param>
        /// <returns></returns>
        public Workbook GetWorkbook(string name)
        {
            Workbook wbk = ExcelApp.Workbooks[name];
            return wbk;
        }

        /// <summary>
        /// 获取workbook对象
        /// </summary>
        /// <param name="index">索引</param>
        /// <returns></returns>
        public Workbook GetWorkbook(int index)
        {
            Workbook wbk = ExcelApp.Workbooks[index];
            return wbk;
        }

        /// <summary>
        /// 获取workbook活动对象
        /// </summary>
        /// <returns></returns>
        public Workbook GetWorkbook()
        {
            Workbook wbk = ExcelApp.ActiveWorkbook;
            return wbk;
        }

        /// <summary>
        /// 主要实现这个方法
        /// </summary>
        /// <param name="path">要打开的文件路径</param>
        public abstract void Handler(string path = null);

        /// <summary>
        /// 批量插入图片
        /// </summary>
        /// <param name="pngdic">单元格范围-图片名</param>
        /// <param name="imgBase">图片根目录</param>
        /// <param name="worksheet">要插入图片的worksheet</param>
        /// <returns>返回处理好的图片日志</returns>
        public string InsertMultipleImages(Dictionary<string, string> pngdic, string imgBase, Worksheet worksheet)
        {
            string msg = null;
            foreach (KeyValuePair<string, string> s in pngdic)
            {
                string imgPath = Path.Combine(imgBase, s.Value);
                if (!Exists(imgPath))
                {
                    continue;
                }

                Range range = worksheet.Range[s.Key];
                AddPic(imgPath, worksheet, range);
                msg = s.Value + "\t" + s.Key + "\t\t\t" + range.Left.ToString() + "\t" + range.Top.ToString() + "\n";
            }

            return msg;
        }

        /// <summary>
        /// 杀死Excel进程
        /// </summary>
        public void Kill()
        {
            if (ExcelApp.Workbooks.Count > 0)
            {
                ExcelApp.DisplayAlerts = false;
                ExcelApp.Workbooks.Close(); //关闭所有工作簿
            }
            ExcelApp.Quit();
            GC.Collect();
            KeyMyExcelProcess.Kill(ExcelApp);
        }

        /// <summary>
        /// 打开或者查找表
        /// </summary>
        /// <param name="path"></param>
        /// <param name="filename"></param>
        /// <returns></returns>
        public Workbook OpenAndFindWorkbook(string path, string filename)
        {
            string pathFull = Path.Combine(path, filename);
            if (!Exists(pathFull))
            {
                pathFull = Directory.GetFiles(path, filename)[0];
            }
            //如果没有找到就直接打开文件
            return OpenAndFindWorkbook(filename);
        }

        /// <summary>
        /// 打开或者查找表
        /// </summary>
        /// <param name="filename">文件名全路径</param>
        /// <returns></returns>
        public Workbook OpenAndFindWorkbook(string filename)
        {
            string pathFull = filename;
            string fileName;
            string path = Path.GetDirectoryName(filename);
            if (!Exists(pathFull))
            {
                pathFull = Directory.GetFiles(path ?? throw new InvalidOperationException(), filename)[0];
                fileName = Path.GetFileName(pathFull);
            }
            else
            {
                fileName = Path.GetFileName(filename);
            }

            Workbook res = null;
            //遍历所有已打开的工作簿
            foreach (Workbook ws in ExcelApp.Workbooks)
            {
                if (ws.Name != fileName) continue;
                res = GetWorkbook(fileName); //OpenFromFile(umts_path).Worksheets[1];
                break;
            }

            //如果没有找到就直接打开文件
            return res ?? (OpenFromFile(pathFull));
        }

        /// <summary>
        /// 打开工作簿
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public Workbook OpenFromFile(string path)
        {
            Workbook workbook = ExcelApp.Workbooks.Open(path);
            return workbook;
        }

        /// <summary>
        /// 指定Sheet替换
        /// </summary>
        /// <param name="sheet">要进行替换的Sheet</param>
        /// <param name="oldstr">要进行替换的字符串</param>
        /// <param name="newstr">新的字符串</param>
        public void Replace(Worksheet sheet, string oldstr, string newstr)
        {
            sheet.UsedRange.Replace(oldstr, newstr, XlLookAt.xlPart, XlSearchOrder.xlByRows);
        }

        /// <summary>
        /// 全部Sheet替换
        /// </summary>
        /// <param name="sheet">要进行替换的Sheets</param>
        /// <param name="oldstr">要进行替换的字符串</param>
        /// <param name="newstr">新的字符串</param>
        public void Replace(Workbook workbook, string oldstr, string newstr)
        {
            Sheets sheets = workbook.Worksheets;
            foreach (Worksheet sheet in sheets)
            {
                Replace(sheet, oldstr, newstr);
            }
        }

        /// <summary>
        /// 保存工作簿
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="path"></param>
        public void SaveWorkbook(Workbook workbook, string path)
        {
            workbook.SaveAs(path);
        }

        /// <summary>
        /// 设置特定列的数据
        /// </summary>
        /// <param name="worksheet">源表</param>
        /// <param name="row">要设置的列号</param>
        /// <param name="len">长度</param>
        /// <param name="value">要设的值</param>
        /// ///
        public void SetSheetRow(Worksheet worksheet, int row, int len, string value)
        {
            //int intrngEnd = this.GetEndRow(worksheet);//取特定列最后一列的长度
            worksheet.Cells[65536, row].End[XlDirection.xlUp].Offset(1, 0).Resize(len, 1).Value = value;
        }

        /// <summary>
        /// 显示Excel窗口
        /// </summary>
        public void Show()
        {
            if (!ExcelApp.Visible)
            {
                ExcelApp.Visible = true;
            }
        }
    }

    /// <summary>
    /// 关闭Excel进程
    /// </summary>
    public class KeyMyExcelProcess
    {
        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int id);

        public static void Kill(Application excel)
        {
            try
            {
                IntPtr t = new IntPtr(excel.Hwnd);   //得到这个句柄，具体作用是得到这块内存入口
                GetWindowThreadProcessId(t, out int k);   //得到本进程唯一标志k
                System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);   //得到对进程k的引用
                p.Kill();     //关闭进程k
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
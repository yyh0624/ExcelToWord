using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NPOI.HSSF;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace ExcelToWord
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private static List<DataModel> _lisData = new List<DataModel>();
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string name = openFileDialog1.FileName;
                    textBox1.Text = openFileDialog1.SafeFileName;

                    var dt = ImportExcelDS(name);
                     if(dt.Tables.Count>0)
                    {
                        ExecDataSet(dt);
                    }
                }
            }
            catch (Exception ex)
            {

                throw;
            }
     
        }
     private   Action<DataSet> ExecDataSet = ds => {
         _lisData.Clear();
            foreach (DataTable item in ds.Tables)
            {
                var dm = new DataModel();
                dm._fileName = item.TableName;
                foreach (DataRow dr in item.Rows)
                {
                    if(string.IsNullOrEmpty(dr[0].ToString())&&string.IsNullOrEmpty(dr[1].ToString()))
                    {
                        if (dr[2].ToString().Equals("小写"))
                            dm._allPriceNum = Convert.ToDecimal(dr[4].ToString());
                        else if (dr[2].ToString().Equals("大写"))
                            dm._allPriceCN = dr[4].ToString();
                    }
                    else
                    {
                        dm._detils.Add(new DetilsModel()
                        {
                            _wzmc = dr[0].ToString(),
                            _ggxh = dr[1].ToString(),
                            _jldw = dr[2].ToString(),
                            sl = Convert.ToInt32(dr[3]),
                            zj = Convert.ToDecimal(dr[4]),
                            dj = Convert.ToDecimal(dr[5]),
                            zjclf = Convert.ToDecimal(dr[6]),
                            wgcjf = Convert.ToDecimal(dr[7]),
                            rljdlf = Convert.ToDecimal(dr[8]),
                            zjrgf = Convert.ToDecimal(dr[9]),
                            fpssf = Convert.ToDecimal(dr[10]),
                            glfy = Convert.ToDecimal(dr[11]),
                            lr = Convert.ToDecimal(dr[12]),
                            sj = Convert.ToDecimal(dr[13]),
                            bjgjf = Convert.ToDecimal(dr[14]),
                            aztsf = Convert.ToDecimal(dr[15]),
                            jsfwf = Convert.ToDecimal(dr[16]),
                            yzf = Convert.ToDecimal(dr[17]),
                            _pinpai = dr[18].ToString(),
                            _zxbz = dr[19].ToString(),
                            _chandi=dr[20].ToString()
                        });
                    }
                }
             _lisData.Add(dm);
                
            }
        };

        /// <summary>
        /// 获取excel内容
        /// </summary>
        /// <param name="filePath">excel文件路径</param>
        /// <returns></returns>
        public static DataTable ImportExcelDT(string filePath)
        {
            DataTable dt = new DataTable();
            using (FileStream fsRead = System.IO.File.OpenRead(filePath))
            {
                IWorkbook wk = null;
                //获取后缀名
                string extension = filePath.Substring(filePath.LastIndexOf(".")).ToString().ToLower();
                //判断是否是excel文件
                if (extension == ".xlsx" || extension == ".xls")
                {
                    //判断excel的版本
                    if (extension == ".xlsx")
                    {
                        wk = new XSSFWorkbook(fsRead);
                    }
                    else
                    {
                        wk = new HSSFWorkbook(fsRead);
                    }

                    //获取第一个sheet
                    ISheet sheet = wk.GetSheetAt(0);
                    //获取第一行
                    IRow headrow = sheet.GetRow(0);
                    //创建列
                    for (int i = headrow.FirstCellNum; i < headrow.Cells.Count; i++)
                    {
                        //  DataColumn datacolum = new DataColumn(headrow.GetCell(i).StringCellValue);
                        DataColumn datacolum = new DataColumn("F" + (i + 1));
                        dt.Columns.Add(datacolum);
                    }
                    //读取每行,从第二行起
                    for (int r = 0; r <= sheet.LastRowNum; r++)
                    {
                        bool result = false;
                        DataRow dr = dt.NewRow();
                        //获取当前行
                        IRow row = sheet.GetRow(r);
                        //读取每列
                        for (int j = 0; j < row.Cells.Count; j++)
                        {
                            NPOI.SS.UserModel.ICell cell = row.GetCell(j); //一个单元格
                            dr[j] = GetCellValue(cell); //获取单元格的值
                                                        //全为空则不取
                            if (dr[j].ToString() != "")
                            {
                                result = true;
                            }
                        }
                        if (result == true)
                        {
                            dt.Rows.Add(dr); //把每行追加到DataTable
                        }
                    }
                }

            }
            return dt;
        }
        public static DataSet ImportExcelDS(string filePath)
        {
            DataSet ds = new DataSet();
            using (FileStream fsRead = System.IO.File.OpenRead(filePath))
            {
                IWorkbook wk = null;
                //获取后缀名
                string extension = filePath.Substring(filePath.LastIndexOf(".")).ToString().ToLower();
                //判断是否是excel文件
                if (extension == ".xlsx" || extension == ".xls")
                {
                    //判断excel的版本
                    if (extension == ".xlsx")
                    {
                        wk = new XSSFWorkbook(fsRead);
                    }
                    else
                    {
                        wk = new HSSFWorkbook(fsRead);
                    }
                    for (int z = 0; z < wk.NumberOfSheets; z++)
                    {
                        DataTable dt = new DataTable();
                        //获取第一个sheet
                        ISheet sheet = wk.GetSheetAt(z);
                        //获取第一行
                        IRow headrow = sheet.GetRow(0);
                        //创建列
                        for (int i = headrow.FirstCellNum; i < headrow.Cells.Count; i++)
                        {
                              DataColumn datacolum = new DataColumn(headrow.GetCell(i).StringCellValue);
                           // DataColumn datacolum = new DataColumn("F" + (i + 1));
                            dt.Columns.Add(datacolum);
                        }
                        //读取每行,从第二行起
                        for (int r = 1; r <= sheet.LastRowNum; r++)
                        {
                            bool result = false;
                            DataRow dr = dt.NewRow();
                            //获取当前行
                            IRow row = sheet.GetRow(r);
                            //读取每列
                            for (int j = 0; j < row.Cells.Count; j++)
                            {
                                NPOI.SS.UserModel.ICell cell = row.GetCell(j); //一个单元格
                                dr[j] = GetCellValue(cell); //获取单元格的值
                                                            //全为空则不取
                                if (dr[j].ToString() != "")
                                {
                                    result = true;
                                }
                            }
                            if (result == true)
                            {
                                dt.Rows.Add(dr); //把每行追加到DataTable
                            }
                        }
                        dt.TableName = sheet.SheetName;
                        ds.Tables.Add(dt);

                    } 
                }

            }
            return ds;
        }
        //对单元格进行判断取值
        private static string GetCellValue(NPOI.SS.UserModel.ICell cell)
        {
            if (cell == null)
                return string.Empty;
            switch (cell.CellType)
            {
                case CellType.Blank: //空数据类型 这里类型注意一下，不同版本NPOI大小写可能不一样,有的版本是Blank（首字母大写)
                    return string.Empty;
                case CellType.Boolean: //bool类型
                    return cell.BooleanCellValue.ToString();
                case CellType.Error:
                    return cell.ErrorCellValue.ToString();
                case CellType.Numeric: //数字类型
                    if (HSSFDateUtil.IsCellDateFormatted(cell))//日期类型
                    {
                        return cell.DateCellValue.ToString();
                    }
                    else //其它数字
                    {
                        return cell.NumericCellValue.ToString();
                    }
                case CellType.Unknown: //无法识别类型
                default: //默认类型
                    return cell.ToString();//
                case CellType.String: //string 类型
                    return cell.StringCellValue;
                case CellType.Formula: //带公式类型
                    try
                    {
                        HSSFFormulaEvaluator e = new HSSFFormulaEvaluator(cell.Sheet.Workbook);
                        e.EvaluateInCell(cell);
                        return cell.ToString();
                    }
                    catch
                    {
                        return cell.NumericCellValue.ToString();
                    }
            }
        }
        private  void SaveToWord(string path)
        {
            XWPFDocument doc = new XWPFDocument();      //创建新的word文档

            XWPFParagraph p1 = doc.CreateParagraph();   //向新文档中添加段落
          

            XWPFParagraph p2 = doc.CreateParagraph(); 

            XWPFRun r2 = p2.CreateRun();
            r2.SetText("测试段落二");
       

            FileStream sw = File.Create("cutput.docx"); //...
            doc.Write(sw);                              //...
            sw.Close();                                 //在服务端生成文件

            FileInfo file = new FileInfo("cutput.docx");//文件保存路径及名称  
           
             
        }
    }
}

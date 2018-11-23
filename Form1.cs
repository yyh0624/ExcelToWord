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
        private static string _choosePath = string.Empty;
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    _choosePath = Path.GetDirectoryName(openFileDialog1.FileName) + "\\newDoc\\";
                    if (!Directory.Exists(_choosePath))
                    {
                        Directory.CreateDirectory(_choosePath);
                    }
                    string name = openFileDialog1.FileName;
                    textBox1.Text = openFileDialog1.SafeFileName;

                    var dt = ImportExcelDS(name);
                    if (dt.Tables.Count > 0)
                    {
                        ExecDataSet(dt);
                        foreach (var item in _lisData)
                        {
                            SaveToWord(_choosePath, item);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }

        }
        private Action<DataSet> ExecDataSet = ds =>
        {
            _lisData.Clear();
            foreach (DataTable item in ds.Tables)
            {
                var dm = new DataModel();
                dm._fileName = item.TableName;
                foreach (DataRow dr in item.Rows)
                {
                    if (string.IsNullOrEmpty(dr[0].ToString()) && string.IsNullOrEmpty(dr[1].ToString()))
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
                            _chandi = dr[20].ToString()
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
        private void SaveToWord(string path, DataModel model)
        {
            XWPFDocument doc = new XWPFDocument();      //创建新的word文档
            XWPFParagraph p1 = doc.CreateParagraph();   //向新文档中添加段落
            p1.Alignment = ParagraphAlignment.CENTER | ParagraphAlignment.BOTH;

            XWPFRun r1 = p1.CreateRun();
            r1.SetText("目录 ");
            doc.Document.body.AddNewP().AddNewR().AddNewInstrText().Value = "1、开标一览表";
            doc.Document.body.AddNewP().AddNewR().AddNewInstrText().Value = "2、价格构成表";
            doc.Document.body.AddNewP().AddNewR().AddNewInstrText().Value = "3、货物材料、部件、工具价格明细表";
            doc.Document.body.AddNewP().AddNewR().AddNewInstrText().Value = "4、 其他与价格有关的资料、文件 ";

            doc.Document.body.AddNewP().AddNewR().AddNewInstrText().Value = "开标一览表";
            doc.Document.body.AddNewP().AddNewR().AddNewInstrText().Value = "投标人全称：北京航天长峰股份有限公司   项目名称：卫生物资   项目编号：2016 - ZHSH - 1049   包号：第十二包   金额单位：元";
            XWPFTable table = doc.CreateTable(1, 10);//创建table

            table.SetColumnWidth(0, 500);//设置列的宽度
            table.SetColumnWidth(1, 2000);//设置列的宽度
            table.SetColumnWidth(2, 1000);//设置列的宽度
            table.SetColumnWidth(3, 1000);//设置列的宽度
            table.SetColumnWidth(4, 500);//设置列的宽度
            table.SetColumnWidth(5, 500);//设置列的宽度
            table.SetColumnWidth(6, 900);//设置列的宽度
            table.SetColumnWidth(7, 900);//设置列的宽度
            table.SetColumnWidth(8, 2000);//设置列的宽度
            table.SetColumnWidth(9, 500);//设置列的宽度
            string[] t1 = { "序号", "货物名称", "品牌", "规格型号", "计量单位", "数量", "单价（含税）", "金额（含税）", "交货时间", "备注" };
            for (int i = 0; i < 10; i++)
            {
                table.Rows[0].GetCell(i).SetText(t1[i]);
            }

            table.Rows[0].GetCTRow().AddNewTrPr().AddNewTrHeight().val = (ulong)1000;//设置行高
            int _t1Count = 1;
            foreach (var item in model._detils)
            {
                var r = table.CreateRow();
                r.GetCell(0).SetText(_t1Count.ToString());
                r.GetCell(1).SetText(item._wzmc);
                r.GetCell(2).SetText(item._pinpai);
                r.GetCell(3).SetText(item._ggxh);
                r.GetCell(4).SetText(item._jldw);
                r.GetCell(5).SetText(item.sl.ToString());
                r.GetCell(6).SetText(item.dj.ToString());
                r.GetCell(7).SetText(item.zj.ToString());
                r.GetCell(8).SetText("正式合同签订后3个月");
                r.GetCell(9).SetText("");
                _t1Count++;
            }
            var hj = table.CreateRow();

            hj.GetCell(0).GetCTTc().AddNewTcPr().AddNewGridspan().val = "2";
            hj.GetCell(0).SetText("合计");
            hj.GetCell(6).SetText(model._allPriceNum.ToString());
            hj.RemoveCell(9);
            var zjnewrow = new CT_Row();
            var zjm_Row = new XWPFTableRow(zjnewrow, table);
            table.AddRow(zjm_Row);
            zjm_Row.CreateCell();
            zjm_Row.CreateCell();
            zjm_Row.CreateCell();
            zjm_Row.GetCell(2).GetCTTc().AddNewTcPr().AddNewGridspan().val = "8";
            zjm_Row.GetCell(2).SetText("投标总价（人民币大写）：" + model._allPriceCN + "     （小写）¥：" + model._allPriceNum);

            doc.Document.body.AddNewP().AddNewR().AddNewInstrText().Value = "说明：金额=单价×数量，投标总价=金额之和。";


            doc.Document.body.AddNewP().AddNewR().AddNewInstrText().Value = "价格构成表";
            doc.Document.body.AddNewP().AddNewR().AddNewInstrText().Value = " 项目名称：卫生物资 项目编号：2016 - ZHSH - 1049             包号：第十二包 金额单位：元";


            XWPFTable table2 = doc.CreateTable(1, 18);//创建table
            table2.SetColumnWidth(0, 1200);//设置列的宽度
            table2.SetColumnWidth(1, 800);//设置列的宽度
            table2.SetColumnWidth(2, 500);//设置列的宽度
            table2.SetColumnWidth(3, 500);//设置列的宽度
            table2.SetColumnWidth(4, 500);//设置列的宽度
            table2.SetColumnWidth(5, 500);//设置列的宽度
            table2.SetColumnWidth(6, 500);//设置列的宽度
            table2.SetColumnWidth(7, 500);//设置列的宽度
            table2.SetColumnWidth(8, 500);//设置列的宽度
            table2.SetColumnWidth(9, 500);//设置列的宽度
            table2.SetColumnWidth(10, 500);//设置列的宽度
            table2.SetColumnWidth(11, 500);//设置列的宽度
            table2.SetColumnWidth(12, 500);//设置列的宽度
            table2.SetColumnWidth(13, 500);//设置列的宽度
            table2.SetColumnWidth(14, 500);//设置列的宽度
            table2.SetColumnWidth(15, 500);//设置列的宽度
            table2.SetColumnWidth(16, 500);//设置列的宽度
            table2.SetColumnWidth(17, 500);//设置列的宽度
            string[] t2 = { "货物名称", "规格型号", "计量单位", "数量", "总价" };
            string[] t2_2 = { "单价", "直接材料费", "外购成件费", "燃料及动力费", "直接人工费", "废品损失费", "管理费用", "利润", "税金", "备件工具费", "安装调试费", "技术服务费", "运杂费" };

            var newrow = new CT_Row();
            var m_Row = new XWPFTableRow(newrow, table2);
            table2.AddRow(m_Row);
            m_Row.CreateCell();
            m_Row.CreateCell();
            m_Row.CreateCell();
            m_Row.CreateCell();
            m_Row.CreateCell();
            var newrow2 = new CT_Row();
            var m_Row2 = new XWPFTableRow(newrow2, table2);
            table2.AddRow(m_Row2);
            m_Row2.CreateCell();
            m_Row2.CreateCell();
            m_Row2.CreateCell();
            m_Row2.CreateCell();
            m_Row2.CreateCell();
            for (int i = 0; i < 5; i++)
            {
                // var cell = m_Row.CreateCell();

                // cell.GetCTTc().AddNewTcPr().AddNewVAlign().val = ST_VerticalJc.center;//垂直居中 
                // m_Row.GetCell(i).GetCTTc().AddNewTcPr().AddNewGridspan().val = "2";
                // m_Row.GetCell(i).GetCTTc().AddNewTcPr().AddNewVMerge().val = ST_Merge.@continue;
                // cell.GetCTTc().GetPList()[0].AddNewR().AddNewT().Value = t2[i];
                m_Row.GetCell(i).SetText(t2[i]);

            }
            var cell2 = m_Row.CreateCell();
            cell2.GetCTTc().AddNewTcPr().AddNewGridspan().val = "13";
            cell2.SetText(" 价  格  组  成");

            for (int i = 5; i < 18; i++)
            {
                var cell = m_Row2.CreateCell();
                cell.SetText(t2_2[i - 5]);
            }


            foreach (var item in model._detils)
            {
                var r = table2.CreateRow();
                r.GetCell(0).SetText(item._wzmc);
                r.GetCell(1).SetText(item._ggxh);
                r.GetCell(2).SetText(item._jldw);
                r.GetCell(3).SetText(item.sl.ToString());
                r.GetCell(4).SetText(item.zj.ToString());
                r.GetCell(5).SetText(item.dj.ToString());
                r.GetCell(6).SetText(item.zjclf.ToString());
                r.GetCell(7).SetText(item.wgcjf.ToString());
                r.GetCell(8).SetText(item.rljdlf.ToString());
                r.GetCell(9).SetText(item.zjrgf.ToString());
                r.GetCell(10).SetText(item.fpssf.ToString());
                r.GetCell(11).SetText(item.glfy.ToString());
                r.GetCell(12).SetText(item.lr.ToString());
                r.GetCell(13).SetText(item.sj.ToString());
                r.GetCell(14).SetText(item.bjgjf.ToString());
                r.GetCell(15).SetText(item.aztsf.ToString());
                r.GetCell(16).SetText(item.jsfwf.ToString());
                r.GetCell(17).SetText(item.yzf.ToString());
            }
            var zj2newrow = new CT_Row();
            var zj2m_Row = new XWPFTableRow(zj2newrow, table2);
            table2.AddRow(zj2m_Row);
            zj2m_Row.CreateCell();
            zj2m_Row.GetCell(0).GetCTTc().AddNewTcPr().AddNewGridspan().val = "18";
            zj2m_Row.GetCell(0).SetText("投标总价（人民币大写）：" + model._allPriceCN + "     （小写）¥：" + model._allPriceNum);



            table2.RemoveRow(0);//去掉第一行空白的
            doc.Document.body.AddNewP().AddNewR().AddNewInstrText().Value = " 说明:1.项5=项6×项4";
            doc.Document.body.AddNewP().AddNewR().AddNewInstrText().Value = " 2.项6 = 项7 + 项8 + 项9 + 项10 + 项11 + 项12 + 项13 + 项14 + 项15 + 项16 + 项17 + 项18";

            doc.Document.body.AddNewP().AddNewR().AddNewInstrText().Value = "货物材料、部件、工具价格明细表 ";
            doc.Document.body.AddNewP().AddNewR().AddNewInstrText().Value = "项目名称：卫生物资                         项目编号：2016 - ZHSH - 1049                               包号：第十二包";

            string[] t3 = { "序号", "项目", "规格型号", "执行标准", "计量单位", "定额/消耗数量", "单价(元)", "金额(元)", "产地或生产企业" };
            XWPFTable table3 = doc.CreateTable(1, 9);//创建table
            table3.SetColumnWidth(0, 500);//设置列的宽度
            table3.SetColumnWidth(1, 2000);//设置列的宽度
            table3.SetColumnWidth(2, 800);//设置列的宽度
            table3.SetColumnWidth(3, 2500);//设置列的宽度
            table3.SetColumnWidth(4, 500);//设置列的宽度
            table3.SetColumnWidth(5, 500);//设置列的宽度
            table3.SetColumnWidth(6, 1100);//设置列的宽度
            table3.SetColumnWidth(7, 1100);//设置列的宽度
            table3.SetColumnWidth(8, 800);//设置列的宽度  
            for (int i = 0; i < 9; i++)
            {
                table3.Rows[0].GetCell(i).SetText(t3[i]);
            }
            decimal? totalPrice = 0;
            var t3_1 = table3.CreateRow();
            t3_1.GetCell(0).SetText("一");
            t3_1.GetCell(1).SetText("直接材料费");
            int _t3count1 = 1;
            foreach (var item in model._detils)
            {
                var _r = table3.CreateRow();
                _r.GetCell(0).SetText(_t3count1.ToString());
                _r.GetCell(1).SetText(item._wzmc);
                _r.GetCell(2).SetText(item._ggxh);
                _r.GetCell(3).SetText(item._zxbz);
                _r.GetCell(4).SetText("套");
                _r.GetCell(5).SetText("1");
                _r.GetCell(6).SetText(item.zjclf.ToString());
                _r.GetCell(7).SetText(item.zjclf.ToString());
                _r.GetCell(8).SetText(item._chandi);
                _t3count1++;
            }
            var t3_2 = table3.CreateRow();
            t3_2.GetCell(0).SetText("二");
            t3_2.GetCell(1).SetText("外购成件费");
            int _t3count2 = 1;
            foreach (var item in model._detils)
            {
                var _r = table3.CreateRow();
                _r.GetCell(0).SetText(_t3count2.ToString());
                _r.GetCell(1).SetText(item._wzmc);
                _r.GetCell(2).SetText(item._ggxh);
                _r.GetCell(3).SetText(item._zxbz);
                _r.GetCell(4).SetText("套");
                _r.GetCell(5).SetText("1");
                _r.GetCell(6).SetText(item.wgcjf.ToString());
                _r.GetCell(7).SetText(item.wgcjf.ToString());
                _r.GetCell(8).SetText(item._chandi);
                _t3count2++;
            }
            var t3_3 = table3.CreateRow();
            t3_3.GetCell(0).SetText("三");
            t3_3.GetCell(1).SetText("备件工具费");
            int _t3count3 = 1;
            foreach (var item in model._detils)
            {
                var _r = table3.CreateRow();
                _r.GetCell(0).SetText(_t3count3.ToString());
                _r.GetCell(1).SetText(item._wzmc);
                _r.GetCell(2).SetText(item._ggxh);
                _r.GetCell(3).SetText(item._zxbz);
                _r.GetCell(4).SetText("套");
                _r.GetCell(5).SetText("1");
                _r.GetCell(6).SetText(item.bjgjf.ToString());
                _r.GetCell(7).SetText(item.bjgjf.ToString());
                _r.GetCell(8).SetText(item._chandi);
                _t3count3++;
            }

            totalPrice += model._detils.Sum(p => p.zjclf);
            totalPrice += model._detils.Sum(p => p.bjgjf);
            totalPrice += model._detils.Sum(p => p.wgcjf);

            var t3_4 = table3.CreateRow();
            t3_4.GetCell(1).SetText("合计");
            t3_4.GetCell(7).SetText(totalPrice.Value.ToString());
            doc.Document.body.AddNewP().AddNewR().AddNewInstrText().Value = "说明：以一套货物的所用材料为基本单位，项目填列直接材料明细。";
            using (FileStream sw = File.Create(path + model._fileName + ".doc"))
            {
                doc.Write(sw);
            }
        }
    }
}

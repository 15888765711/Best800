using _800Best.ExcelHelpModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _800Best.ExcelHelpDAL
{
    public class MyExcelDal
    {
        /// <summary>
        /// 改变结算类型开始
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public ISheet ChangeExcel(ISheet sheet)
        {
            IRow row = sheet.GetRow(0);
            if ((row != null) && (row.GetCell(0) != null))
            {
                int lastRowNum = sheet.LastRowNum;
                int lastCellNum = row.LastCellNum;
                int cellnum = -1;
                int num4 = -1;
                int num5 = -1;
                for (int i = 0; i < lastCellNum; i++)
                {
                    if (row.GetCell(i).StringCellValue == "结算类型")
                    {
                        cellnum = i;
                    }
                    else if (row.GetCell(i).StringCellValue == "结算金额")
                    {
                        num4 = i;
                    }
                    else if (row.GetCell(i).StringCellValue == "开户站点")
                    {
                        num5 = i;
                    }
                }
                if ((num4 == -1) || (cellnum == -1)||num5==-1)
                {
                    return null;
                }
                for (int j = 1; j <= lastRowNum; j++)
                {
                    ICell cell = sheet.GetRow(j).GetCell(cellnum);
                    this.ChangeExcelType(cell, sheet.GetRow(j).GetCell(num4).NumericCellValue);
                }
            }
            return sheet;
        }
        /// <summary>
        /// 修改结算类型-实现
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="numericCellValue"></param>
        private void ChangeExcelType(ICell cell, double numericCellValue)
        {
            string stringCellValue = cell.StringCellValue;
            if (stringCellValue != null)
            {
                switch (stringCellValue)
                {

                    case "寄件派费":
                        if (numericCellValue > 0)
                        { cell.SetCellValue("寄件派费调整"); }
                        break;
                    case "付有偿派费": cell.SetCellValue("派件费");break;
                    case "代转件费":
                        if (numericCellValue > 0)
                        {
                            cell.SetCellValue("代转件费调整");
                        }
                        break;
                    case "中转费调整":
                        if (numericCellValue > 0)
                        {
                            cell.SetCellValue("中转费取消");
                        }
                        break;
                    case "派件派费":
                        if (numericCellValue > 0)
                        {
                            cell.SetCellValue("派件派费调整");
                        }
                        break;
                    case "网点派件派费-赋能":
                        if (numericCellValue > 0)
                        {
                            cell.SetCellValue("网点派件派费-赋能调整");
                        }
                        break;
                    //派件派费
                    case "计重收费调整":
                        if (numericCellValue < 0)
                        {
                            cell.SetCellValue("计重收费");
                        }
                        break;
                    case "违禁品罚款":
                        cell.SetCellValue("航空违禁品罚款");
                        break;
                    case "扣有偿中转调整":
                        if (numericCellValue > 0)
                        {
                            cell.SetCellValue("扣有偿中转取消");
                        }
                        break;
                    case "扫描费":
                        cell.SetCellValue("扫描费调整");
                        break;                   
                    case "保价手续费":
                        cell.SetCellValue("手续费");
                        break;
                    case "批货大货费":
                        cell.SetCellValue("大货费");
                        break;
                    case "大货手续费":
                        if (numericCellValue > 0)
                        {
                            cell.SetCellValue("大货费调整");
                        }
                        break;
                    case "中转费-应集未集":

                        cell.SetCellValue("应集未集补收罚款");
                        break;
                    case "中转费-应集未集调整":

                        cell.SetCellValue("应集未集补收罚款调整");
                        break;
                    case "错集率罚款":
                    case "短信服务费":
                    case "环保袋滞留费":
                    case "站点留仓率罚款":
                    case "菜鸟及时揽收率罚款":
                    case "我要寄百世爽约率考核罚款":
                    case "计重收费返款":
                    case "激励政策返款":
                    case "环保袋使用费":
                        //计重收费返款
                        //我要寄百世爽约率考核罚款
                        if (cell.Row.GetCell(0) != null)
                        {
                            cell.Row.GetCell(4).SetCellValue(cell.Row.GetCell(0).StringCellValue);
                            cell.Row.GetCell(0).SetCellValue("");
                        }
                        break;

                }
            }
        }

        public ISheet GetXinQiaoSummarySheet(ISheet sheet)
        {
            IRow row = sheet.CreateRow(1);
            row.CreateCell(4).SetCellValue("上传数量");
            row.CreateCell(5).SetCellValue("上传金额");
            row.CreateCell(6).SetCellValue("未上传数量");
            row.CreateCell(7).SetCellValue("未上传金额");
            row.CreateCell(8).SetCellValue("总数量");
            row.CreateCell(9).SetCellValue("总金额");
            row.CreateCell(10).SetCellValue("核对");
            row.CreateCell(11).SetCellValue("差异");
            row = sheet.CreateRow(2);
            row.CreateCell(3).SetCellValue("系统扣费");
            row.CreateCell(4).SetCellFormula("COUNTA(藤桥运单扣费1!G:G,藤桥运单扣费2!G:G,藤桥运单扣费3!G:G,藤桥运单扣费4!G:G)-4");
            row.CreateCell(5).SetCellFormula("SUM(藤桥运单扣费1!D:D,藤桥运单扣费2!D:D,藤桥运单扣费3!D:D,藤桥运单扣费4!D:D)");
            row.CreateCell(6).SetCellFormula("COUNTA(未分类站点!G:G)-1");
            row.CreateCell(7).SetCellFormula("SUM(未分类站点!D:D)");
            row.CreateCell(8).SetCellFormula("E3+G3");
            row.CreateCell(9).SetCellFormula("F3+H3");
            row.CreateCell(11).SetCellFormula("ROUND(K3-J3,2)");
            row = sheet.CreateRow(3);
            row.CreateCell(3).SetCellValue("集包扣费");
            row.CreateCell(4).SetCellFormula("COUNTA(藤桥集包!C:C)-1");
            row.CreateCell(5).SetCellFormula("SUM(藤桥集包!D:D)");
            row.CreateCell(8).SetCellFormula("E4+G4");
            row.CreateCell(9).SetCellFormula("F4+H4");
            row = sheet.CreateRow(4);
            row.CreateCell(3).SetCellValue("合计");
            row.CreateCell(4).SetCellFormula("SUM(E3:E4)");
            row.CreateCell(5).SetCellFormula("SUM(F3:F4)");
            row.CreateCell(6).SetCellFormula("SUM(G3:G4)");
            row.CreateCell(7).SetCellFormula("SUM(H3:H4)");
            row.CreateCell(8).SetCellFormula("SUM(I3:I4)");
            row.CreateCell(9).SetCellFormula("SUM(J3:J4)");
            row = sheet.CreateRow(6);
            row.CreateCell(3).SetCellValue("1.代集包费3KG以下0.35，3KG以上0.1*重量，取两位小数");
            row = sheet.CreateRow(7);
            row.CreateCell(3).SetCellValue("2.003站点代集包费不收集包费");
            row = sheet.CreateRow(8);
            row.CreateCell(3).SetCellValue("3.系统扣费中先关联S9数据，分配到归属站点先，分配不到归属站点再关联派件数据，做到温州藤桥一部，其余做到未分类站点，关联发放网点");
            row = sheet.CreateRow(9);
            row.CreateCell(3).SetCellValue("4.网络资讯服务费做到温州藤桥分部001");
            row = sheet.CreateRow(10);
            row.CreateCell(3).SetCellValue("5.付有偿派费→派件费");
            row.CreateCell(9).SetCellValue(DateTime.Today.AddDays(-1.0).ToShortDateString());
            sheet.CreateRow(12).CreateCell(3).SetCellValue("分批上传");
            sheet.GetRow(12).CreateCell(4).SetCellValue("上传数量");
            sheet.GetRow(12).CreateCell(5).SetCellValue("上传金额");
            sheet.CreateRow(13).CreateCell(3).SetCellValue("第一批");
            sheet.GetRow(13).CreateCell(4).SetCellFormula("COUNTA(藤桥运单扣费1!G:G)-1");
            sheet.GetRow(13).CreateCell(5).SetCellFormula("SUM(藤桥运单扣费1!D:D)");
            sheet.CreateRow(14).CreateCell(3).SetCellValue("第二批");
            sheet.GetRow(14).CreateCell(4).SetCellFormula("COUNTA(藤桥运单扣费2!G:G)-1");
            sheet.GetRow(14).CreateCell(5).SetCellFormula("SUM(藤桥运单扣费2!D:D)");
            sheet.CreateRow(15).CreateCell(3).SetCellValue("第三批");
            sheet.GetRow(15).CreateCell(4).SetCellFormula("COUNTA(藤桥运单扣费3!G:G)-1");
            sheet.GetRow(15).CreateCell(5).SetCellFormula("SUM(藤桥运单扣费3!D:D)");
            sheet.CreateRow(16).CreateCell(3).SetCellValue("第四批");
            sheet.GetRow(16).CreateCell(4).SetCellFormula("COUNTA(藤桥运单扣费4!G:G,藤桥集包!C:C)-2");
            sheet.GetRow(16).CreateCell(5).SetCellFormula("SUM(藤桥运单扣费4!D:D,藤桥集包!D:D)");
            sheet.CreateRow(17).CreateCell(3).SetCellValue("合计：");
            sheet.GetRow(17).CreateCell(4).SetCellFormula("SUM(E14:E17)");
            sheet.GetRow(17).CreateCell(5).SetCellFormula("SUM(F14:F17)");

            return sheet;
        }

        /// <summary>
        /// 根据单元格格式转换成对应c#格式
        /// </summary>
        /// <param name="myCell"></param>
        /// <param name="cell"></param>
        /// <param name="cellType"></param>
        private void CopyCell(ICell myCell, ICell cell, CellType cellType)
        {
            switch (cellType)
            {
                case CellType.Unknown:
                    myCell.SetCellValue(cell.StringCellValue);
                    myCell.SetCellType(cellType);
                    break;

                case CellType.Numeric:
                    myCell.SetCellValue(cell.NumericCellValue);
                    myCell.SetCellType(cellType);
                    break;

                case CellType.String:
                    myCell.SetCellValue(cell.StringCellValue);
                    myCell.SetCellType(cellType);
                    break;

                case CellType.Formula:
                    myCell.SetCellValue(cell.CellFormula);
                    myCell.SetCellType(cellType);
                    break;

                case CellType.Blank:
                    myCell.SetCellType(cellType);
                    break;

                case CellType.Boolean:
                    myCell.SetCellValue(cell.BooleanCellValue);
                    myCell.SetCellType(cellType);
                    break;

                case CellType.Error:
                    myCell.SetCellValue((double)cell.ErrorCellValue);
                    myCell.SetCellType(cellType);
                    break;

                default:
                    myCell.SetCellValue(cell.StringCellValue);
                    myCell.SetCellType(cellType);
                    break;
            }
        }
        /// <summary>
        /// 根据sql语句从数据库查询数据
        /// </summary>
        /// <param name="sheet">传入返回的表格</param>
        /// <param name="sqlstr">sql语句</param>
        /// <param name="starttime">参数开始时间</param>
        /// <param name="endtime">参数结束时间</param>
        /// <returns></returns>
        public ISheet GetSheet(ISheet sheet, string sqlstr, DateTime starttime, DateTime endtime)
        {
            SqlDataReader reader;
            SqlParameter[] sp = new SqlParameter[] { new SqlParameter("@starttime", starttime), new SqlParameter("@endtime", endtime) };
            CommandType commandType = CommandType.StoredProcedure;
           
            if (sheet.SheetName== "未分类站点")
            {
                commandType = CommandType.Text;
            }


            reader = SqlHelper.ExecuteReader(sqlstr, commandType, sp);

            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < reader.FieldCount; i++)
            {
                row.CreateCell(i).SetCellValue(reader.GetName(i));
            }
            if (reader.HasRows)
            {
                for (int j = 1; reader.Read(); j++)
                {
                    IRow row2 = sheet.CreateRow(j);
                    for (int k = 0; k < reader.FieldCount; k++)
                    {
                        object obj2 = reader[k];
                        ICell cell = row2.CreateCell(k);
                        if (reader.GetName(k) == "结算金额")
                        {//包号扣费，公式计算
                            if ((sheet.SheetName == "包号扣费"))
                            {
                                cell.SetCellFormula(string.Format("ROUND(F{0}*H{0}/SUMIFS(H:H,J:J,J{0}),2)", j + 1));
                            }
                            else
                            {
                                cell.SetCellValue(Convert.ToDouble(obj2 ?? "0"));
                            }
                        }
                        else if ((reader.GetName(k) == "重量") || (reader.GetName(k) == "扣费金额") || (reader.GetName(k) == "金额"))
                        {
                            cell.SetCellValue(Convert.ToDouble(obj2 ==DBNull.Value? "0":obj2));
                        }
                        else
                        {
                            cell.SetCellValue(obj2.ToString());
                        }
                    }
                }
            }
            return sheet;
        }
        /// <summary>
        /// 制作汇总表
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public ISheet GetSummarySheet(ISheet sheet)
        {
            IRow row = sheet.CreateRow(1);
            row.CreateCell(5).SetCellValue("上传");
            row.CreateCell(6).SetCellValue("不上传");
            row.CreateCell(7).SetCellValue("合计");
            row.CreateCell(8).SetCellValue("系统实际扣费");
            row.CreateCell(9).SetCellValue("差异");
            row = sheet.CreateRow(2);
            row.CreateCell(4).SetCellValue("系统扣费");
            row.CreateCell(5).SetCellFormula("SUM(运单扣费!D:D,包号扣费!D:D,非匹配数据!D:D,刷单扣费!D:D,'001运单扣费'!D:D,取消上传!D:D)");
            row.CreateCell(6).SetCellFormula("SUM(应收余额数据!D:D,包号费!D:D)");
            row.CreateCell(7).SetCellFormula("F3+G3");
            row.CreateCell(9).SetCellFormula("H3-I3");
            row = sheet.CreateRow(3);
            row.CreateCell(4).SetCellValue("结算扣费");
            row.CreateCell(5).SetCellFormula("SUM(集包收费!D:D,集包费取消!D:D,'001集包收费'!D:D)");
            row = sheet.CreateRow(4);
            row.CreateCell(4).SetCellValue("合计");
            row.CreateCell(5).SetCellFormula("SUM(F3:F4)");
            row.CreateCell(6).SetCellValue("上传数量");
            row = sheet.CreateRow(5);
            row.CreateCell(4).SetCellValue("运单扣费");
            row.CreateCell(5).SetCellFormula("SUM(运单扣费!D:D)");
            row.CreateCell(6).SetCellFormula("COUNTA(运单扣费!C:C)-1");
            row = sheet.CreateRow(6);
            row.CreateCell(4).SetCellValue("包号扣费上传");
            //row.CreateCell(5).SetCellFormula("SUM(包号扣费!D:D)");
            //row.CreateCell(6).SetCellFormula("COUNTA(包号扣费!C:C)-1");
            row = sheet.CreateRow(7);
            row.CreateCell(4).SetCellValue("其他上传扣费");
            row.CreateCell(5).SetCellFormula("SUM('001运单扣费'!D:D,取消上传!D:D,非匹配数据!D:D,刷单扣费!D:D,集包收费!D:D,'001集包收费'!D:D,集包费取消!D:D,包号扣费!D:D)");
            row.CreateCell(6).SetCellFormula("COUNTA('001运单扣费'!C:C,取消上传!C:C,非匹配数据!C:C,刷单扣费!B:B,集包收费!C:C,'001集包收费'!C:C,集包费取消!C:C,包号扣费!C:C)-8");
            row = sheet.CreateRow(8);
            row.CreateCell(4).SetCellValue("合计");
            row.CreateCell(5).SetCellFormula("SUM(F6:F8)");
            row.CreateCell(6).SetCellFormula("SUM(G6:G8)");
            row.CreateCell(8).SetCellValue("注意：1.未匹配数据部分还有一些不上传");
            sheet.CreateRow(9).CreateCell(8).SetCellValue(" 2.包号扣费部分站点数据未知，需要补齐");
            sheet.CreateRow(10).CreateCell(8).SetCellValue(" 3.包号扣费 非航空3KG以下转到取消上传，站点名改成鹅湖分部");
            sheet.CreateRow(11).CreateCell(8).SetCellValue(DateTime.Today.AddDays(-1.0).ToShortDateString());
            return sheet;
        }
        /// <summary>
        /// 合并表格
        /// </summary>
        /// <param name="myExcel">表格数据model</param>
        /// <param name="mySheet">传入和返回的工作薄</param>
        /// <param name="dataSouceFileNames">源数据路径</param>
        /// <param name="isAddFilename">是否加标题</param>
        /// <returns></returns>
        public ISheet MergeExcel(MyExcel myExcel, ISheet mySheet, string dataSouceFileNames, bool isAddFilename)
        {
            int count = myExcel.AddFileNames.Count;
            int num2 = myExcel.AddFileNames.Count;
            int rownum = myExcel.SouceStartRow - 1;
            using (FileStream stream = File.OpenRead(dataSouceFileNames))
            {
                ISheet sheetAt = WorkbookFactory.Create(stream).GetSheetAt(0);
                int lastRowNum = sheetAt.LastRowNum;
                IRow row = sheetAt.GetRow(rownum);
                IRow row2 = mySheet.CreateRow(myExcel.CurrentRow);
                int lastCellNum = row.LastCellNum;
                string str = sheetAt.GetRow(0).GetCell(1).ToString();
                if (isAddFilename)
                {
                    if (num2 > 0)
                    {
                        for (int k = 0; k < myExcel.AddFileNames.Count; k++)
                        {
                            row2.CreateCell(k).SetCellValue(myExcel.AddFileNames[k]);
                        }
                    }
                    for (int j = 0; j < lastCellNum; j++)
                    {
                        if (row.GetCell(j).StringCellValue == "运单编号")
                        {
                            row2.CreateCell(count + j).SetCellValue("运单号");
                        }
                        else if (row.GetCell(j).StringCellValue == "第一次出件时间")
                        {
                            row2.CreateCell(count + j).SetCellValue("入网日期");
                        }
                        else
                        {
                            this.CopyCell(row2.CreateCell(count + j), row.GetCell(j), row.GetCell(j).CellType);
                        }
                    }
                    myExcel.CurrentRow++;
                }
                rownum++;
                for (int i = rownum; i <= lastRowNum; i++)
                {
                    row = sheetAt.GetRow(rownum);
                    row2 = mySheet.CreateRow(myExcel.CurrentRow);
                    if (count > 0)
                    {
                        row2.CreateCell(0).SetCellValue(str);
                        row2.CreateCell(1).SetCellFormula(string.Format("Max(J{0}:N{0})", myExcel.CurrentRow + 1));
                    }
                    for (int j = 0; j < lastCellNum; j++)
                    {
                        ICell myCell = row2.CreateCell(count + j);
                        this.CopyCell(myCell, row.GetCell(j), row.GetCell(j).CellType);
                    }
                    rownum++;
                    myExcel.CurrentRow++;
                }
            }
            return mySheet;
        }
        /// <summary>
        /// 上传派件数据
        /// </summary>
        /// <param name="partlist"></param>
        /// <returns></returns>
        public int UploadPartstoDataBase(List<Parts> partlist)
        {
            int resultNum = 0;
            string sql = "insert into Parts values( @id, @scanSite, @scanTime, @scanPeople, @recordtime, @worker);";
            foreach (Parts model in partlist)
            {
                if (model != null)
                {
                    SqlParameter[] sp = new SqlParameter[]
                    {   new SqlParameter("@id", SqlDbType.NVarChar, 50){ Value = model.ID == null?((object)DBNull.Value) : ((object)model.ID)
                    },
                        new SqlParameter("@scanSite", SqlDbType.NVarChar, 50){Value = (model.ScanSite == null) ? ((object)DBNull.Value) : ((object)model.ScanSite)},
                        new SqlParameter("@scanTime", SqlDbType.DateTime){Value = model.ScanTime},
                        new SqlParameter("@scanPeople", SqlDbType.NVarChar, 50){Value = model.ScanPeople==null?(object)DBNull.Value:model.ScanPeople},
                        new SqlParameter("@recordtime", SqlDbType.DateTime){Value = model.Recordtime },
                        new SqlParameter("@worker", SqlDbType.NVarChar, 50){Value = model.Worker==null?(object)DBNull.Value:model.Worker}


                };

                    //SqlParameter[] sp = parameterArray1;
                    resultNum += SqlHelper.ExecuteNonQuery(sql, CommandType.Text, sp);
                }

            }
            return resultNum;




        }

        /// <summary>
        /// 更新重量
        /// </summary>
        /// <param name="dateTime"></param>
        /// <returns></returns>
        public int UpdateData(DateTime startTime, DateTime endTime)
        {
            string sql = "pro_checkHeight";

            SqlParameter[] sp = new SqlParameter[]
           { new SqlParameter("@starttime", SqlDbType.DateTime) { Value=startTime},
                new SqlParameter("@endtime", SqlDbType.DateTime) { Value = endTime } };
            try
            {//不知道什么原因，第一次执行就是连接不上数据库，所以返回-1先，后边再次执行
                return SqlHelper.ExecuteNonQuery(sql, CommandType.StoredProcedure, sp);

            }
            catch (Exception)
            {
                return -1;
            }
           
           

           
        }
        /// <summary>
        /// 集包数据上传
        /// </summary>
        /// <param name="collectlist"></param>
        /// <returns></returns>
        public int UploadCollectBagtoDataBase(List<CollectPackBag> collectlist)
        {
            int resultNum=0;
            string sql = "insert into collecbags values( @ScanSite, @ScanType, @BagID, @ID, @ScanPeople, @ScanTime, @RecordTime, @Weight, @DestinationProvince, @DestinationCity, @Site);";
            foreach (CollectPackBag model in collectlist)
            {
                if (model != null)
                {
                    SqlParameter[] sp = new SqlParameter[]
                    {   new SqlParameter("@ScanSite", SqlDbType.NVarChar, 50){ Value = (model.ScanSite ==null) ? ((object)DBNull.Value) : ((object)model.ScanSite)},
                        new SqlParameter("@ScanType", SqlDbType.NVarChar, 50){Value = (model.ScanType == null) ? ((object)DBNull.Value) : ((object)model.ScanType)},
                        new SqlParameter("@BagID", SqlDbType.NVarChar, 50){Value = model.BagID},
                        new SqlParameter("@ID", SqlDbType.NVarChar, 50){Value = model.ID},
                        new SqlParameter("@ScanPeople", SqlDbType.NVarChar, 50){Value = (model.ScanPeople == null) ? ((object)DBNull.Value) : ((object)model.ScanPeople)},
                        new SqlParameter("@ScanTime", SqlDbType.DateTime){ Value = model.ScanTime},
                        new SqlParameter("@RecordTime", SqlDbType.NVarChar, 50){Value = model.RecordTime},
                        new SqlParameter("@Weight", SqlDbType.Float){Value = !model.Weight.HasValue ? 0.0 : model.Weight},
                        new SqlParameter("@DestinationProvince", SqlDbType.NVarChar, 50){Value = (model.DestinationProvince == null) ? ((object)DBNull.Value) : ((object)model.DestinationProvince)},
                        new SqlParameter("@DestinationCity", SqlDbType.NVarChar, 50){Value=(model.DestinationCity == null) ? ((object)DBNull.Value):((object)model.DestinationCity)},
                        new SqlParameter("@Site", SqlDbType.NVarChar, 50){ Value = (model.Site == null) ? ((object)DBNull.Value) : ((object)model.Site)}
                };

                    //SqlParameter[] sp = parameterArray1;
                    resultNum+= SqlHelper.ExecuteNonQuery(sql, CommandType.Text, sp);
                }

            }
            return resultNum;
        }
        /// <summary>
        /// 扣费数据上传
        /// </summary>
        /// <param name="costlist"></param>
        /// <returns></returns>
        public int UploadCosttoDataBase(List<Cost> costlist)
        {
            int resultRows = 0;
            string sql = "insert into cost values( @costID, @costtype, @time, @costnum, @amount, @amounttype, @remark);";
            foreach (Cost model in costlist)
            {
                if (model != null)
                {
                    SqlParameter[] sp = new SqlParameter[]
                    {    new SqlParameter("@costID", SqlDbType.NVarChar, 50){Value = (model.CostID ==         null) ? ((object)DBNull.Value) : ((object)model.CostID)},
                         new SqlParameter("@costtype", SqlDbType.NVarChar, 50){Value = model.CostType},
                         new SqlParameter("@time", SqlDbType.NVarChar, 50){Value = model.CostTime},
                         new SqlParameter("@costnum", SqlDbType.NVarChar, 50){Value = model.CostNum},
                         new SqlParameter("@amount", SqlDbType.NVarChar, 50){ Value = model.CostAmount},
                         new SqlParameter("@amounttype", SqlDbType.NVarChar, 50){Value =                model.CostAmountType},
                         new SqlParameter("@remark", SqlDbType.NVarChar, 100){Value = (model.Remarks ==    null) ? ((object)DBNull.Value) : (model.Remarks)}
                };

                    resultRows +=SqlHelper.ExecuteNonQuery(sql, CommandType.Text, sp);
                }
            }
            return resultRows;
        }
        /// <summary>
        /// 订单数据上传
        /// </summary>
        /// <param name="costlist"></param>
        /// <returns></returns>
        public int UploadCustomertoDataBase(List<Customer> costlist)
        {
            int resultRows = 0;
            string sql = "insert into customer values( @Date, @ID, @Address1, @Address2, @Address3, @Weight, @Site, @WeightFenbo, @WeightJisancan, @WeightYiji, @WeightErji,@WeightJipao);";
            foreach (Customer model in costlist)
            {
                if (model != null)
                {
                    SqlParameter[] sp = new SqlParameter[] {
                        new SqlParameter("@Date", SqlDbType.DateTime){ Value = model.Date}, new SqlParameter("@ID", SqlDbType.NVarChar, 50){ Value = model.ID },
                        new SqlParameter("@Address1", SqlDbType.NVarChar, 50){ Value = (model.Address1 == null) ? ((object)DBNull.Value) : ((object)model.Address1) },
                        new SqlParameter("@Address2", SqlDbType.NVarChar, 50){ Value = (model.Address2 == null) ? ((object)DBNull.Value) : ((object)model.Address2) },
                        new SqlParameter("@Address3", SqlDbType.NVarChar, 50){ Value = (model.Address3 == null) ? ((object)DBNull.Value) : ((object)model.Address3) },
                        new SqlParameter("@Weight", SqlDbType.Float) { Value = model.Weight },
                        new SqlParameter("@Site", SqlDbType.NVarChar, 50){ Value = (model.Site == null) ? ((object)DBNull.Value) : ((object)model.Site)},
                        new SqlParameter("@WeightFenbo", SqlDbType.Float){ Value = model.WeightFenbo },
                        new SqlParameter("@WeightJisancan", SqlDbType.Float){Value=model.WeightJisancan },
                        new SqlParameter("@WeightYiji", SqlDbType.Float){ Value = model.WeightYiji },
                        new SqlParameter("@WeightErji", SqlDbType.Float){ Value = model.WeightErji },
                        new SqlParameter("@WeightJipao", SqlDbType.Float){ Value = model.WeightJipao }
                };

                    //SqlParameter[] sp = parameterArray1;
                resultRows +=   SqlHelper.ExecuteNonQuery(sql, CommandType.Text, sp);
                }
            }
            return resultRows;
        }
    }
}






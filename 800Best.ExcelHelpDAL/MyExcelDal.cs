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

            public ISheet ChangeExcel(ISheet sheet)
            {
                IRow row = sheet.GetRow(0);
                if ((row != null) && (row.GetCell(0) != null))
                {
                    int lastRowNum = sheet.LastRowNum;
                    int lastCellNum = row.LastCellNum;
                    int cellnum = -1;
                    int num4 = -1;
                    int num5 = 1;
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
                    if ((num4 == -1) || (cellnum == -1))
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

            private void ChangeExcelType(ICell cell, double numericCellValue)
            {
                string stringCellValue = cell.StringCellValue;
                if (stringCellValue != null)
                {
                    switch (stringCellValue)
                    {
                        case "站点留仓率罚款":
                         
                                if (cell.Row.GetCell(0) != null)
                                {
                                    cell.Row.GetCell(4).SetCellValue(cell.Row.GetCell(0).StringCellValue);
                                    cell.Row.GetCell(0).SetCellValue("");
                                }
                                break;
                        case "寄件派费":
                          
                                if (numericCellValue > 0.0)
                                {
                                    cell.SetCellValue("寄件派费调整");
                                }
                                break;
                         

                        case "代转件费":
                          
                                if (numericCellValue > 0.0)
                                {
                                    cell.SetCellValue("代转件费调整");
                                }
                            
                            break;

                        case "中转费调整":
                          
                                if (numericCellValue > 0.0)
                                {
                                    cell.SetCellValue("中转费取消");
                                }
                                break;
                          

                        case "计重收费调整":
                          
                                if (numericCellValue < 0.0)
                                {
                                    cell.SetCellValue("计重收费");
                                }
                                break;
                           
                        case "违禁品罚款":
                          
                                cell.SetCellValue("航空违禁品罚款");
                                break;
                           

                        case "扣有偿中转调整":
                          
                                if (numericCellValue > 0.0)
                                {
                                    cell.SetCellValue("扣有偿中转取消");
                                }
                                break;
                           

                        case "扫描费":
                          
                                cell.SetCellValue("扫描费调整");
                                break;
                          

                        case "短信服务费":
                            
                                if (cell.Row.GetCell(0) != null)
                                {
                                    cell.Row.GetCell(0).SetCellValue("");
                                }
                                break;
                         

                        case "保价手续费":
                          
                                cell.SetCellValue("手续费");
                                break;
                           

                        case "批货大货费":
                           
                                cell.SetCellValue("大货费");
                                break;
                           
                        case "大货手续费":
                          
                                if (numericCellValue > 0.0)
                                {
                                    cell.SetCellValue("大货费调整");
                                }
                                break;
                         

                        case "错集率罚款":
                          
                                if (cell.Row.GetCell(0)!=null)
                                {
                                    cell.Row.GetCell(4).SetCellValue(cell.Row.GetCell(0).StringCellValue);
                                    cell.Row.GetCell(0).SetCellValue("");
                                }
                                break;
                          
                    }
                }
            }

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

            public ISheet GetSheet(ISheet sheet, string sqlstr, DateTime starttime, DateTime endtime)
            {
                SqlParameter[] sp = new SqlParameter[] { new SqlParameter("@starttime", starttime), new SqlParameter("@endtime", endtime) };
                SqlDataReader reader = SqlHelper.ExecuteReader(sqlstr, CommandType.Text, sp);
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
                            {
                                if ((sheet.SheetName == "包号扣费") && ((reader[k - 1].ToString() == "中转费") || (reader[k - 1].ToString() == "中转费调整")))
                                {
                                    string formula = string.Format("ROUND(F{0}/SUMIFS(H:H,C:C,C{0},E:E,E{0})*H{0},2)", j + 1);
                                    cell.SetCellFormula(formula);
                                }
                                else
                                {
                                    cell.SetCellValue(Convert.ToDouble(obj2 ?? "0"));
                                }
                            }
                            else if ((reader.GetName(k) == "重量") || (reader.GetName(k) == "金额"))
                            {
                                cell.SetCellValue(Convert.ToDouble(obj2 ?? "0"));
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
                row.CreateCell(5).SetCellFormula("SUM(包号扣费!D:D)");
                row.CreateCell(6).SetCellFormula("COUNTA(包号扣费!C:C)-1");
                row = sheet.CreateRow(7);
                row.CreateCell(4).SetCellValue("其他上传扣费");
                row.CreateCell(5).SetCellFormula("F5-F6-F7");
                row.CreateCell(6).SetCellFormula("COUNTA('001运单扣费'!C:C,取消上传!C:C,非匹配数据!C:C,刷单扣费!B:B,集包收费!C:C,'001集包收费'!C:C,集包费取消!C:C)-7");
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

            public int UpdateData(DateTime dateTime)
            {
                string sql = "pro_checkHeight";
                SqlParameter[] sp = new SqlParameter[1];
                SqlParameter parameter1 = new SqlParameter("@starttime", SqlDbType.DateTime)
                {
                    Value = dateTime
                };
                sp[0] = parameter1;
                return SqlHelper.ExecuteNonQuery(sql, CommandType.StoredProcedure, sp);
            }

            public bool UploadCollectBagtoDataBase(List<CollectPackBag> collectlist)
            {
                string sql = "insert into collecbags values( @ScanSite, @ScanType, @BagID, @ID, @ScanPeople, @ScanTime, @RecordTime, @Weight, @DestinationProvince, @DestinationCity, @Site);";
                foreach (CollectPackBag model in collectlist)
                {
                    if (model != null)
                    {
                        SqlParameter[] parameterArray1 = new SqlParameter[11];
                        SqlParameter parameter1 = new SqlParameter("@ScanSite", SqlDbType.NVarChar, 50)
                        {
                            Value = (model.ScanSite == null) ? ((object)DBNull.Value) : ((object)model.ScanSite)
                        };
                        parameterArray1[0] = parameter1;
                        SqlParameter parameter2 = new SqlParameter("@ScanType", SqlDbType.NVarChar, 50)
                        {
                            Value = (model.ScanType == null) ? ((object)DBNull.Value) : ((object)model.ScanType)
                        };
                        parameterArray1[1] = parameter2;
                        SqlParameter parameter3 = new SqlParameter("@BagID", SqlDbType.NVarChar, 50)
                        {
                            Value = model.BagID
                        };
                        parameterArray1[2] = parameter3;
                        SqlParameter parameter4 = new SqlParameter("@ID", SqlDbType.NVarChar, 50)
                        {
                            Value = model.ID
                        };
                        parameterArray1[3] = parameter4;
                        SqlParameter parameter5 = new SqlParameter("@ScanPeople", SqlDbType.NVarChar, 50)
                        {
                            Value = (model.ScanPeople == null) ? ((object)DBNull.Value) : ((object)model.ScanPeople)
                        };
                        parameterArray1[4] = parameter5;
                        SqlParameter parameter6 = new SqlParameter("@ScanTime", SqlDbType.DateTime)
                        {
                            Value = model.ScanTime
                        };
                        parameterArray1[5] = parameter6;
                        SqlParameter parameter7 = new SqlParameter("@RecordTime", SqlDbType.NVarChar, 50)
                        {
                            Value = model.RecordTime
                        };
                        parameterArray1[6] = parameter7;
                        SqlParameter parameter8 = new SqlParameter("@Weight", SqlDbType.Float)
                        {
                            Value = !model.Weight.HasValue ? 0.0 : model.Weight
                        };
                        parameterArray1[7] = parameter8;
                        SqlParameter parameter9 = new SqlParameter("@DestinationProvince", SqlDbType.NVarChar, 50)
                        {
                            Value = (model.DestinationProvince == null) ? ((object)DBNull.Value) : ((object)model.DestinationProvince)
                        };
                        parameterArray1[8] = parameter9;
                        SqlParameter parameter10 = new SqlParameter("@DestinationCity", SqlDbType.NVarChar, 50)
                        {
                            Value = (model.DestinationCity == null) ? ((object)DBNull.Value) : ((object)model.DestinationCity)
                        };
                        parameterArray1[9] = parameter10;
                        SqlParameter parameter11 = new SqlParameter("@Site", SqlDbType.NVarChar, 50)
                        {
                            Value = (model.Site == null) ? ((object)DBNull.Value) : ((object)model.Site)
                        };
                        parameterArray1[10] = parameter11;
                        SqlParameter[] sp = parameterArray1;
                        SqlHelper.ExecuteNonQuery(sql, CommandType.Text, sp);
                    }
                }
                return true;
            }

            public bool UploadCosttoDataBase(List<Cost> costlist)
            {
                string sql = "insert into cost values( @costID, @costtype, @time, @costnum, @amount, @amounttype, @remark);";
                foreach (Cost model in costlist)
                {
                    if (model != null)
                    {
                        SqlParameter[] parameterArray1 = new SqlParameter[7];
                        SqlParameter parameter1 = new SqlParameter("@costID", SqlDbType.NVarChar, 50)
                        {
                            Value = (model.CostID == null) ? ((object)DBNull.Value) : ((object)model.CostID)
                        };
                        parameterArray1[0] = parameter1;
                        SqlParameter parameter2 = new SqlParameter("@costtype", SqlDbType.NVarChar, 50)
                        {
                            Value = model.CostType
                        };
                        parameterArray1[1] = parameter2;
                        SqlParameter parameter3 = new SqlParameter("@time", SqlDbType.NVarChar, 50)
                        {
                            Value = model.CostTime
                        };
                        parameterArray1[2] = parameter3;
                        SqlParameter parameter4 = new SqlParameter("@costnum", SqlDbType.NVarChar, 50)
                        {
                            Value = model.CostNum
                        };
                        parameterArray1[3] = parameter4;
                        SqlParameter parameter5 = new SqlParameter("@amount", SqlDbType.NVarChar, 50)
                        {
                            Value = model.CostAmount
                        };
                        parameterArray1[4] = parameter5;
                        SqlParameter parameter6 = new SqlParameter("@amounttype", SqlDbType.NVarChar, 50)
                        {
                            Value = model.CostAmountType
                        };
                        parameterArray1[5] = parameter6;
                        SqlParameter parameter7 = new SqlParameter("@remark", SqlDbType.NVarChar, 100)
                        {
                            Value = (model.Remarks == null) ? ((object)DBNull.Value) : ((object)model.Remarks)
                        };
                        parameterArray1[6] = parameter7;
                        SqlParameter[] sp = parameterArray1;
                        SqlHelper.ExecuteNonQuery(sql, CommandType.Text, sp);
                    }
                }
                return true;
            }

            public bool UploadCustomertoDataBase(List<Customer> costlist)
            {
                string sql = "insert into customer values( @Date, @ID, @Address1, @Address2, @Address3, @Weight, @Site, @WeightFenbo, @WeightJisancan, @WeightYiji, @WeightErji,@WeightJipao);";
                foreach (Customer model in costlist)
                {
                    if (model != null)
                    {
                        SqlParameter[] parameterArray1 = new SqlParameter[12];
                        SqlParameter parameter1 = new SqlParameter("@Date", SqlDbType.DateTime)
                        {
                            Value = model.Date
                        };
                        parameterArray1[0] = parameter1;
                        SqlParameter parameter2 = new SqlParameter("@ID", SqlDbType.NVarChar, 50)
                        {
                            Value = model.ID
                        };
                        parameterArray1[1] = parameter2;
                        SqlParameter parameter3 = new SqlParameter("@Address1", SqlDbType.NVarChar, 50)
                        {
                            Value = (model.Address1 == null) ? ((object)DBNull.Value) : ((object)model.Address1)
                        };
                        parameterArray1[2] = parameter3;
                        SqlParameter parameter4 = new SqlParameter("@Address2", SqlDbType.NVarChar, 50)
                        {
                            Value = (model.Address2 == null) ? ((object)DBNull.Value) : ((object)model.Address2)
                        };
                        parameterArray1[3] = parameter4;
                        SqlParameter parameter5 = new SqlParameter("@Address3", SqlDbType.NVarChar, 50)
                        {
                            Value = (model.Address3 == null) ? ((object)DBNull.Value) : ((object)model.Address3)
                        };
                        parameterArray1[4] = parameter5;
                        SqlParameter parameter6 = new SqlParameter("@Weight", SqlDbType.Float)
                        {
                            Value = model.Weight
                        };
                        parameterArray1[5] = parameter6;
                        SqlParameter parameter7 = new SqlParameter("@Site", SqlDbType.NVarChar, 50)
                        {
                            Value = (model.Site == null) ? ((object)DBNull.Value) : ((object)model.Site)
                        };
                        parameterArray1[6] = parameter7;
                        SqlParameter parameter8 = new SqlParameter("@WeightFenbo", SqlDbType.Float)
                        {
                            Value = model.WeightFenbo
                        };
                        parameterArray1[7] = parameter8;
                        SqlParameter parameter9 = new SqlParameter("@WeightJisancan", SqlDbType.Float)
                        {
                            Value = model.WeightJisancan
                        };
                        parameterArray1[8] = parameter9;
                        SqlParameter parameter10 = new SqlParameter("@WeightYiji", SqlDbType.Float)
                        {
                            Value = model.WeightYiji
                        };
                        parameterArray1[9] = parameter10;
                        SqlParameter parameter11 = new SqlParameter("@WeightErji", SqlDbType.Float)
                        {
                            Value = model.WeightErji
                        };
                        parameterArray1[10] = parameter11;
                        SqlParameter parameter12 = new SqlParameter("@WeightJipao", SqlDbType.Float)
                        {
                            Value = model.WeightJipao
                        };
                        parameterArray1[11] = parameter12;
                        SqlParameter[] sp = parameterArray1;
                        SqlHelper.ExecuteNonQuery(sql, CommandType.Text, sp);
                    }
                }
                return true;
            }
        }
    }






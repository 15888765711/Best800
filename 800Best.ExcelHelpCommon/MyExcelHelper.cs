using _800Best.ExcelHelpModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _800Best.ExcelHelpCommon
{
 

        public class MyExcelHelper
        {
            public static List<Customer> CustomerList(string file)
            {
                string[] filename = new string[] { "入网日期", "运单号", "电子面单打印", "电商下单", "签收目的地", "重量", "归属站点", "分拨中心称重", "集散仓称重", "一级站点称重", "二级站点称重", "计泡重量" };
                IWorkbook workbook = WorkbookFactory.Create(file);
                ISheet sheetAt = workbook.GetSheetAt(0);
                List<Customer> customerModels = new List<Customer>();
                customerModels = GetListCustomerModel(sheetAt, customerModels, filename);
                ISheet sheet = workbook.GetSheet("未客户发放");
                if (sheet != null)
                {
                    customerModels = GetListCustomerModel(sheet, customerModels, filename);
                    workbook.Close();
                }
                return customerModels;
            }

            public static List<CollectPackBag> GetCollectList(string file)
            {
                string[] strArray = new string[] { "扫描站点", "扫描类型", "包号", "运单编号", "扫描人", "扫描日期", "录入时间", "重量", "目的分拨(省)", "目的分拨(市)", "面单发放网点" };
                IWorkbook workbook = WorkbookFactory.Create(file);
                ISheet sheetAt = workbook.GetSheetAt(0);
                List<CollectPackBag> list = new List<CollectPackBag>();
                int lastRowNum = sheetAt.LastRowNum;
                int lastCellNum = sheetAt.GetRow(0).LastCellNum;
                IRow row = sheetAt.GetRow(0);
                int[] numArray = new int[strArray.Length];
                for (int i = 0; i < strArray.Length; i++)
                {
                    numArray[i] = -1;
                    for (int k = 0; k < lastCellNum; k++)
                    {
                        if (strArray[i] == row.GetCell(k).StringCellValue)
                        {
                            numArray[i] = k;
                            break;
                        }
                    }
                    if (numArray[i] == -1)
                    {
                        return null;
                    }
                }
                for (int j = 1; j <= lastRowNum; j++)
                {
                    row = sheetAt.GetRow(j);
                    CollectPackBag item = new CollectPackBag
                    {
                        ScanSite = (row.GetCell(numArray[0]) == null) ? null : row.GetCell(numArray[0]).StringCellValue,
                        ScanType = (row.GetCell(numArray[1]) == null) ? null : row.GetCell(numArray[1]).StringCellValue,
                        BagID = (row.GetCell(numArray[2]) == null) ? null : row.GetCell(numArray[2]).StringCellValue,
                        ID = (row.GetCell(numArray[3]) == null) ? null : row.GetCell(numArray[3]).StringCellValue,
                        ScanPeople = (row.GetCell(numArray[4]) == null) ? null : row.GetCell(numArray[4]).StringCellValue,
                        ScanTime = Convert.ToDateTime(row.GetCell(numArray[5]).StringCellValue),
                        RecordTime = row.GetCell(numArray[6]).StringCellValue,
                        Weight = new double?((row.GetCell(numArray[7]) == null) ? 0.2 : row.GetCell(numArray[7]).NumericCellValue),
                        DestinationProvince = (row.GetCell(numArray[8]) == null) ? null : row.GetCell(numArray[8]).StringCellValue,
                        DestinationCity = (row.GetCell(numArray[9]) == null) ? null : row.GetCell(numArray[9]).StringCellValue,
                        Site = (row.GetCell(numArray[10]) == null) ? null : row.GetCell(numArray[10]).StringCellValue
                    };
                    list.Add(item);
                }
                workbook.Close();
                return list;
            }

            public static List<Cost> GetCostList(string file)
            {
                string[] strArray = new string[] { "运单编号", "结算类型", "结算/上传日期", "结算流水号", "金额", "入账余额", "备注" };
                IWorkbook workbook = WorkbookFactory.Create(file);
                ISheet sheetAt = workbook.GetSheetAt(0);
                List<Cost> list = new List<Cost>();
                int lastRowNum = sheetAt.LastRowNum;
                int lastCellNum = sheetAt.GetRow(0).LastCellNum;
                IRow row = sheetAt.GetRow(0);
                int[] numArray = new int[strArray.Length];
                for (int i = 0; i < strArray.Length; i++)
                {
                    numArray[i] = -1;
                    for (int k = 0; k < lastCellNum; k++)
                    {
                        if (strArray[i] == row.GetCell(k).StringCellValue)
                        {
                            numArray[i] = k;
                            break;
                        }
                    }
                    if (numArray[i] == -1)
                    {
                        return null;
                    }
                }
                for (int j = 1; j <= lastRowNum; j++)
                {
                    row = sheetAt.GetRow(j);
                    Cost item = new Cost
                    {
                        CostID = (row.GetCell(numArray[0]) == null) ? null : row.GetCell(numArray[0]).StringCellValue,
                        CostType = row.GetCell(numArray[1]).StringCellValue,
                        CostTime = Convert.ToDateTime(row.GetCell(numArray[2]).StringCellValue),
                        CostNum = row.GetCell(numArray[3]).NumericCellValue.ToString(),
                        CostAmount = row.GetCell(numArray[4]).NumericCellValue,
                        CostAmountType = row.GetCell(numArray[5]).StringCellValue,
                        Remarks = (row.GetCell(numArray[6]) == null) ? null : row.GetCell(numArray[6]).StringCellValue
                    };
                    list.Add(item);
                }
                workbook.Close();
                return list;
            }

            private static List<Customer> GetListCustomerModel(ISheet sheet, List<Customer> customerModels, string[] filename)
            {
                int lastRowNum = sheet.LastRowNum;
                int lastCellNum = sheet.GetRow(0).LastCellNum;
                IRow row = sheet.GetRow(0);
                int[] numArray = new int[filename.Length];
                for (int i = 0; i < filename.Length; i++)
                {
                    numArray[i] = -1;
                    for (int k = 0; k < lastCellNum; k++)
                    {
                        if (filename[i] == row.GetCell(k).StringCellValue)
                        {
                            numArray[i] = k;
                            break;
                        }
                    }
                    if (numArray[i] == -1)
                    {
                        return null;
                    }
                }
                for (int j = 1; j <= lastRowNum; j++)
                {
                    row = sheet.GetRow(j);
                    Customer item = new Customer
                    {
                        Date = DateTime.FromOADate(row.GetCell(numArray[0]).NumericCellValue),
                        ID = row.GetCell(numArray[1]).StringCellValue,
                        Address1 = (row.GetCell(numArray[2]) == null) ? null : row.GetCell(numArray[2]).StringCellValue,
                        Address2 = (row.GetCell(numArray[3]) == null) ? null : row.GetCell(numArray[3]).StringCellValue,
                        Address3 = (row.GetCell(numArray[4]) == null) ? null : row.GetCell(numArray[4]).StringCellValue,
                        Weight = row.GetCell(numArray[5]).NumericCellValue,
                        Site = (row.GetCell(numArray[6]) == null) ? null : row.GetCell(numArray[6]).StringCellValue,
                        WeightFenbo = (row.GetCell(numArray[7]) == null) ? 0.0 : row.GetCell(numArray[7]).NumericCellValue,
                        WeightJisancan = (row.GetCell(numArray[8]) == null) ? 0.0 : row.GetCell(numArray[8]).NumericCellValue,
                        WeightYiji = (row.GetCell(numArray[9]) == null) ? 0.0 : row.GetCell(numArray[9]).NumericCellValue,
                        WeightErji = (row.GetCell(numArray[10]) == null) ? 0.0 : row.GetCell(numArray[10]).NumericCellValue,
                        WeightJipao = (row.GetCell(numArray[11]) == null) ? 0.0 : row.GetCell(numArray[11]).NumericCellValue
                    };
                    customerModels.Add(item);
                }
                return customerModels;
            }

            private delegate T GetModelDelegate<T>(List<T> listmodel, string[] strs);
        }
    }





using _800Best.ExcelHelpCommon;
using _800Best.ExcelHelpDAL;
using _800Best.ExcelHelpModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _800Best.ExcelHelpBLL
{
    public class MyExcelBll
    {

        private readonly MyExcelDal myDal = new MyExcelDal();
        private bool isAddFileName;
        /// <summary>
        /// 修改结算类型
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public bool ChangeExcel(string fileName)
        {
            try
            {
                using (FileStream stream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = WorkbookFactory.Create(stream);
                    int numberOfSheets = workbook.NumberOfSheets;
                    for (int i = 0; i < numberOfSheets; i++)
                    {
                        ISheet sheetAt = workbook.GetSheetAt(i);
                        if (sheetAt.SheetName != "汇总表")
                        {
                            sheetAt = this.myDal.ChangeExcel(sheetAt);
                        }
                    }
                    FileStream stream2 = File.Create(fileName);
                    workbook.Write(stream2);
                    stream2.Close();
                    workbook.Close();
                }
                return true;
            }
            catch (Exception exception)
            {
                MessageBox.Show("打开文件失败；请重新检查路径\r\n" + exception.Message.ToString());
                return false;
            }
        }
        /// <summary>
        /// 获取导出数据（站点修改）
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="starttime"></param>
        /// <param name="endtime"></param>
        /// <returns></returns>
        public bool GetExportData(string filename, DateTime starttime, DateTime endtime, bool isXinqiao)
        {
            IWorkbook workbook = new XSSFWorkbook();
            string[] strArray = null;
            if (isXinqiao)
            {
                strArray = new string[] { "藤桥集包", "藤桥运单扣费1", "藤桥运单扣费2", "藤桥运单扣费3", "藤桥运单扣费4", "未分类站点", "汇总表" };
            }
            else
            {
                strArray = new string[] { "运单扣费", "001运单扣费", "取消上传", "包号扣费", "非匹配数据", "刷单扣费", "集包收费", "001集包收费", "集包费取消", "应收余额数据", "包号费", "汇总表" };
            }

            //
            int length = strArray.Length;
            ISheet[] sheetArray = new ISheet[length];
            for (int i = 0; i < (length - 1); i++)
            {
                sheetArray[i] = workbook.CreateSheet(strArray[i]);
            }
            for (int j = 0; j < (length - 1); j++)
            {
                string sqlStr = this.GetSqlStr(strArray[j]);
                sheetArray[j] = this.myDal.GetSheet(sheetArray[j], sqlStr, starttime, endtime);
            }
            sheetArray[length - 1] = workbook.CreateSheet(strArray[length - 1]);
            if (isXinqiao)
            {
                sheetArray[length - 1] = this.myDal.GetXinQiaoSummarySheet(sheetArray[length - 1]);
            }
            else
            {

                sheetArray[length - 1] = this.myDal.GetSummarySheet(sheetArray[length - 1]);
            }

            workbook.Write(File.OpenWrite(filename));
            workbook.Close();
            return true;
        }
        /// <summary>
        /// 获取sql语句
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        private string GetSqlStr(string sql)
        {
            string s = sql;
            if (s != null)
            {
                switch (s)
                {
                    case "藤桥集包": return "pro_CollectBagIncome4TQ";
                    case "藤桥集包003": return "pro_CollectBag003Income4TQ";
                    case "藤桥运单扣费1":
                        return "pro_CustomerAndPartsIDCost1";
                    case "藤桥运单扣费2":
                        return "pro_CustomerAndPartsIDCost2";
                    case "藤桥运单扣费3":
                        return "pro_CustomerAndPartsIDCost3";
                    case "藤桥运单扣费4":
                        return "pro_CustomerAndPartsIDCost4";
                    case "未分类站点":
                        return " SELECT t1.CostID AS 运单编号, '' AS 开户站点, t1.CostType AS 结算类型, t1.CostAmount AS 结算金额, 1.CostTime AS 备注, t1.CostAmountType AS 入账类型, t1.CostNum AS 结算流水号,  '' AS 派件单号, 0 AS 重量, t4.Site AS 面单发放网点 FROM dbo.Cost t1 LEFT OUTER JOIN dbo.Customer t2  ON t1.CostID = t2.ID LEFT OUTER JOIN dbo.Parts t3 ON t1.CostID = t3.ID LEFT OUTER JOIN dbo.Collecbags t4 ON t1.CostID = t4.ID WHERE(t1.CostTime >= @starttime) AND(t1.CostTime < @endtime) and t2.ID is null and t3.ID is null group by t1.CostID,t2.Site,t1.CostType,t1.CostAmount,t1.CostTime,t1.CostAmountType,t1.CostNum,t3.ID,t2.Weight,t4.Site ";
                 
                    case "包号费": return "pro_BagCost";
                    case "刷单扣费": return "pro_ShuadanCost";
                    case "001运单扣费": return "pro_CustomerIDCost001";
                    case "001集包收费": return "pro_CollectBagCost001";
                    case "集包费取消": return "pro_CollectBagCostCancel";
                    case "集包收费": return "pro_CollectBagCost";
                    case "包号扣费": return "pro_SelectBagCost";
                    case "非匹配数据": return "pro_NoMatchCost";
                    case "运单扣费": return "pro_CustomerIDCost";
                    case "应收余额数据": return "pro_SelectYingshou";
                    case "取消上传": return "pro_CancelUpdate";
                    default: return null;


                }
            }
            return null;
        }
        /// <summary>
        /// 把list转换成几行的字符串
        /// </summary>
        /// <param name="failedFileNames"></param>
        /// <returns></returns>
        private string ListtoString(List<string> failedFileNames)
        {
            string str = null;
            foreach (string str2 in failedFileNames)
            {
                str = str + str2 + "\r\n";
            }
            return str;
        }
        /// <summary>
        /// 合并单元格bll逻辑
        //// </summary>
        /// <param name="myExcel"></param>
        /// <param name="souceFileNames"></param>
        public void MergeExcel(MyExcel myExcel, List<string> souceFileNames)
        {
            this.isAddFileName = true;
            IWorkbook workbook = new XSSFWorkbook();
            ISheet mySheet = workbook.CreateSheet("承包区");
            List<string> failedFileNames = new List<string>();
            foreach (string str in souceFileNames)
            {
                mySheet = this.myDal.MergeExcel(myExcel, mySheet, str, this.isAddFileName);
                if (mySheet == null)
                {
                    failedFileNames.Add(str);
                }
                if (this.isAddFileName)
                {
                    this.isAddFileName = false;
                }
                myExcel.CurrentRow = mySheet.LastRowNum + 1;
            }
            workbook.Write(File.Create(myExcel.SaveFile));
            workbook.Close();
            MessageBox.Show($"成功复制{souceFileNames.Count - failedFileNames.Count}个表，\r\n失败{ failedFileNames.Count}个表,\r\n失败表名为{ this.ListtoString(failedFileNames)}");
        }
        /// <summary>
        /// 更新重量
        /// </summary>
        /// <param name="dateTime"></param>
        /// <returns></returns>
        public int UpdateData(DateTime startTime, DateTime endTime)
        {
            return myDal.UpdateData(startTime, endTime);
        }

        /// <summary>
        /// 上传集包
        /// </summary>
        /// <param name="upLoadFiles"></param>
        /// <returns></returns>
        public int UpLoadCollectBagToDataBase(string upLoadFiles)
        {
            if (File.Exists(upLoadFiles))
            {
                List<CollectPackBag> collectList = MyExcelHelper.GetCollectList(upLoadFiles);
                if (collectList == null)
                {
                    return 0;
                }
                return this.myDal.UploadCollectBagtoDataBase(collectList);
            }
            return 0;
        }
        /// <summary>
        /// 上传s9
        /// </summary>
        /// <param name="upLoadFiles"></param>
        /// <returns></returns>
        public int UpLoadCustomerToDataBase(string upLoadFiles)
        {
            if (File.Exists(upLoadFiles))
            {
                List<Customer> costlist = MyExcelHelper.CustomerList(upLoadFiles);
                if (costlist == null)
                {
                    return 0;
                }
                return this.myDal.UploadCustomertoDataBase(costlist);
            }
            return 0;
        }
        /// <summary>
        /// 上传派件
        /// </summary>
        /// <param name="upLoadFiles"></param>
        /// <returns></returns>
        public int UpLoadPartsToDataBase(string upLoadFiles)
        {
            if (File.Exists(upLoadFiles))
            {
                List<Parts> partlist = MyExcelHelper.GetPartsList(upLoadFiles);
                if (partlist == null)
                {
                    return 0;
                }
                return this.myDal.UploadPartstoDataBase(partlist);
            }
            return 0;
        }
        /// <summary>
        /// 上传cost
        /// </summary>
        /// <param name="upLoadFiles"></param>
        /// <returns></returns>
        public int UpLoadToDataBase(string upLoadFiles)
        {
            if (File.Exists(upLoadFiles))//判断是否存在
            {
                List<Cost> costList = MyExcelHelper.GetCostList(upLoadFiles);
                if (costList == null)
                {
                    return 0;
                }
                return this.myDal.UploadCosttoDataBase(costList);
            }
            return 0;
        }
    }
}





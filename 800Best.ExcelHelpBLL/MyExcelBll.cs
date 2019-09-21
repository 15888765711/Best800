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

            public bool GetExportData(string filename, DateTime starttime, DateTime endtime)
            {
                IWorkbook workbook = new XSSFWorkbook();
                string[] strArray = new string[] { "运单扣费", "001运单扣费", "取消上传", "包号扣费", "非匹配数据", "刷单扣费", "集包收费", "001集包收费", "集包费取消", "应收余额数据", "包号费", "汇总表" };
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
                sheetArray[length - 1] = this.myDal.GetSummarySheet(sheetArray[length - 1]);
                workbook.Write(File.OpenWrite(filename));
                workbook.Close();
                return true;
            }

            private string GetSqlStr(string sql)
            {
                string s = sql;
                if (s != null)
                {
                    switch (s)
                    {
                        case "包号费":

                            return "SELECT t1.运单编号, '温州瓯海茶山二部' AS 开户站点, t1.结算类型, t1.金额 AS 结算金额, t1.[结算/上传日期] AS 备注, t1.结算流水号 FROM dbo.cost t1  WHERE(t1.[结算/上传日期] >= @starttime) AND(t1.结算类型 ='包号费' or t1.结算类型 ='代付进港集包费' or t1.结算类型 ='走件费' or t1.结算类型 ='存款' ) AND(t1.入账余额 = '可用余额') and(t1.[结算/上传日期] < @endtime) GROUP BY t1.运单编号, t1.结算类型, t1.金额, t1.结算流水号, t1.[结算/上传日期]";


                        case "刷单扣费":

                            return "SELECT  '' AS 运单编号, '温州瓯海茶山二部' AS 开户站点,t2.结算类型, ROUND(SUM(t2.金额),2) AS 结算金额, \r\n                '刷单' AS 备注 FROM  dbo.customer t1 INNER JOIN\r\n                dbo.cost t2 ON t1.运单号 = t2.运单编号 WHERE   (t2.[结算/上传日期] >= @starttime) and (t2.[结算/上传日期] < @endtime) AND (t2.入账余额 = '可用余额') AND (t1.归属站点='刷单') Group by  t2.结算类型";

                        case "001运单扣费":

                            return "SELECT  t2.运单编号, t1.归属站点 AS 开户站点,t2.结算类型, t2.金额 AS 结算金额, \r\n                t2.[结算/上传日期] AS 备注, t2.结算流水号, t1.重量,LEFT(电子面单打印,2) AS 地区 FROM  dbo.customer t1 INNER JOIN\r\n                dbo.cost t2 ON t1.运单号 = t2.运单编号 WHERE   (t2.[结算/上传日期] >= @starttime) and (t2.[结算/上传日期] < @endtime) AND (t2.入账余额 = '可用余额')AND (t1.归属站点!='刷单') AND (t1.归属站点='温州南白象001') AND (LEFT(电子面单打印,2) IN (SELECT Province.Province FROM Province)  OR t1.重量>3)";


                        case "001集包收费":

                            return "SELECT   运单号 AS 运单编号, 归属站点 AS 开户站点, '代转件费' AS 结算类型, \r\nCASE WHEN LEFT(电子面单打印,2) IN('新疆','西藏','内蒙','宁夏','青海','海南') THEN -ceiling(重量) \r\nwhen 重量<=0.5 then 0.2\r\nwhen 重量<=1 then 0\r\nwhen 重量<=3 then -0.5\r\nELSE round(重量 * (- 0.1), 2) END AS 结算金额,\r\n 入网日期 AS 备注, 重量,LEFT(电子面单打印,2) AS 地区 \r\n FROM  dbo.customer  WHERE (入网日期 >= @starttime) and (入网日期 < @endtime) AND (归属站点='温州南白象001') ";


                        case "集包费取消":

                            return "SELECT   t2.运单编号, t1.归属站点 AS 开户站点, '代扣进港集包费取消' AS 结算类型,   -t2.金额 AS 结算金额, t2.[结算/上传日期] AS 备注, t2.结算流水号, t1.重量 FROM  dbo.customer  t1 INNER JOIN  dbo.cost t2 ON t1.运单号 = t2.运单编号 WHERE   (t2.[结算/上传日期] >= @starttime) and (t2.[结算/上传日期] < @endtime) AND (t2.结算类型 ='代扣进港集包费' ) AND (t2.入账余额 = '可用余额')";

                        case "集包收费":

                            return "SELECT   运单号 AS 运单编号, 归属站点 AS 开户站点, \r\n                CASE WHEN 重量 > 3 THEN '计重收费' ELSE '代集包费' END AS 结算类型, \r\n                CASE WHEN 重量 > 3 THEN round(重量 * (- 0.1), 2) ELSE - 0.35 END AS 结算金额, 入网日期 AS 备注, 重量,LEFT(电子面单打印,2) AS 地区 FROM  dbo.customer WHERE (入网日期 >= @starttime) and (入网日期 < @endtime) AND (归属站点!='温州南白象001') AND (归属站点!='刷单')";

                        case "包号扣费":

                            return "SELECT   t1.运单编号, t3.归属站点 AS 开户站点, t2.结算类型, \r\n                CASE t2.结算类型 WHEN '扫描费' THEN - 0.07 WHEN '寄件派费' THEN - 0.2 WHEN '中转费' THEN t2.金额 ELSE 0 END AS 结算金额, t1.包号 AS 备注, t2.金额, t1.面单发放网点, CASE t1.重量 WHEN 0 THEN 0.2 ELSE t1.重量 END AS 重量 ,LEFT(t3.电子面单打印,2) AS 地址 FROM dbo.collecbags AS t1 INNER JOIN dbo.cost AS t2 ON t1.包号 = t2.运单编号 LEFT OUTER JOIN\r\n                dbo.customer AS t3 ON t1.运单编号 = t3.运单号 WHERE(t2.[结算/上传日期] >= @starttime) AND(t2.结算类型 <> '包号费') AND(t2.入账余额 = '可用余额') and (t2.[结算/上传日期] < @endtime) GROUP BY t1.包号, t1.运单编号, t2.金额, t2.结算类型, t1.重量, t1.面单发放网点, t3.归属站点, LEFT(t3.电子面单打印,2)";


                        case "非匹配数据":

                            return "SELECT   t2.运单编号, '温州瓯海茶山二部' AS 开户站点, t2.结算类型, t2.金额 AS 结算金额,  t2.[结算/上传日期] AS 备注, t2.结算流水号,t3.面单发放网点 FROM dbo.cost t2 LEFT OUTER JOIN dbo.allId t1 ON t2.运单编号 = t1.运单号 LEFT OUTER JOIN collecbags t3 on t2.运单编号=t3.运单编号  WHERE   (t2.[结算/上传日期] >= @starttime) and (t2.[结算/上传日期] < @endtime) AND (t1.运单号 IS NULL) AND (t2.入账余额 = '可用余额') AND(t2.结算类型!='代付进港集包费') AND(t2.结算类型!='包号费') AND(t2.结算类型!='走件费') AND(t2.结算类型!='存款')";
                        case "运单扣费":

                            return "SELECT  t2.运单编号, t1.归属站点 AS 开户站点,t2.结算类型, t2.金额 AS 结算金额, \r\n                t2.[结算/上传日期] AS 备注, t2.结算流水号, t1.重量,LEFT(t1.电子面单打印,2) AS 地区 FROM  dbo.customer t1 INNER JOIN\r\n                dbo.cost t2 ON t1.运单号 = t2.运单编号 WHERE   (t2.[结算/上传日期] >= @starttime) and (t2.[结算/上传日期] < @endtime) AND (t2.入账余额 = '可用余额')AND (t1.归属站点!='刷单') AND (t1.归属站点!='温州南白象001') AND(t2.结算类型 !='代付进港集包费') ";

                        case "应收余额数据":

                            return "SELECT   运单编号, '温州瓯海茶山二部' AS 开户站点, 结算类型, 金额 AS 结算金额, [结算/上传日期] AS 备注, 结算流水号 FROM dbo.cost WHERE   ([结算/上传日期] >= @starttime) and ([结算/上传日期] < @endtime) AND (入账余额 = '应收余额')";

                        case "取消上传":
                            return "SELECT  t2.运单编号, case when t2.结算类型 in ( select  tb001.结算类型 from tb001) then '温州瓯海鹅湖分部'\r\n      else '温州南白象001' end  AS 开户站点,t2.结算类型, t2.金额 AS 结算金额, \r\n                t2.[结算/上传日期] AS 备注, t2.结算流水号, '' as 面单发放网点, t1.重量,LEFT(电子面单打印,2) AS 地区 FROM  dbo.customer t1 INNER JOIN\r\n                dbo.cost t2 ON t1.运单号 = t2.运单编号 WHERE   (t2.[结算/上传日期] >= @starttime) and (t2.[结算/上传日期] < @endtime) AND (t2.入账余额 = '可用余额')AND (t1.归属站点!='刷单') AND (t1.归属站点='温州南白象001') AND (LEFT(电子面单打印,2) NOT IN (SELECT Province.Province FROM Province) OR LEFT(电子面单打印,2)  IS NULL)  AND (t1.重量<=3)";
                        default:
                            return null;


                    }
                }
                return null;
            }

            private string ListtoString(List<string> failedFileNames)
            {
                string str = null;
                foreach (string str2 in failedFileNames)
                {
                    str = str + str2 + "\r\n";
                }
                return str;
            }

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
                MessageBox.Show($"成功复制{souceFileNames.Count - failedFileNames.Count}个表，失败{ failedFileNames.Count}个表,失败表名为{ this.ListtoString(failedFileNames)}");
            }

            public bool UpdateData(DateTime dateTime) =>
                (this.myDal.UpdateData(dateTime) > 0);

            public bool UpLoadCollectBagToDataBase(string upLoadFiles)
            {
                if (File.Exists(upLoadFiles))
                {
                    List<CollectPackBag> collectList = MyExcelHelper.GetCollectList(upLoadFiles);
                    if (collectList == null)
                    {
                        return false;
                    }
                    return this.myDal.UploadCollectBagtoDataBase(collectList);
                }
                return false;
            }

            public bool UpLoadCustomerToDataBase(string upLoadFiles)
            {
                if (File.Exists(upLoadFiles))
                {
                    List<Customer> costlist = MyExcelHelper.CustomerList(upLoadFiles);
                    if (costlist == null)
                    {
                        return false;
                    }
                    return this.myDal.UploadCustomertoDataBase(costlist);
                }
                return false;
            }

            public bool UpLoadToDataBase(string upLoadFiles)
            {
                if (File.Exists(upLoadFiles))//判断是否存在
                {
                    List<Cost> costList = MyExcelHelper.GetCostList(upLoadFiles);
                    if (costList == null)
                    {
                        return false;
                    }
                    return this.myDal.UploadCosttoDataBase(costList);
                }
                return false;
            }
        }
    }





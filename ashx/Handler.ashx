<%@ WebHandler Language="C#" Class="Handler" %>

using System;
using System.Web;
using LitJson;
using System.Data;
using System.Text;
using OfficeOpenXml;
using System.IO;
using System.Runtime;
using Rsdn.Framework.Data;
using System.Drawing;
using System.Diagnostics;

public class Handler : IHttpHandler {
    
    public void ProcessRequest(HttpContext context)
        {
           
            JsonData data = new JsonData();
            System.IO.Stream s = System.Web.HttpContext.Current.Request.InputStream;
            BinaryReader reader = new BinaryReader(s);
            string fileclass = "";
            try
            {
                for (int i = 0; i < 2; i++)
                {
                    fileclass += reader.ReadByte().ToString();
                }
                //confirm your import file is excel
                if(fileclass=="4545")
                {
                    string fileType = fileclass;
                    DataTable dt = new DataTable();
                    //transfer your excel into datatable
                    using (ExcelPackage pck = new ExcelPackage(s))
                    {
                        ExcelWorksheet ws = pck.Workbook.Worksheets[1];
                        int minColumnNum = ws.Dimension.Start.Column;//工作区开始列
                        int maxColumnNum = ws.Dimension.End.Column; //工作区结束列
                        int minRowNum = ws.Dimension.Start.Row; //工作区开始行号
                        int maxRowNum = ws.Dimension.End.Row; //工作区结束行号
                        DataColumn vC;
                        //column name
                        string[] strcolum = { "school", "xq", "building", "EquType", "equName", "devicesn", "status", "remark", "usingDate" };
                        foreach (string i in strcolum)
                        {
                            vC = new DataColumn(i, typeof(string));
                            dt.Columns.Add(vC);
                        }
                        if (maxRowNum > 200)
                        {
                            maxRowNum = 200;
                        }
                        for (int n = 2; n <= maxRowNum; n++)
                        {
                            DataRow vRow = dt.NewRow();
                            for (int m = 1; m <= maxColumnNum; m++)
                            {
                                vRow[m - 1] = ws.Cells[n, m].Value;
                            }
                            dt.Rows.Add(vRow);
                        }
                    }

                    if (dealdata.insertIntodb(dt)!=true)
                    {
                        data["errcode"] = 3;
                        data["error"] = "数据导入失败,请确保所填数据的正确性！";
                    }else{
                        data["errcode"] = 0;
                        data["message"] = "import excel success"; }
                }
            }
            catch (Exception ex)
            {
                data["errcode"] = "1";
                data["error"] = "系统错误：" + ex.Message;
            }
            finally
            {
                context.Response.Write(data.ToJson());
            }
        }
    public bool IsReusable {
        get {
            return false;
        }
    }

}
using LitJson;
using Rsdn.Framework.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

/// <summary>
/// dealdata 的摘要说明
/// </summary>
public class dealdata
{
    public dealdata()
    {
        //
        // TODO: 在此处添加构造函数逻辑
        //

    }
    public static bool insertIntodb(DataTable dt)
    {
        try
        {
            using (DbManager dbm = new DbManager("name"))
            {
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr["school"].ToString() != "")
                    {
                        JsonData jd = new JsonData();
                        string strsql = string.Empty;
                        string school = dr["school"].ToString();
                        string xq = dr["xq"].ToString();
                        string building = dr["building"].ToString();
                        string EquType = dr["EquType"].ToString();
                        string equName = dr["equName"].ToString();
                        string devicesn = dr["devicesn"].ToString();
                        string status = dr["status"].ToString();
                        string remark = dr["remark"].ToString();
                        string usingDate = dr["usingDate"].ToString();
                        JsonData jdlist = new JsonData();
                        jdlist.SetJsonType(JsonType.Array);
                        //System.DateTime createDate = new System.DateTime();
                        string createDate = System.DateTime.Now.ToString();
                        string deviceno = GetSpellCode(xq) + "-" + GetSpellCode(building) + "-" + GetSpellCode(EquType);
                        using (DbManager db = new DbManager("name"))
                        {
                            string[] insertdata = { EquType, school, xq, equName, deviceno, devicesn, status, remark, usingDate };
                            if (school != "" || xq != "" || EquType != "" || equName != "" || deviceno != "" || devicesn != "" || status != "" || remark != "")
                            {

                                strsql = "insert into [equ_info]([type_id],[xxid],[xqid],[equname] ,[deviceno],[devicesn] ,[state] ,[remark] ,[usingdate],[createdate]) values('" + EquType + "','" + school + "','" + xq + "','" + equName + "','" + deviceno + "','" + devicesn + "','" + status + "','" + remark + "','" + usingDate + "','" + createDate + "')select @@identity as infoid";
                                dbm.SetCommand(strsql);
                                System.Data.SqlClient.SqlDataReader reader = (System.Data.SqlClient.SqlDataReader)dbm.ExecuteReader();
                                while (reader.Read())
                                { if (reader["infoid"].ToString() == "") { return false; } }
                            }
                        }
                    }
                    else { return true; }
                }
            }
            return true;
        }
        catch
        {
            return false;
        }

    }

    public static string GetSpellCode(string CnStr)
    {
        string strTemp = "";
        int iLen = CnStr.Length;
        int i = 0;
        for (i = 0; i <= iLen - 1; i++)
        {
            strTemp += GetCharSpellCode(CnStr.Substring(i, 1));
        }
        return strTemp;
    }

    private static string GetCharSpellCode(string CnChar)
    {
        long iCnChar;
        byte[] ZW = System.Text.Encoding.Default.GetBytes(CnChar);
        //如果是字母，则直接返回
        if (ZW.Length == 1)
        {
            return CnChar.ToUpper();
        }
        else
        {
            // get the array of byte from the single char
            int i1 = (short)(ZW[0]);
            int i2 = (short)(ZW[1]);
            iCnChar = i1 * 256 + i2;
        }
        // iCnChar match the constant
        if ((iCnChar >= 45217) && (iCnChar <= 45252))
        {
            return "A";
        }
        else if ((iCnChar >= 45253) && (iCnChar <= 45760))
        {
            return "B";
        }
        else if ((iCnChar >= 45761) && (iCnChar <= 46317))
        {
            return "C";
        }
        else if ((iCnChar >= 46318) && (iCnChar <= 46825))
        {
            return "D";
        }
        else if ((iCnChar >= 46826) && (iCnChar <= 47009))
        {
            return "E";
        }
        else if ((iCnChar >= 47010) && (iCnChar <= 47296))
        {
            return "F";
        }
        else if ((iCnChar >= 47297) && (iCnChar <= 47613))
        {
            return "G";
        }
        else if ((iCnChar >= 47614) && (iCnChar <= 48118))
        {
            return "H";
        }
        else if ((iCnChar >= 48119) && (iCnChar <= 49061))
        {
            return "J";
        }
        else if ((iCnChar >= 49062) && (iCnChar <= 49323))
        {
            return "K";
        }
        else if ((iCnChar >= 49324) && (iCnChar <= 49895))
        {
            return "L";
        }
        else if ((iCnChar >= 49896) && (iCnChar <= 50370))
        {
            return "M";
        }
        else if ((iCnChar >= 50371) && (iCnChar <= 50613))
        {
            return "N";
        }
        else if ((iCnChar >= 50614) && (iCnChar <= 50621))
        {
            return "O";
        }
        else if ((iCnChar >= 50622) && (iCnChar <= 50905))
        {
            return "P";
        }
        else if ((iCnChar >= 50906) && (iCnChar <= 51386))
        {
            return "Q";
        }
        else if ((iCnChar >= 51387) && (iCnChar <= 51445))
        {
            return "R";
        }
        else if ((iCnChar >= 51446) && (iCnChar <= 52217))
        {
            return "S";
        }
        else if ((iCnChar >= 52218) && (iCnChar <= 52697))
        {
            return "T";
        }
        else if ((iCnChar >= 52698) && (iCnChar <= 52979))
        {
            return "W";
        }
        else if ((iCnChar >= 52980) && (iCnChar <= 53640))
        {
            return "X";
        }
        else if ((iCnChar >= 53689) && (iCnChar <= 54480))
        {
            return "Y";
        }
        else if ((iCnChar >= 54481) && (iCnChar <= 55289))
        {
            return "Z";
        }
        else
            return ("?");
    }
}
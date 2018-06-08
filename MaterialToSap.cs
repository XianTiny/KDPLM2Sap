using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using SAP.Middleware.Connector;
using System.IO;

namespace DoSap1Rfc
{
    public class MaterialToSap
    {
        protected string __conn;
        protected string __flowId;

        public MaterialToSap()
        {
            __conn = "";
            __flowId = "";
        }

        public MaterialToSap(string conn, string flowId)
        {
            __conn = conn;
            __flowId = flowId;
        }

        public string doActions()
        {
            try
            {
                SqlConnection sqlconn = new SqlConnection(__conn);
                string sqlstr = "SELECT ObjectId FROM WF_Object WHERE BaseId='" + __flowId + "'";


                SqlDataAdapter sda = new SqlDataAdapter(sqlstr, sqlconn);
                //获取
                DataSet ds = new DataSet();
                sda.Fill(ds);
                DataTable dt3 = new DataTable();
                dt3 = ds.Tables[0];

                // 链接
                RfcConfigParameters rfcPar = new RfcConfigParameters();
                rfcPar.Add(RfcConfigParameters.Name, "DEV");
                rfcPar.Add(RfcConfigParameters.AppServerHost, "10.98.0.22");
                rfcPar.Add(RfcConfigParameters.Client, "400");
                rfcPar.Add(RfcConfigParameters.User, "USER01");
                rfcPar.Add(RfcConfigParameters.Password, "1234567890");
                rfcPar.Add(RfcConfigParameters.SystemNumber, "00");//SAP实例
                rfcPar.Add(RfcConfigParameters.Language, "ZH");
                RfcDestination dest = RfcDestinationManager.GetDestination(rfcPar);
                RfcRepository rfcrep = dest.Repository;
                IRfcFunction myfun = dest.Repository.CreateFunction("ZSET_MATERIEL_MARA");
                myfun.SetValue("P_CODE", "N");//SAP里面的传入参数        // N 新增;M 修改

                IRfcStructure import = null;
                IRfcTable table;
                bool wrong = true;
                string error = "";
                DataRow dr;

                DataTable dt2 = new DataTable();
                dt2.Columns.Add("DATA1", typeof(string));
                dt2.Columns.Add("DATA2", typeof(string));
                dt2.Columns.Add("DATA3", typeof(string));
                dt2.Columns.Add("DATA4", typeof(string));
                dt2.Columns.Add("DATA5", typeof(string));
                table = myfun.GetTable("T_ZMARA");
                string strsql2 = "";
                DataTable dt = new DataTable();
                SqlDataAdapter sda2;

                if (dt3.Rows.Count <= 0)
                    return "Error:流程行不足1";

                for (int k = 0; k < dt3.Rows.Count; k++)
                {
                    strsql2 = "exec  GETmat_woo  '" + dt3.Rows[k]["ObjectId"] + "'";
                    sda2 = new SqlDataAdapter(strsql2, sqlconn);
                    sqlconn.Open();
                    ds.Clear();
                    dt.Clear();
                    sda2.Fill(ds);
                    dt = ds.Tables[0];
                    sqlconn.Close();

                    if (dt.Rows.Count <= 0)
                        continue;
                    else
                        if(k!=0)
                            loopstr2 = loopstr2 + " union";

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        import = rfcrep.GetStructureMetadata("ZMARA").CreateStructure();

                        //赋值
                        import.SetValue("MATNR", dt.Rows[i]["物料编码"].ToString());    //物料编码  
                        import.SetValue("LTEXT", dt.Rows[i]["物料長描述"]);   //物料描述长文本  
                        import.SetValue("MAKTX", dt.Rows[i]["物料短描述"]);   //物料描述
                        import.SetValue("MATKL", dt.Rows[i]["物料组"]);// dt.Rows[0][2]);   //物料组  
                     

                        table.Append(import);
                     }
                }
                myfun.Invoke(dest);

                table = myfun.GetTable("T_ZMARA_RESULT");

                for (int j = 0; j < table.Count; j++)
                {
                    table.CurrentIndex = j;
                    dr = dt2.NewRow();

                    dr["DATA1"] = table[j].GetValue("MANDT").ToString();
                    dr["DATA2"] = table[j].GetValue("MATNR").ToString();
                    dr["DATA3"] = table[j].GetValue("MAKTX").ToString();
                    dr["DATA4"] = table[j].GetValue("ZRESULT").ToString();
                    dr["DATA5"] = table[j].GetValue("ZRESULT_TEXT").ToString();

                    if (Convert.ToString(dr["DATA4"]) == "N" && (!(table[j].GetValue("ZRESULT_TEXT").ToString().Contains("已在创建过程"))) && (table[j].GetValue("ZRESULT_TEXT").ToString() != ""))
                    {
                        wrong = false;
                        error = error + "第" + (j + 1) + "行物料[" + (j + 1) + "]" + table[j].GetValue("MATNR").ToString() + "[" + table[j].GetValue("MAKTX").ToString() + "]" + "<出错原因:" + table[j].GetValue("ZRESULT_TEXT").ToString() + ">;\r\n";
                    }

                    dt2.Rows.Add(dr);
                }

                // dt 包含返回结果
                if (wrong == false)
                    return error;
                else
                    return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

        }
    }

    // 
    public class BomToSap
    {
        protected string __conn;
        protected string __flowId;

        public BomToSap()
        {
            __conn = "";
            __flowId = "";

        }

        public BomToSap(string conn, string flowId)
        {
            __conn = conn;
            __flowId = flowId;
        }

        public string doAction()
        {
            try
            {
                SqlConnection sqlconn = new SqlConnection(__conn);
                string sqlstr = @"SELECT ObjectId FROM WF_Object WHERE BaseId='" + __flowId + "'";

                SqlDataAdapter sda = new SqlDataAdapter(sqlstr, sqlconn);
                //获取
                DataSet ds = new DataSet();
                sda.Fill(ds);
                DataTable dt = new DataTable();
                dt = ds.Tables[0];

                // 链接
                RfcConfigParameters rfcPar = new RfcConfigParameters();
                rfcPar.Add(RfcConfigParameters.Name, "DEV");
                rfcPar.Add(RfcConfigParameters.AppServerHost, "10.98.0.22");
                rfcPar.Add(RfcConfigParameters.Client, "400");
                rfcPar.Add(RfcConfigParameters.User, "USER01");
                rfcPar.Add(RfcConfigParameters.Password, "1234567890");
                rfcPar.Add(RfcConfigParameters.SystemNumber, "00");//SAP实例
                rfcPar.Add(RfcConfigParameters.Language, "ZH");
                RfcDestination dest = RfcDestinationManager.GetDestination(rfcPar);
                RfcRepository rfcrep = dest.Repository;

                IRfcFunction myfun = dest.Repository.CreateFunction("Z_RFC_PLM_SAP_BOM_CREATE");

                IRfcStructure import = null;
                IRfcTable table = myfun.GetTable("ZT_PP0011");
                IRfcTable retable;
                SqlDataAdapter sda2;
                DataTable dt2;
                DataTable dt3;
                bool wrong = true;
                string error = "";
                DataRow dr;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    sqlstr = @"exec GETBOM_woo  '" + dt.Rows[i][0] + "'";
                    sda2 = new SqlDataAdapter(sqlstr, sqlconn);

                    dt2 = new DataTable();
                    sda2.Fill(dt2);

                    for (int j = 0; j < dt2.Rows.Count; j++)
                    {
                        import = rfcrep.GetStructureMetadata("ZPP0011").CreateStructure();

                        import.SetValue("MATNR", dt2.Rows[j][0]);//物料编号
                        import.SetValue("DATUV", dt2.Rows[j][1]);//有效起始日期
                        import.SetValue("IDNRK", dt2.Rows[j][2]);//BOM 组件
                        import.SetValue("BMENG", dt2.Rows[j][3]);//基本数量
                        import.SetValue("STLST", dt2.Rows[j][4]);//BOM 状态 
                        import.SetValue("MENGE", dt2.Rows[j][5]);//组件数量
                        import.SetValue("MEINS", dt2.Rows[j][6]);//组件计量单位
                        import.SetValue("DATUV_I", dt2.Rows[j][7]);//有效起始日期
                        import.SetValue("DATUB_I", dt2.Rows[j][8]);//有效截止日期
                        table.Append(import);

                    }
                    myfun.Invoke(dest);
                }
                dt3 = new DataTable();
                dt3.Columns.Add("MATNR", typeof(string));
                dt3.Columns.Add("POSNR", typeof(string));
                dt3.Columns.Add("IDNRK", typeof(string));
                dt3.Columns.Add("FLAG", typeof(string));
                dt3.Columns.Add("ZRESULTS", typeof(string));
                retable = myfun.GetTable("ZT_PP0011_RLT");

                //re1 = re1 + " kmax=" + Convert.ToString(retable.RowCount) + " ; table= " + Convert.ToString(table.RowCount);
                for (int k = 0; k < retable.Count; k++)
                {
                    retable.CurrentIndex = k;
                    dr = dt3.NewRow();

                    dr["MATNR"] = retable[k].GetValue("MATNR").ToString();
                    dr["POSNR"] = retable[k].GetValue("POSNR").ToString();
                    dr["IDNRK"] = retable[k].GetValue("IDNRK").ToString();
                    dr["FLAG"] = retable[k].GetValue("FLAG").ToString();
                    dr["ZRESULTS"] = retable[k].GetValue("ZRESULTS").ToString();

                    if (Convert.ToString(dr["FLAG"]) == "E")
                    {
                        if (dr["ZRESULTS"].ToString() == "" || dr["ZRESULTS"].ToString().Contains("创建过程中"))
                            continue;

                        wrong = false;
                        error = error + "第" + (k + 1) + "行子物料" + retable[k].GetValue("MATNR").ToString() + "<出错原因:" + dr["ZRESULTS"].ToString() + ">;";
                    }
                    dt3.Rows.Add(dr);
                }
                if (wrong == false)
                    return error;
                else
                    return "";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
    }
}

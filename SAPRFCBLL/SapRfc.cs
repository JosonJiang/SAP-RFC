using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace SAPRFCBLL
{
    public class SapRfcBLL
    {

 
        public static DataTable CRMRFC()
        {

            string Conn = "172.20.1.36";
            string rfcName = "z_rfc_crm_0002";

#if DEBUG

            SapRFCHelper saph = new SapRFCHelper();

            //string[] param = { "ZTEST|Y" };
            //SAP.Middleware.Connector.IRfcStructure itab = saph.GetRfcStructure("", rfcName, "WERKS");

            List<string> e = new List<string>();
            e.Add(CreateParam("VKORG", "1000"));                //销售组织
            e.Add(CreateParam("FKDAT_LOW", "20151212"));        //起始更新日期，（必填）
            e.Add(CreateParam("FKDAT_HIGH", "20160130"));       //终止更新日期，（必填）
            //e.Add("VBELN_LOW|");                                //起始交货单号 
            //e.Add("VBELN_HIGH|");                               //终止交货单号
            //e.Add("KUNAG_LOW|");                             //起始售达方
            //e.Add("KUNAG_HIGH|");                            //终止售达方 

            //String[] arry = { "MATNR|4706-T18250-0320", "BUKRS|SDT" };

            DataTable dt = saph.GetRfcOutDt(Conn, e.ToArray(), rfcName, "ZSTR_CRM_DNLIST");

            string returncode = saph.GetRfcString("", e.ToArray(), rfcName, "RETCODE");
            string RETURNstr = saph.GetRfcString("", e.ToArray(), rfcName, "RETMSG");

            //RETURNstr = saph.GetRfcString("", arry, rfcName, "RETURN");

#else

            Dictionary<String, String> Dict = new Dictionary<string, string>();
            Dict.Add("VKORG", "1000");
            Dict.Add("FKDAT_LOW", "20151212");
            Dict.Add("FKDAT_HIGH", "20160130");

            dt = SapRFCHelper.GetRfcOutDt(Dict, rfcName, "IPPLFIX");
            returncode = SapRFCHelper.GetRFCString(Dict, rfcName, "RETCODE");
            RETURNstr = SapRFCHelper.GetRFCString(Dict, rfcName, "RETURN");
#endif

            return dt;

        }


        public static DataTable RFC()
        {

            string rfcName = "Z_RFC_PPL_REAL";

#if DEBUG

            SapRFCHelper saph = new SapRFCHelper();

            //string[] param = { "ZTEST|Y" };
            //SAP.Middleware.Connector.IRfcStructure itab = saph.GetRfcStructure("", rfcName, "WERKS");

            List<string> e = new List<string>();
            e.Add(CreateParam("MATNR", "4706-T18250-0320"));
            e.Add(CreateParam("BUKRS", "SDT"));
            //e.Add("WERKS|7000");
            //e.Add("MATNR|4706-T18250-0320");
            //e.Add("BUKRS|SDT");

            String[] arry = { "MATNR|4706-T18250-0320", "BUKRS|SDT" };

            DataTable dt = saph.GetRfcOutTable("", e.ToArray(), rfcName, "IPPLFIX");
            string returncode = saph.GetRfcString("", e.ToArray(), rfcName, "RETCODE");
            string RETURNstr = saph.GetRfcString("", e.ToArray(), rfcName, "RETURN");

            RETURNstr = saph.GetRfcString("", arry, rfcName, "RETURN");

#else

            Dictionary<String, String> Dict = new Dictionary<string, string>();
            Dict.Add("MATNR", "4706-T18250-0320");
            Dict.Add("BUKRS", "SDT");

            dt = SapRFCHelper.GetRfcOutDt(Dict, rfcName, "IPPLFIX");
            returncode = SapRFCHelper.GetRFCString(Dict, rfcName, "RETCODE");
            RETURNstr = SapRFCHelper.GetRFCString(Dict, rfcName, "RETURN");
#endif

            return dt;

        }

        public DataTable GetModuleList(string PartNumber)
        {
            SapRFCHelper saph = new SapRFCHelper();
            string rfcName = "ZJTESTCS15";
            List<string> parameters = new List<string>();
            parameters.Add(CreateParam("ZMATNR", PartNumber));
            parameters.Add("ZWERKS|7000");
            //parameters.Add("BUKRS|SDT");
            DataTable dt = saph.GetRfcOutTable("", parameters.ToArray(), rfcName, "ZSTPO");
            return dt;
        }




        public List<String> GetModelList(string pn)
        {
            List<string> str = new List<string>();
            DataTable dt = GetModuleList(pn);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i][0].ToString() == "0")
                {
                    if (!str.Contains(dt.Rows[i - 1][1].ToString()))
                    {
                        str.Add(dt.Rows[i - 1][1].ToString());
                    }
                }
            }
            return str;
        }



        public static string CreateParam(string paramName, string val)
        {
            return paramName.Trim() + "|" + val.Trim();
        }
    }
}

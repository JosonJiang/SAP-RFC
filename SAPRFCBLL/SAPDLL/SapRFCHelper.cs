using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using SAP.Middleware.Connector;
using System.Net;

namespace saphelper
{
    [Serializable] 
    public class SapRFCHelper
    {

        #region --基本方法 ------------
        /// <summary>
        /// 获取登录SAP参数
        /// </summary>
        /// <returns></returns>
        public RfcConfigParameters GetRfcLoginParameters(string conn)
        {
            //1.获取连接参数
            RfcConfigParameters parameters = new RfcConfigParameters();
            if (conn == "PRD") //正式版
            {
                parameters[RfcConfigParameters.Name] = "PRD";
                parameters[RfcConfigParameters.User] = "IT09";
                parameters[RfcConfigParameters.Password] = "123";
                parameters[RfcConfigParameters.Client] = "200";
                parameters[RfcConfigParameters.Language] = "ZH";
                parameters[RfcConfigParameters.AppServerHost] = "172.20.1.31";
                parameters[RfcConfigParameters.SystemNumber] = "00";
                parameters[RfcConfigParameters.IdleTimeout] = "60000";
                parameters[RfcConfigParameters.PoolSize] = "5";


            }
            else if (conn == "QAS")
            {
                parameters[RfcConfigParameters.Name] = "QAS";
                parameters[RfcConfigParameters.User] = "IT09";
                parameters[RfcConfigParameters.Password] = "abc.123";
                parameters[RfcConfigParameters.Client] = "200";
                parameters[RfcConfigParameters.Language] = "ZH";
                parameters[RfcConfigParameters.AppServerHost] = "172.20.1.36";
                parameters[RfcConfigParameters.SystemNumber] = "00";
                parameters[RfcConfigParameters.IdleTimeout] = "60000";
                parameters[RfcConfigParameters.PoolSize] = "5";
            }
            return parameters;    
        }

        /// <summary>
        /// 调用RFC，返回一个SAP Structure
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="rfcName"></param>
        /// <param name="IRfcStructureName"></param>
        /// <returns></returns>
        public IRfcStructure GetRfcStructure(string conn, string rfcName, string IRfcStructureName)
        {
            //1.登录SAP
            RfcConfigParameters parameters = GetRfcLoginParameters(conn);//获取登录参数
            RfcDestination rd = RfcDestinationManager.GetDestination(parameters);

            RfcRepository repo = rd.Repository;
            IRfcFunction f = repo.CreateFunction(rfcName);   //调用函数名

            IRfcStructure istr = f.GetStructure(IRfcStructureName);


            return istr;
        }

        /// <summary>
        /// 调用RFC,返回一个空的SAP 内部表结构
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="rfcName"></param>
        /// <param name="tableName"></param>
        /// <returns></returns> 
        public IRfcTable GetRfcTable(string conn, string rfcName, string IRfcTableName)
        {
            //1.登录SAP
            RfcConfigParameters parameters = GetRfcLoginParameters(conn);//获取登录参数
            RfcDestination rd = RfcDestinationManager.GetDestination(parameters);

            RfcRepository repo = rd.Repository;
            IRfcFunction f = repo.CreateFunction(rfcName);   //调用函数名

            IRfcTable itab = f.GetTable(IRfcTableName); //获取表结构

            //itab.Append();

            return itab;            
        }

        /// <summary>
        /// 给空内部赋值，返回一个带数据的IRfcTable
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="rfcName"></param>
        /// <param name="IRfcTableName"></param>
        /// <param name="valueTable"></param>
        /// <returns></returns>
        public IRfcTable GetRfcTableWithValue(string conn, string rfcName, string IRfcTableName, DataTable valueTable)
        {
            //获取空的内部结构
            IRfcTable itb = GetRfcTable(conn, rfcName, IRfcTableName);

            //根据传入valueTable，给内表赋值（字段名必须一致）
            if (valueTable.Rows.Count > 0)
            {
                for (int r = 0; r < valueTable.Rows.Count; r++)
                {
                    itb.Insert();
                    for (int c = 0; c < valueTable.Columns.Count; c++)
                    {
                        itb.CurrentRow.SetValue(valueTable.Columns[c].ColumnName.ToString(), valueTable.Rows[r][c].ToString());
                    }
                }
            }
            return itb;
        }

        /// <summary>
        /// 给空的结构赋值，返回一个带数据的IRfcStructure
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="rfcName"></param>
        /// <param name="IRfcStructureName"></param>
        /// <param name="valueTable"></param>
        /// <returns></returns>
        public IRfcStructure GetRfcStructureWithValue(string conn, string rfcName, string IRfcStructureName, DataTable valueTable)
        {
            //获取空的内部结构
            IRfcStructure istr = GetRfcStructure(conn, rfcName, IRfcStructureName);

            //根据传入valueTable，给内表赋值（字段名必须一致）
            for (int i = 0; i < valueTable.Columns.Count; i++)
            {
                istr.SetValue(valueTable.Columns[i].ColumnName, valueTable.Rows[0][i].ToString());
            }
            return istr;
        }

        /// <summary>
        /// 将SAP内表转换成DataTable
        /// </summary>
        /// <param name="rfcTable">内表名称</param>
        /// <returns>返回一个DataTable</returns>
        public DataTable ConvertToTable(IRfcTable rfcTable)
        {
            DataTable dt = new DataTable();

            //建立表结构
            for (int col = 0; col < rfcTable.ElementCount; col++)
            {
                RfcElementMetadata rfcCol = rfcTable.GetElementMetadata(col);
                string columnName = rfcCol.Name;
                dt.Columns.Add(columnName);
            }

            for (int rx = 0; rx < rfcTable.RowCount; rx++)
            {
                object[] dr = new object[rfcTable.ElementCount];

                for (int cx = 0; cx < dt.Columns.Count; cx++)
                {
                    dr[cx] = rfcTable[rx][dt.Columns[cx].ColumnName].GetValue();
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        /// <summary>
        /// 将SAP 结构转成DataTable
        /// </summary>
        /// <param name="rfcStru"></param>
        /// <returns></returns>
        public DataTable ConvertToTable(IRfcStructure rfcStru)
        {
            DataTable dt = new DataTable();

            //建立表结构
            for (int col = 0; col < rfcStru.ElementCount; col++)
            {
                RfcElementMetadata rfcCol = rfcStru.GetElementMetadata(col);
                string columnName = rfcCol.Name;
                dt.Columns.Add(columnName);
            }


            object[] dr = new object[rfcStru.ElementCount];
            for (int cx = 0; cx < dt.Columns.Count; cx++)
            {
                dr[cx] = rfcStru[dt.Columns[cx].ColumnName].GetValue();
            }
            dt.Rows.Add(dr);

            return dt;
        }


        #endregion ------------------

        #region --调用RFC 返回IRfcFunction ----------------

        /// <summary>
        /// 调用RFC，返回一个IRfcFuntion，当需要返回多个参数的时候调用
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="rfcName"></param>
        /// <param name="param"></param>
        /// <returns></returns>
        public IRfcFunction GetIRfcFun(string conn, string rfcName, string[] param)
        {
            //1.登录SAP
            RfcConfigParameters parameters = GetRfcLoginParameters(conn);//获取登录参数
            RfcDestination rd = RfcDestinationManager.GetDestination(parameters);

            RfcRepository repo = rd.Repository;
            IRfcFunction f = repo.CreateFunction(rfcName);   //调用函数名

            foreach (string value in param)
            {
                string[] keyvalue = value.Split('|');
                f.SetValue(keyvalue[0], keyvalue[1].Trim());//传递入参数
            }

            f.Invoke(rd); //执行函数

            rd = null;
            
            return f;
        }

        /// <summary>
        /// 调用RFC，返回一个IRfcFuntion，当需要返回多个参数的时候调用
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="rfcName"></param>
        /// <param name="param"></param>
        /// <param name="irfcStructureName"></param>
        /// <param name="irfcStr"></param>
        /// <returns></returns>
        public IRfcFunction GetIRfcFun(string conn, string rfcName, string[] param, string irfcStructureName, IRfcStructure irfcStr)
        {
            //1.登录SAP
            RfcConfigParameters parameters = GetRfcLoginParameters(conn);//获取登录参数
            RfcDestination rd = RfcDestinationManager.GetDestination(parameters);

            RfcRepository repo = rd.Repository;
            IRfcFunction f = repo.CreateFunction(rfcName);   //调用函数名

            //传入字符参数
            foreach (string value in param)
            {
                string[] keyvalue = value.Split('|');
                f.SetValue(keyvalue[0], keyvalue[1].Trim());//传递入参数
            }

            //--传入 结构参数--
            f.SetValue(irfcStructureName, irfcStr);

            f.Invoke(rd); //执行函数

            rd = null;
            return f;
        }

        /// <summary>
        /// 调用RFC，返回一个IRfcFuntion，当需要返回多个参数的时候调用
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="rfcName"></param>
        /// <param name="param"></param>
        /// <param name="irfcTableName"></param>
        /// <param name="irfcTab"></param>
        /// <returns></returns>
        public IRfcFunction GetIRfcFun(string conn, string rfcName, string[] param, string irfcTableName, IRfcTable irfcTab)
        {
            //1.登录SAP
            RfcConfigParameters parameters = GetRfcLoginParameters(conn);//获取登录参数
            RfcDestination rd = RfcDestinationManager.GetDestination(parameters);

            RfcRepository repo = rd.Repository;
            IRfcFunction f = repo.CreateFunction(rfcName);   //调用函数名
            //传入字符参数
            foreach (string value in param)
            {
                string[] keyvalue = value.Split('|');
                f.SetValue(keyvalue[0], keyvalue[1].Trim());//传递入参数
            }

            //--传入 内表 参数
            f.SetValue(irfcTableName, irfcTab);

            f.Invoke(rd); //执行函数

            rd = null;

            return f;
        }

        /// <summary>
        /// 调用RFC，返回一个IRfcFuntion，当需要返回多个参数的时候调用
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="rfcName"></param>
        /// <param name="param"></param>
        /// <param name="irfcStructureName"></param>
        /// <param name="irfcStr"></param>
        /// <param name="irfcTableName"></param>
        /// <param name="irfcTab"></param>
        /// <returns></returns>
        public IRfcFunction GetIRfcFun(string conn, string rfcName, string[] param, string irfcStructureName, IRfcStructure irfcStr, string irfcTableName, IRfcTable irfcTab)
        {
            //1.登录SAP
            RfcConfigParameters parameters = GetRfcLoginParameters(conn);//获取登录参数
            RfcDestination rd = RfcDestinationManager.GetDestination(parameters);

            RfcRepository repo = rd.Repository;
            IRfcFunction f = repo.CreateFunction(rfcName);   //调用函数名
            //传入字符参数 
            foreach (string value in param)
            {
                string[] keyvalue = value.Split('|');
                f.SetValue(keyvalue[0], keyvalue[1].Trim());//传递入参数
            }
            //--传入 结构参数--
            f.SetValue(irfcStructureName, irfcStr);
            //--传入 内部 参数
            f.SetValue(irfcTableName, irfcTab);

            f.Invoke(rd); //执行函数

            rd = null;
            return f;
        }

        #endregion

        #region--调用RFC 不返回任何数据-------------
        /// <summary>
        /// 直接调用RFC，不返回任何东西
        /// </summary>
        /// <param name="param">需要传入的字符串参数 例：{"P1|value1","P2|value2"}</param>
        /// <param name="rfcName">Rfc名称</param>
        public void ExecuteRFC(string conn, string[] param, string rfcName)
        {
            //1.登录SAP
            RfcConfigParameters parameters = GetRfcLoginParameters(conn);//获取登录参数
            RfcDestination rd = RfcDestinationManager.GetDestination(parameters);

            RfcRepository repo = rd.Repository;
            IRfcFunction f = repo.CreateFunction(rfcName);   //调用函数名
            //传入字符参数
            foreach (string value in param)
            {
                string[] keyvalue = value.Split('|');
                f.SetValue(keyvalue[0], keyvalue[1].Trim());//传递入参数
            }
            f.Invoke(rd); //执行函数

            rd = null;
        }

        /// <summary>
        /// 直接调用RFC，不返回任何东西
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="param">需要传入的字符串参数 例：{"P1|value1","P2|value2"}</param>
        /// <param name="rfcName"></param>
        /// <param name="irfcTableName"></param>
        /// <param name="irfcTab"></param>
        public void ExecuteRFC(string conn, string[] param, string rfcName, string irfcTableName, IRfcTable irfcTab)
        {
            //1.登录SAP
            RfcConfigParameters parameters = GetRfcLoginParameters(conn);//获取登录参数
            RfcDestination rd = RfcDestinationManager.GetDestination(parameters);

            RfcRepository repo = rd.Repository;
            IRfcFunction f = repo.CreateFunction(rfcName);   //调用函数名
            //传入字符参数
            foreach (string value in param)
            {
                string[] keyvalue = value.Split('|');
                f.SetValue(keyvalue[0], keyvalue[1].Trim());//传递入参数
            }
            //传入内表参数
            f.SetValue(irfcTableName, irfcTab);

            f.Invoke(rd); //执行函数

            rd = null;
        }

        /// <summary>
        /// 直接调用RFC，不返回任何东西
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="param">需要传入的字符串参数 例：{"P1|value1","P2|value2"}</param>
        /// <param name="rfcName"></param>
        /// <param name="irfcStructureName"></param>
        /// <param name="irfcStru"></param>
        public void ExecuteRFC(string conn, string[] param, string rfcName, string irfcStructureName, IRfcStructure irfcStru)
        {
            //1.登录SAP
            RfcConfigParameters parameters = GetRfcLoginParameters(conn);//获取登录参数
            RfcDestination rd = RfcDestinationManager.GetDestination(parameters);

            RfcRepository repo = rd.Repository;
            IRfcFunction f = repo.CreateFunction(rfcName);   //调用函数名
            //传入字符参数
            foreach (string value in param)
            {
                string[] keyvalue = value.Split('|');
                f.SetValue(keyvalue[0], keyvalue[1].Trim());//传递入参数
            }

            //传入 sap结构 参数
            f.SetValue(irfcStructureName, irfcStru);

            f.Invoke(rd); //执行函数

            rd = null;
        }

        public void ExecuteRFC(string conn, string[] param, string rfcName, string irfcStructureName, IRfcStructure irfcStru, string irfcTableName, IRfcTable irfcTab)
        {
            //1.登录SAP
            RfcConfigParameters parameters = GetRfcLoginParameters(conn);//获取登录参数
            RfcDestination rd = RfcDestinationManager.GetDestination(parameters);

            RfcRepository repo = rd.Repository;
            IRfcFunction f = repo.CreateFunction(rfcName);   //调用函数名
            //传入字符参数
            foreach (string value in param)
            {
                string[] keyvalue = value.Split('|');
                f.SetValue(keyvalue[0], keyvalue[1].Trim());//传递入参数
            }

            //传入 sap结构 参数
            f.SetValue(irfcStructureName, irfcStru);

            //传入内表参数
            f.SetValue(irfcTableName, irfcTab);

            f.Invoke(rd); //执行函数

            rd = null;
        }

        #endregion

        #region--调用RFC 返回一个字符参数-------------
        /// <summary>
        /// 调用RFC，返回一个字符串参数
        /// </summary>
        /// <param name="param">需要传入的字符串参数 例：{"P1|value1","P2|value2"}</param>
        /// <param name="rfcName">Rfc名称</param>
        public string GetRfcString(string conn, string[] param, string rfcName, string outStringName)
        {
            //1.登录SAP
            RfcConfigParameters parameters = GetRfcLoginParameters(conn);//获取登录参数
            RfcDestination rd = RfcDestinationManager.GetDestination(parameters);

            RfcRepository repo = rd.Repository;
            IRfcFunction f = repo.CreateFunction(rfcName);   //调用函数名
            //传入字符参数
            foreach (string value in param)
            {
                string[] keyvalue = value.Split('|');
                f.SetValue(keyvalue[0], keyvalue[1].Trim());//传递入参数
            }
            f.Invoke(rd); //执行函数

            rd = null;
            return f.GetString(outStringName).ToString(); //返回字符串  
        }

        /// <summary>
        /// 调用RFC，返回一个字符串参数
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="param">需要传入的字符串参数 例：{"P1|value1","P2|value2"}</param>
        /// <param name="rfcName"></param>
        /// <param name="irfcTableName"></param>
        /// <param name="irfcTab"></param>
        public string ExecuteRFC(string conn, string[] param, string rfcName, string irfcTableName, IRfcTable irfcTab, string outStringName)
        {
            //1.登录SAP
            RfcConfigParameters parameters = GetRfcLoginParameters(conn);//获取登录参数
            RfcDestination rd = RfcDestinationManager.GetDestination(parameters);

            RfcRepository repo = rd.Repository;
            IRfcFunction f = repo.CreateFunction(rfcName);   //调用函数名
            //传入字符参数
            foreach (string value in param)
            {
                string[] keyvalue = value.Split('|');
                f.SetValue(keyvalue[0], keyvalue[1].Trim());//传递入参数
            }
            //传入内表参数
            f.SetValue(irfcTableName, irfcTab);

            f.Invoke(rd); //执行函数
            rd = null;
            return f.GetString(outStringName).ToString(); //返回字符串  
        }

        /// <summary>
        /// 调用RFC，返回一个字符串参数
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="param">需要传入的字符串参数 例：{"P1|value1","P2|value2"}</param>
        /// <param name="rfcName"></param>
        /// <param name="irfcStructureName"></param>
        /// <param name="irfcStru"></param>
        public string ExecuteRFC(string conn, string[] param, string rfcName, string irfcStructureName, IRfcStructure irfcStru, string outStringName)
        {
            //1.登录SAP
            RfcConfigParameters parameters = GetRfcLoginParameters(conn);//获取登录参数
            RfcDestination rd = RfcDestinationManager.GetDestination(parameters);

            RfcRepository repo = rd.Repository;
            IRfcFunction f = repo.CreateFunction(rfcName);   //调用函数名
            //传入字符参数
            foreach (string value in param)
            {
                string[] keyvalue = value.Split('|');
                f.SetValue(keyvalue[0], keyvalue[1].Trim());//传递入参数
            }

            //传入 sap结构 参数
            f.SetValue(irfcStructureName, irfcStru);

            f.Invoke(rd); //执行函数

            rd = null;
            return f.GetString(outStringName).ToString(); //返回字符串  
        }
        /// <summary>
        /// 调用RFC，返回一个字符串参数
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="param"></param>
        /// <param name="rfcName"></param>
        /// <param name="irfcStructureName"></param>
        /// <param name="irfcStru"></param>
        /// <param name="irfcTableName"></param>
        /// <param name="irfcTab"></param>
        /// <param name="outStringName"></param>
        /// <returns></returns>
        public string ExecuteRFC(string conn, string[] param, string rfcName, string irfcStructureName, IRfcStructure irfcStru, string irfcTableName, IRfcTable irfcTab, string outStringName)
        {
            //1.登录SAP
            RfcConfigParameters parameters = GetRfcLoginParameters(conn);//获取登录参数
            RfcDestination rd = RfcDestinationManager.GetDestination(parameters);

            RfcRepository repo = rd.Repository;
            IRfcFunction f = repo.CreateFunction(rfcName);   //调用函数名
            //传入字符参数
            foreach (string value in param)
            {
                string[] keyvalue = value.Split('|');
                f.SetValue(keyvalue[0], keyvalue[1].Trim());//传递入参数
            }

            //传入 sap结构 参数
            f.SetValue(irfcStructureName, irfcStru);

            //传入内表参数
            f.SetValue(irfcTableName, irfcTab);

            f.Invoke(rd); //执行函数

            rd = null;
            return f.GetString(outStringName).ToString(); //返回字符串  
        }

        #endregion

        #region--调用RFC 返回一个内表参数-------------
        /// <summary>
        /// 调用RFC，返回一个内表参数
        /// </summary>
        /// <param name="param">需要传入的字符串参数 例：{"P1|value1","P2|value2"}</param>
        /// <param name="rfcName">Rfc名称</param>
        public DataTable GetRfcOutTable(string conn, string[] param, string rfcName, string outTableName)
        {
            //1.登录SAP
            RfcConfigParameters parameters = GetRfcLoginParameters(conn);//获取登录参数
            RfcDestination rd = RfcDestinationManager.GetDestination(parameters);

            RfcRepository repo = rd.Repository;
            IRfcFunction f = repo.CreateFunction(rfcName);   //调用函数名
            //传入字符参数
            foreach (string value in param)
            {
                string[] keyvalue = value.Split('|');
                f.SetValue(keyvalue[0], keyvalue[1].Trim());//传递入参数
            }
            f.Invoke(rd); //执行函数
            rd = null;
            return ConvertToTable(f.GetTable(outTableName));
        }

        /// <summary>
        /// 调用RFC，返回一个内表参数
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="param">需要传入的字符串参数 例：{"P1|value1","P2|value2"}</param>
        /// <param name="rfcName"></param>
        /// <param name="irfcTableName"></param>
        /// <param name="irfcTab"></param>
        public DataTable GetRfcOutTable(string conn, string[] param, string rfcName, string irfcTableName, IRfcTable irfcTab, string outTableName)
        {
            //1.登录SAP
            RfcConfigParameters parameters = GetRfcLoginParameters(conn);//获取登录参数
            RfcDestination rd = RfcDestinationManager.GetDestination(parameters);

            RfcRepository repo = rd.Repository;
            IRfcFunction f = repo.CreateFunction(rfcName);   //调用函数名
            //传入字符参数
            foreach (string value in param)
            {
                string[] keyvalue = value.Split('|');
                f.SetValue(keyvalue[0], keyvalue[1].Trim());//传递入参数
            }
            //传入内表参数
            f.SetValue(irfcTableName, irfcTab);

            f.Invoke(rd); //执行函数
            rd = null;
            return ConvertToTable(f.GetTable(outTableName));
        }

        /// <summary>
        /// 调用RFC，返回一个内表参数
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="param">需要传入的字符串参数 例：{"P1|value1","P2|value2"}</param>
        /// <param name="rfcName"></param>
        /// <param name="irfcStructureName"></param>
        /// <param name="irfcStru"></param>
        public DataTable GetRfcOutTable(string conn, string[] param, string rfcName, string irfcStructureName, IRfcStructure irfcStru, string outTableName)
        {
            //1.登录SAP
            RfcConfigParameters parameters = GetRfcLoginParameters(conn);//获取登录参数
            RfcDestination rd = RfcDestinationManager.GetDestination(parameters);

            RfcRepository repo = rd.Repository;
            IRfcFunction f = repo.CreateFunction(rfcName);   //调用函数名
            //传入字符参数
            foreach (string value in param)
            {
                string[] keyvalue = value.Split('|');
                f.SetValue(keyvalue[0], keyvalue[1].Trim());//传递入参数
            }

            //传入 sap结构 参数
            f.SetValue(irfcStructureName, irfcStru);

            f.Invoke(rd); //执行函数

            rd = null;
            return ConvertToTable(f.GetTable(outTableName));
        }
        /// <summary>
        /// 调用RFC，返回一个内表参数
        /// </summary>
        /// <param name="conn"></param>
        /// <param name="param"></param>
        /// <param name="rfcName"></param>
        /// <param name="irfcStructureName"></param>
        /// <param name="irfcStru"></param>
        /// <param name="irfcTableName"></param>
        /// <param name="irfcTab"></param>
        /// <param name="outStringName"></param>
        /// <returns></returns>
        public DataTable GetRfcOutTable(string conn, string[] param, string rfcName, string irfcStructureName, IRfcStructure irfcStru, string irfcTableName, IRfcTable irfcTab, string outTableName)
        {
            //1.登录SAP
            RfcConfigParameters parameters = GetRfcLoginParameters(conn);//获取登录参数
            RfcDestination rd = RfcDestinationManager.GetDestination(parameters);

            RfcRepository repo = rd.Repository;
            IRfcFunction f = repo.CreateFunction(rfcName);   //调用函数名
            //传入字符参数
            foreach (string value in param)
            {
                string[] keyvalue = value.Split('|');
                f.SetValue(keyvalue[0], keyvalue[1].Trim());//传递入参数
            }

            //传入 sap结构 参数
            f.SetValue(irfcStructureName, irfcStru);

            //传入内表参数
            f.SetValue(irfcTableName, irfcTab);

            f.Invoke(rd); //执行函数

            rd = null;
            return ConvertToTable(f.GetTable(outTableName));
        }

        #endregion
    }
}
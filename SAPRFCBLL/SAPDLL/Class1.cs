using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SapRFCHelper
{
    class Class1
    {


        string R_Messege;	//返回错误信息
string subrc;    	//返回结果，subrc=0时，表示正确；subrc=1时，错误；
string salesdocumen;	//SAP合同编号

//如果合同为历史记录为完成合同，不需要入SAP
if(Convert.ToString(FormDataSet["GNHT_Applicant_All.History_Status"])=="2")
{
	return;
}

//获取合同类型,	正常合同=ZCQ1，订货单=ZCQ2
string ContractType="";
switch(Convert.ToString(FormDataSet["GNHT_Applicant_All.ContractType"]))
{
case "正常合同":
	ContractType="ZCQ1";
	break;
case "订货单":
	ContractType="ZCQ2";
	break;
default:
	ContractType="";
	break;
}

//如果ContractType为空，表示合同类型不是"正常合同"或"订货单"，不需要与SAP交互
if(string.IsNullOrEmpty(ContractType))
{
	return;
}

//如果SAP合同号为空，表示不需要与SAP交互
if(string.IsNullOrEmpty(Convert.ToString(FormDataSet["GNHT_Applicant_All.SAP_ContractNO"])))
{
	return;
}

Z_RFC_0028.ZSD_CONTRACTHEAD_STR HeadItem = new Z_RFC_0028.ZSD_CONTRACTHEAD_STR();
HeadItem.Ktext = Convert.ToString(FormDataSet["GNHT_Applicant_All.AppSN"]);	      //合同流水号
HeadItem.Vbeln = Convert.ToString(FormDataSet["GNHT_Applicant_All.SAP_ContractNO"]);  //SAP合同号
HeadItem.Auart = ContractType;            					      //来合同类型	正常合同ZCQ1，订货单ZCQ2
HeadItem.Kunnr = Convert.ToString(FormDataSet["GNHT_Applicant_All.CustomerID"]);      //客户编号	售达方
HeadItem.Vkorg = Convert.ToString(FormDataSet["GNHT_Applicant_All.SalesArea_VKORG"]);            //销售组织	1000
HeadItem.Vtweg = Convert.ToString(FormDataSet["GNHT_Applicant_All.SalesArea_VTWEG"]);              //分销渠道
HeadItem.Spart = Convert.ToString(FormDataSet["GNHT_Applicant_All.SalesArea_SPART"]);              //产品组
HeadItem.Zterm = Convert.ToString(FormDataSet["GNHT_Applicant_All.PaymentNO"]);            //付款方式
HeadItem.Bstkd = Convert.ToString(FormDataSet["GNHT_Applicant_All.ContractNO"]);    //合同编号

if(!string.IsNullOrEmpty(Convert.ToString(FormDataSet["GNHT_Applicant_All.ContractStartDate"])))
{
	HeadItem.Guebg = Convert.ToDateTime(FormDataSet["GNHT_Applicant_All.ContractStartDate"]).ToString("yyyyMMdd");                //合同有效期从
}

if(!string.IsNullOrEmpty(Convert.ToString(FormDataSet["GNHT_Applicant_All.ContractEndDate"])))
{
	HeadItem.Gueen = Convert.ToDateTime(FormDataSet["GNHT_Applicant_All.ContractEndDate"]).ToString("yyyyMMdd");                //有效期到
}

if(!string.IsNullOrEmpty(Convert.ToString(FormDataSet["GNHT_Applicant_All.Y_AcceptanceTime"])))
{
	HeadItem.Mahza = Convert.ToInt32(FormDataSet["GNHT_Applicant_All.Y_AcceptanceTime"]);                 //验收时间,单位天
	HeadItem.Text1 = Convert.ToString(FormDataSet["GNHT_Applicant_All.Y_AcceptanceNote"]);                //验收条款文本
}

HeadItem.Text2 = Convert.ToString(FormDataSet["GNHT_Applicant_All.H_Note"]);                //合同备注

if(!string.IsNullOrEmpty(Convert.ToString(FormDataSet["GNHT_Applicant_All.Warranty"])))
{
	HeadItem.Vsnmr_V = Convert.ToString(FormDataSet["GNHT_Applicant_All.Warranty"]) + "个月";         //质保期
}

Z_RFC_0028.ZSD_CONTRACTITEMS_STRTable strTable = new Z_RFC_0028.ZSD_CONTRACTITEMS_STRTable();

//合同行项目信息
FlowDataTable table_all_sup= FormDataSet.Tables["GNHT_Applicant_Order"];

foreach (FlowDataRow row_all_sup in table_all_sup.Rows)
{
   //如果选择进入SAP，才执行
   //if(Convert.ToInt16(row_all_sup["isASP"])==1)
   //{
	Z_RFC_0028.ZSD_CONTRACTITEMS_STR Item = new Z_RFC_0028.ZSD_CONTRACTITEMS_STR();
	
	if(!string.IsNullOrEmpty(row_all_sup["ContractNum"].ToString().Trim()))
	{
		Item.Posnr = row_all_sup["ContractNum"].ToString().Trim();      //合同项目号
	}

	Item.Pstyv = row_all_sup["ContractNumType"].ToString().Trim();    //合同项目类别

	Item.Kdmat = row_all_sup["MachineType"].ToString().Trim();  //型号

	if(!string.IsNullOrEmpty(Convert.ToString(row_all_sup["Count"]).Trim()))
	{
		Item.Zmeng = Convert.ToDecimal(row_all_sup["Count"]);     //合同数量
	}
	
	if(!string.IsNullOrEmpty(Convert.ToString(row_all_sup["UnitPrice"]).Trim()))
	{
		Item.Kbetr = Convert.ToDecimal(row_all_sup["UnitPrice"]);      //单价
	}
	
	Item.Uflag = row_all_sup["isASP"].ToString().Trim();

	strTable.Add(Item);
    //}
}

SAP_GNHT_RFC_0028.Service1 myService = new SAP_GNHT_RFC_0028.Service1();

myService.RFC_0028_Control("IT09", "123", "172.20.1.31", "Z", HeadItem, out R_Messege, out subrc, out salesdocumen, ref strTable);

//throw new BPMException(subrc + "/" +  salesdocumen);

if (subrc!="0")
{
	throw new BPMException("SAP提示：" + R_Messege);
}
    }
}

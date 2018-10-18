//***********************************************************
//***********************************************************
//以下代码由工具自动生成,欢迎联系下面的作者邮箱，提出宝贵意见。
//285474816@qq.com
//***********************************************************
//***********************************************************
using RCMath;
using System;
namespace _3VJ_MV
{
	public partial class PartType Entity
	{
		#region 根据CSV模版生成的成员变量
		/// <summary>
		/// 序号
		/// </summary>
		string no._A_0 {get;set;}
		public string No. { get { return no._A_0 ; } set { no._A_0  = value; } }

		/// <summary>
		/// 基础板件类型
		/// </summary>
		string parttype_B_1 {get;set;}
		public string PartType { get { return parttype_B_1 ; } set { parttype_B_1  = value; } }

		/// <summary>
		/// 模型编号
		/// </summary>
		string partnumber_C_2 {get;set;}
		public string PartNumber { get { return partnumber_C_2 ; } set { partnumber_C_2  = value; } }

		/// <summary>
		/// 排程类型
		/// </summary>
		string productmodel_D_3 {get;set;}
		public string ProductModel { get { return productmodel_D_3 ; } set { productmodel_D_3  = value; } }

		public int colCount;
		#endregion
		public PartType Entity() { colCount = 4; }
		public PartType Entity(string csvString)
		{
			colCount = 4;
			string[] csvstrlist = new string[colCount];
			string[] csvstrlist0 = csvString.Split(',');
			if (csvstrlist0.Length <= 4) csvstrlist0.CopyTo(csvstrlist, 0);
			else csvstrlist = csvstrlist0;
			no._A_0  = csvstrlist[0];
			parttype_B_1  = csvstrlist[1];
			partnumber_C_2  = csvstrlist[2];
			productmodel_D_3  = csvstrlist[3];
		}
		public string OutPutCsvString()
		{
			string retString = "";
			retString += (no._A_0 +",");
			retString += (parttype_B_1 +",");
			retString += (partnumber_C_2 +",");
			retString += (productmodel_D_3 +",");
			retString = retString.Remove(retString.Length - 1);
			return retString;
		}
	}
}
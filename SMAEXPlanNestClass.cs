//***********************************************************
//***********************************************************
//以下代码由工具自动生成,欢迎联系下面的作者邮箱，提出宝贵意见。
//285474816@qq.com
//***********************************************************
//***********************************************************
using System;
namespace WorkList
{
	public partial class NestlistEntity
	{
		#region 根据CSV模版生成的成员变量
		/// <summary>
		/// 序号
		/// </summary>
		string index_A_0 {get;set;}
		public string Index { get { return index_A_0 ; } set { index_A_0  = value; } }

		/// <summary>
		/// 材料
		/// </summary>
		string mat_B_1 {get;set;}
		public string Mat { get { return mat_B_1 ; } set { mat_B_1  = value; } }

		/// <summary>
		/// 封边宽1
		/// </summary>
		string ew1_C_2 {get;set;}
		public string EW1 { get { return ew1_C_2 ; } set { ew1_C_2  = value; } }

		/// <summary>
		/// 封边宽2
		/// </summary>
		string ew2_D_3 {get;set;}
		public string EW2 { get { return ew2_D_3 ; } set { ew2_D_3  = value; } }

		/// <summary>
		/// 封边长1
		/// </summary>
		string el1_E_4 {get;set;}
		public string EL1 { get { return el1_E_4 ; } set { el1_E_4  = value; } }

		/// <summary>
		/// 封边长2
		/// </summary>
		string el2_F_5 {get;set;}
		public string EL2 { get { return el2_F_5 ; } set { el2_F_5  = value; } }

		/// <summary>
		/// 名称
		/// </summary>
		string partname_G_6 {get;set;}
		public string PartName { get { return partname_G_6 ; } set { partname_G_6  = value; } }

		/// <summary>
		/// 高
		/// </summary>
		string length_H_7 {get;set;}
		public string Length { get { return length_H_7 ; } set { length_H_7  = value; } }

		/// <summary>
		/// 宽
		/// </summary>
		string width_I_8 {get;set;}
		public string Width { get { return width_I_8 ; } set { width_I_8  = value; } }

		/// <summary>
		/// 数量
		/// </summary>
		string num_J_9 {get;set;}
		public string Num { get { return num_J_9 ; } set { num_J_9  = value; } }

		/// <summary>
		/// 正面加工码
		/// </summary>
		string filename_K_10 {get;set;}
		public string Filename { get { return filename_K_10 ; } set { filename_K_10  = value; } }

		/// <summary>
		/// 反面加工码
		/// </summary>
		string filename6_L_11 {get;set;}
		public string Filename6 { get { return filename6_L_11 ; } set { filename6_L_11  = value; } }

		/// <summary>
		/// 备注1（批次号）
		/// </summary>
		string batch_M_12 {get;set;}
		public string Batch { get { return batch_M_12 ; } set { batch_M_12  = value; } }

		/// <summary>
		/// 备注2（任务编码）
		/// </summary>
		string tasknum_N_13 {get;set;}
		public string Tasknum { get { return tasknum_N_13 ; } set { tasknum_N_13  = value; } }

		/// <summary>
		/// 备注3（序列号）
		/// </summary>
		string id_O_14 {get;set;}
		public string ID { get { return id_O_14 ; } set { id_O_14  = value; } }

		/// <summary>
		/// 备注4（分流）
		/// </summary>
		string split_P_15 {get;set;}
		public string Split { get { return split_P_15 ; } set { split_P_15  = value; } }

		/// <summary>
		/// 备注5
		/// </summary>
		string common5_Q_16 {get;set;}
		public string Common5 { get { return common5_Q_16 ; } set { common5_Q_16  = value; } }

		/// <summary>
		/// 备注6
		/// </summary>
		string common6_R_17 {get;set;}
		public string Common6 { get { return common6_R_17 ; } set { common6_R_17  = value; } }

		/// <summary>
		/// 备注7
		/// </summary>
		string common7_S_18 {get;set;}
		public string Common7 { get { return common7_S_18 ; } set { common7_S_18  = value; } }

		/// <summary>
		/// 备注8
		/// </summary>
		string common8_T_19 {get;set;}
		public string Common8 { get { return common8_T_19 ; } set { common8_T_19  = value; } }

		public int colCount;
		#endregion
		public NestlistEntity() { colCount = 20; }
		public NestlistEntity(string csvString)
		{
			colCount = 20;
			string[] csvstrlist = new string[colCount];
			string[] csvstrlist0 = csvString.Split(',');
			if (csvstrlist0.Length <= 20) csvstrlist0.CopyTo(csvstrlist, 0);
			else csvstrlist = csvstrlist0;
            index_A_0 = csvstrlist[0];
            mat_B_1 = csvstrlist[1];
            ew1_C_2 = csvstrlist[2];
            ew2_D_3 = csvstrlist[3];
            el1_E_4 = csvstrlist[4];
            el2_F_5 = csvstrlist[5];
            partname_G_6 = csvstrlist[6];
            length_H_7 = csvstrlist[7];
            width_I_8 = csvstrlist[8];
            num_J_9 = csvstrlist[9];
			filename_K_10  = csvstrlist[10];
			filename6_L_11  = csvstrlist[11];
            batch_M_12 = csvstrlist[12];
            tasknum_N_13 = csvstrlist[13];
            id_O_14 = csvstrlist[14];
            split_P_15 = csvstrlist[15];
            common5_Q_16 = csvstrlist[16];
            common6_R_17 = csvstrlist[17];
            common7_S_18 = csvstrlist[18];
			common8_T_19  = csvstrlist[19];
		}
		public string OutPutCsvString()
		{
			string retString = "";
			retString += (index_A_0 +",");
			retString += (mat_B_1 +",");
			retString += (ew1_C_2 +",");
			retString += (ew2_D_3 +",");
			retString += (el1_E_4 +",");
			retString += (el2_F_5 +",");
			retString += (partname_G_6 +",");
			retString += (length_H_7 +",");
			retString += (width_I_8 +",");
			retString += (num_J_9 +",");
			retString += (filename_K_10 +",");
			retString += (filename6_L_11 +",");
			retString += (batch_M_12 +",");
			retString += (tasknum_N_13 +",");
			retString += (id_O_14 +",");
			retString += (split_P_15 +",");
			retString += (common5_Q_16 +",");
			retString += (common6_R_17 +",");
			retString += (common7_S_18 +",");
			retString += (common8_T_19 +",");
			retString = retString.Remove(retString.Length - 1);
			return retString;
		}
	}
}
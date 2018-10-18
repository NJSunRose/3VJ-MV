namespace SWJXMLToCSV
{
	public partial class ClassEntity
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
		string material_B_1 {get;set;}
		public string Material { get { return material_B_1 ; } set { material_B_1  = value; } }

		/// <summary>
		/// 封边宽1
		/// </summary>
		string ebw1_C_2 {get;set;}
		public string EbW1 { get { return ebw1_C_2 ; } set { ebw1_C_2  = value; } }

		/// <summary>
		/// 封边宽2
		/// </summary>
		string ebw2_D_3 {get;set;}
		public string EbW2 { get { return ebw2_D_3 ; } set { ebw2_D_3  = value; } }

		/// <summary>
		/// 封边长1
		/// </summary>
		string ebl1_E_4 {get;set;}
		public string EbL1 { get { return ebl1_E_4 ; } set { ebl1_E_4  = value; } }

		/// <summary>
		/// 封边长2
		/// </summary>
		string ebl2_F_5 {get;set;}
		public string EbL2 { get { return ebl2_F_5 ; } set { ebl2_F_5  = value; } }

		/// <summary>
		/// 名称
		/// </summary>
		string partname_G_6 {get;set;}
		public string PartName { get { return partname_G_6 ; } set { partname_G_6  = value; } }

		/// <summary>
		/// 长度
		/// </summary>
		string length_H_7 {get;set;}
		public string Length { get { return length_H_7 ; } set { length_H_7  = value; } }

		/// <summary>
		/// 宽度
		/// </summary>
		string width_I_8 {get;set;}
		public string Width { get { return width_I_8 ; } set { width_I_8  = value; } }

		/// <summary>
		/// 数量
		/// </summary>
		string num_J_9 {get;set;}
		public string Num { get { return num_J_9 ; } set { num_J_9  = value; } }

		/// <summary>
		/// 加工代码
		/// </summary>
		string f5filename_K_10 {get;set;}
		public string F5FileName { get { return f5filename_K_10 ; } set { f5filename_K_10  = value; } }

		/// <summary>
		/// 反面加工代码
		/// </summary>
		string f6filename_L_11 {get;set;}
		public string F6FileName { get { return f6filename_L_11 ; } set { f6filename_L_11  = value; } }

		/// <summary>
		/// 备注1批次号
		/// </summary>
		string batchnum_M_12 {get;set;}
		public string BatchNum { get { return batchnum_M_12 ; } set { batchnum_M_12  = value; } }

		/// <summary>
		/// 备注2分拣号
		/// </summary>
		string boxnumber_N_13 {get;set;}
		public string BoxNumber { get { return boxnumber_N_13 ; } set { boxnumber_N_13  = value; } }

		/// <summary>
		/// 备注3板件号
		/// </summary>
		string partnumber_O_14 {get;set;}
		public string PartNumber { get { return partnumber_O_14 ; } set { partnumber_O_14  = value; } }

		/// <summary>
		/// 备注4分流
		/// </summary>
		string modelname_P_15 {get;set;}
		public string ModelName { get { return modelname_P_15 ; } set { modelname_P_15  = value; } }

		/// <summary>
		/// 备注5优化号
		/// </summary>
		string nestingnumber_Q_16 {get;set;}
		public string NestingNumber { get { return nestingnumber_Q_16 ; } set { nestingnumber_Q_16  = value; } }

		/// <summary>
		/// 备注6FTP目录
		/// </summary>
		string f5ftpadress_R_17 {get;set;}
		public string F5FTPAdress { get { return f5ftpadress_R_17 ; } set { f5ftpadress_R_17  = value; } }

		/// <summary>
		/// 备注7反面FTP目录
		/// </summary>
		string f6ftpadress_S_18 {get;set;}
		public string F6FTPAdress { get { return f6ftpadress_S_18 ; } set { f6ftpadress_S_18  = value; } }

		/// <summary>
		/// 备注8
		/// </summary>
		string nest_num_T_19 {get;set;}
		public string Nest_Num { get { return nest_num_T_19 ; } set { nest_num_T_19  = value; } }

		/// <summary>
		/// 订单号
		/// </summary>
		string order_U_20 {get;set;}
		public string Order { get { return order_U_20 ; } set { order_U_20  = value; } }

		/// <summary>
		/// 行号
		/// </summary>
		string linenumber_V_21 {get;set;}
		public string LineNumber { get { return linenumber_V_21 ; } set { linenumber_V_21  = value; } }

		public int colCount;
		#endregion
		public ClassEntity() { colCount = 22; }
		public ClassEntity(string csvString)
		{
			colCount = 22;
			string[] csvstrlist = new string[colCount];
			string[] csvstrlist0 = csvString.Split(',');
			if (csvstrlist0.Length <= 22) csvstrlist0.CopyTo(csvstrlist, 0);
			else csvstrlist = csvstrlist0;
			index_A_0  = csvstrlist[0];
			material_B_1  = csvstrlist[1];
			ebw1_C_2  = csvstrlist[2];
			ebw2_D_3  = csvstrlist[3];
			ebl1_E_4  = csvstrlist[4];
			ebl2_F_5  = csvstrlist[5];
			partname_G_6  = csvstrlist[6];
			length_H_7  = csvstrlist[7];
			width_I_8  = csvstrlist[8];
			num_J_9  = csvstrlist[9];
			f5filename_K_10  = csvstrlist[10];
			f6filename_L_11  = csvstrlist[11];
			batchnum_M_12  = csvstrlist[12];
			boxnumber_N_13  = csvstrlist[13];
			partnumber_O_14  = csvstrlist[14];
			modelname_P_15  = csvstrlist[15];
			nestingnumber_Q_16  = csvstrlist[16];
			f5ftpadress_R_17  = csvstrlist[17];
			f6ftpadress_S_18  = csvstrlist[18];
			nest_num_T_19  = csvstrlist[19];
			order_U_20  = csvstrlist[20];
			linenumber_V_21  = csvstrlist[21];
		}
		public string OutPutCsvString()
		{
			string retString = "";
			retString += (index_A_0 +",");
			retString += (material_B_1 +",");
			retString += (ebw1_C_2 +",");
			retString += (ebw2_D_3 +",");
			retString += (ebl1_E_4 +",");
			retString += (ebl2_F_5 +",");
			retString += (partname_G_6 +",");
			retString += (length_H_7 +",");
			retString += (width_I_8 +",");
			retString += (num_J_9 +",");
			retString += (f5filename_K_10 +",");
			retString += (f6filename_L_11 +",");
			retString += (batchnum_M_12 +",");
			retString += (boxnumber_N_13 +",");
			retString += (partnumber_O_14 +",");
			retString += (modelname_P_15 +",");
			retString += (nestingnumber_Q_16 +",");
			retString += (f5ftpadress_R_17 +",");
			retString += (f6ftpadress_S_18 +",");
			retString += (nest_num_T_19 +",");
            retString += (order_U_20 + ",");
            retString += (linenumber_V_21 + ",");
			retString = retString.Remove(retString.Length - 1);
			return retString;
		}
	}
}
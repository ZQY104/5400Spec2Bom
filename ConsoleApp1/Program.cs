using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;

namespace Sizing
{
    public class Program
    {
        private static void Main(string[] args)
        {
            //Console.WriteLine("Please enter the full file path of your Word document (without quotes):");
            //object path = Console.ReadLine();
            object path = @"c:/spec.doc"; //如果想自己从控制台输入 目标文件 路径,请解除上两行注释
            //Console.WriteLine("Please enter the file path of the text document in which you want to store the text of your word document (without quotes):");
            //string txtPath = Console.ReadLine();
            string txtPath = @"c:/5400BomNumber.txt"; //如果想自己从控制台输入 结果文件 路径,请解除上两行注释

            Word.Application app = new Word.Application();
            Word.Document doc;
            object missing = Type.Missing;
            object readOnly = true;
            try
            {
                doc = app.Documents.Open(ref path, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                string text = doc.Content.Text;
                text = text.Replace("\r", "").Replace("\n", "").Replace("\a", ""); //这里 \r, \n, \a都是来替代word转TXT后的ASCII转义字符,需要将他们替换为空
                string[] text1 = Regex.Split(text, "Sliding Stem", RegexOptions.IgnoreCase);
                string text_output = "";
                string output_5400 = "";

                #region 逐item转换

                for (int n = 1; n < text1.Length; n++)
                {
                    #region 第一步,将单张spec拆分为离散的字段

                    string Customer = getBetween("Sliding Stem Vavle Specification" + text1[n], "Customer:", "Contact:"); //Customer;
                    string Item = getBetween("Sliding Stem Vavle Specification" + text1[n], "Item:", "Rev:"); //Item
                    string Rev = getBetween("Sliding Stem Vavle Specification" + text1[n], "Rev:", "Qty:"); //Spec Rev
                    string Qty = getBetween("Sliding Stem Vavle Specification" + text1[n], "Qty:", "Quote:"); //Qty
                    string Tag = getBetween("Sliding Stem Vavle Specification" + text1[n], "Tags:", "Date Last Modified:"); //Tag
                    string Size_Type = getBetween("Sliding Stem Vavle Specification" + text1[n], "Size and Type:", "Input Signal:"); //Size and Type
                    string Body_Style = getBetween("Sliding Stem Vavle Specification" + text1[n], "Body Style:", "Access:"); //Bosy Style
                    string Design_Temp = getBetween("Sliding Stem Vavle Specification" + text1[n], "Design Temp:", "Gauges:"); //Design Temp
                    string Design_Pressure = getBetween("Sliding Stem Vavle Specification" + text1[n], "Design Press:", "Action:"); //Design Pressure
                    string End_Connection = getBetween("Sliding Stem Vavle Specification" + text1[n], "End Connect/In/Out:", "Certification:"); //End Connection
                    string Material = getBetween("Sliding Stem Vavle Specification" + text1[n], "Material:", "Controller Type:"); //Material
                    string Port = getBetween("Sliding Stem Vavle Specification" + text1[n], "Ports:", "Action:"); //Port
                    string Flow_Direction = getBetween("Sliding Stem Vavle Specification" + text1[n], "Flow Direction:", "Measure Element:"); //Flow direction
                    string Trim_Number = getBetween("Sliding Stem Vavle Specification" + text1[n], "Trim Number:", "Range:"); //Trim number
                    string Cage_Material = getBetween("Sliding Stem Vavle Specification" + text1[n], "Cage Matl:", "Output:"); //Cage material
                    string Retainer_Material = getBetween("Sliding Stem Vavle Specification" + text1[n], "Retainer Matl:", "Mounting:"); //Cage retainer mateial
                    string Bushing_Material = getBetween("Sliding Stem Vavle Specification" + text1[n], "Bushing Matl:", "Airset:"); //Bushing Material
                    string Seatring_Material = getBetween("Sliding Stem Vavle Specification" + text1[n], "Seat Ring Matl:", "Mounting:"); //seat ring material
                    string Plug_Material_1 = getBetween("Sliding Stem Vavle Specification" + text1[n], "Transducer:", "Guiding:"); //Plug material_1 is needed because there is another "Material" (body material) before this plug material, to get plug material , we firstly pick string between "transducer:" and "Guiding:", then pick string between "Material:" and "Input Singnal:"
                    string Plug_Material = getBetween(Plug_Material_1, "Material:", "Input Signal:"); //Plug material
                    string Guiding = getBetween("Sliding Stem Vavle Specification" + text1[n], "Guiding:", "Output Signal:"); //Plug gudiding method
                    string Balance = getBetween("Sliding Stem Vavle Specification" + text1[n], "Balance:", "Action:"); //trim balance method
                    string Shutoff = getBetween("Sliding Stem Vavle Specification" + text1[n], "Shutoff Class:", "Mounting:"); //shut off class
                    string Port_Size = getBetween("Sliding Stem Vavle Specification" + text1[n], "Port Size:", "Airset:"); //Port size
                    string Characteristic = getBetween("Sliding Stem Vavle Specification" + text1[n], "Characteristic:", "Certifications:"); //flow characteristic
                    string Stem_Material = getBetween("Sliding Stem Vavle Specification" + text1[n], "Stem Material:", "Line In:"); //Stem Material
                    string Stem_Size = getBetween("Sliding Stem Vavle Specification" + text1[n], "Stem Size:", "Line Out:"); //Stem size
                    string Bonnet_Style = getBetween("Sliding Stem Vavle Specification" + text1[n], "Bonnet Style:", "Insulation:"); //Bonnet style
                    string Boss_Size = getBetween("Sliding Stem Vavle Specification" + text1[n], "Boss Size:", "Service Cond:"); //Boss size
                    string Packing = getBetween("Sliding Stem Vavle Specification" + text1[n], "Packing:", "Process Fluid:"); //Packing
                    string Access = getBetween("Sliding Stem Vavle Specification" + text1[n], "Access:", "Critical Pressure:"); //Bonnet access
                    string Bolt_Bonnet = getBetween("Sliding Stem Vavle Specification" + text1[n], "Bolt, Bonnet:", "Shutoff Drop:"); //Bonnet Bolt
                    string Bolt_Flange_Packing = getBetween("Sliding Stem Vavle Specification" + text1[n], "PackFlg/Bltg:", "|"); //packing bolt and flange
                    string Actuator = getBetween("Sliding Stem Vavle Specification" + text1[n], "Actuator:", "|"); //Actuator
                    string Actuator_Type_Size = getBetween("Sliding Stem Vavle Specification" + text1[n], "Type/Size:", "|"); //actuator type and size
                    string Travel = getBetween("Sliding Stem Vavle Specification" + text1[n], "Travel:", "|"); // valve travel

                    #endregion 第一步,将单张spec拆分为离散的字段

                    #region 第二步,对复合字段进行进一步拆分

                    //判定客户
                    string Customer_name;
                    Customer_name = "";
                    string Note_Customer_name;
                    Note_Customer_name = "";
                    switch (Customer)
                    {
                        case "":
                            Note_Customer_name = "Customer is note specified";
                            break;

                        default:
                            Customer_name = Customer.Replace("\a", "") + "";
                            break;
                    }

                    //判定Item #
                    string Item_number;
                    Item_number = "";
                    string Note_Item_number;
                    Note_Item_number = "";
                    switch (Item)
                    {
                        case "":
                            Item_number = "Item # is not specified";
                            break;

                        default:
                            Item_number = Item.Replace("\a", "");
                            break;
                    }

                    //判定阀门型号 和 尺寸
                    string Valve_type; //阀门型号
                    Valve_type = "";
                    string Note_Valve_type;
                    Note_Valve_type = ""; //如果阀门型号不能判定,则给予note提示;

                    string Valve_size; //阀门尺寸
                    Valve_size = "";
                    string Note_Valve_size;
                    Note_Valve_size = "";

                    if (Size_Type.Contains("5400")) //5400
                    {
                        Valve_type = "Type 5400";
                        Valve_size = Size_Type.Replace("5400", "").Replace("\a", "").Replace(" ", "");
                    }
                    else if (Size_Type.Contains("5100")) //5100
                    {
                        Valve_type = "Type 5100";
                        Valve_size = Size_Type.Replace("5100", "").Replace("\a", "").Replace(" ", "");
                    }
                    else if (Size_Type.Contains("Type 5364")) //5364
                    {
                        Valve_type = "Type 5364";
                        Valve_size = Size_Type.Replace("5364", "").Replace("\a", "").Replace(" ", "");
                    }
                    else if (Size_Type.Contains("Type 5366")) //5366
                    {
                        Valve_type = "Type 5366";
                        Valve_size = Size_Type.Replace("5366", "").Replace("\a", "").Replace(" ", "");
                    }
                    /*
                    else if (Size_Type.Contains("Type xxxx")) // 增加产品型号请看这里!
                    {
                        Valve_type = "Type xxxx";
                        Valve_size = Size_Type.Replace("5400", "").Replace("\a", "").Replace(" ", "");
                    }
                    */
                    else
                    {
                        Note_Valve_type = "Cannot indentify product type";
                    }

                    //判定阀门设计温度
                    Double Design_temperature; //阀门设计温度 double
                    Design_temperature = GetNumber(Design_Temp, @"[a - zA - Z]");
                    string Design_temperatrue_unit = ""; //阀门温度单位
                    string Note_Design_temperature = "";
                    if (Design_Temp.Contains("C"))
                    {
                        Design_temperatrue_unit = "Degree C";
                    }
                    else if (Design_Temp.Contains("F"))
                    {
                        Design_temperature = (Design_temperature - 32) * 5 / 9;
                        Design_temperatrue_unit = "Degree C";
                    }
                    else
                    {
                        Note_Design_temperature = "temperatrue unit can not be recognized, please change the temperature unit to Degree F or Degree C";
                    }
                    //判定阀门设计压力
                    Double Design_pressure; //阀门设计压力 double
                    Design_pressure = GetNumber(Design_Pressure, @"[a - zA - Z]");
                    string Design_pressure_unit = GetLetter(Design_Pressure, @"[a - zA - Z]"); //阀门压力单位
                    string Note_Design_pressure = "";
                    if (Design_Pressure.ToUpper().Contains("MPA"))
                    {
                        Design_pressure_unit = "MPa";
                    }
                    else if (Design_Pressure.ToUpper().Contains("BAR"))
                    {
                        Design_pressure = Design_pressure / 10;
                        Design_pressure_unit = "MPa";
                    }
                    else if (Design_Pressure.ToUpper().Contains("KPA"))
                    {
                        Design_pressure = Design_pressure / 1000;
                        Design_pressure_unit = "MPa";
                    }
                    else if (Design_Pressure.ToUpper().Contains("PSI"))
                    {
                        Design_pressure = Design_pressure / 145.03773800722;
                        Design_pressure_unit = "MPa";
                    }
                    else
                    {
                        Design_pressure_unit = "NA";
                        Note_Design_pressure = "pressure unit can not be recognized, please change the temperature unit to Degree F or Degree C";
                    }

                    //压力等级
                    string Valve_class;
                    if (End_Connection.ToUpper().Contains("150") || End_Connection.ToUpper().Replace(" ", "").Contains("PN20"))
                    {
                        Valve_class = "CL150";
                    }
                    else if (End_Connection.ToUpper().Contains("300") || End_Connection.ToUpper().Replace(" ", "").Contains("PN50"))
                    {
                        Valve_class = "CL300";
                    }
                    else if (End_Connection.ToUpper().Contains("600") || End_Connection.ToUpper().Replace(" ", "").Contains("PN110"))
                    {
                        Valve_class = "CL600";
                    }
                    else if (End_Connection.ToUpper().Contains("400") || End_Connection.ToUpper().Replace(" ", "").Contains("PN68"))
                    {
                        Valve_class = "CL400";
                    }
                    else if (End_Connection.ToUpper().Contains("900") || End_Connection.ToUpper().Replace(" ", "").Contains("PN150"))
                    {
                        Valve_class = "CL900";
                    }
                    else if (End_Connection.ToUpper().Contains("1500") || End_Connection.ToUpper().Replace(" ", "").Contains("PN260"))
                    {
                        Valve_class = "CL1500";
                    }
                    else if (End_Connection.ToUpper().Contains("2500") || End_Connection.ToUpper().Replace(" ", "").Contains("PN420"))
                    {
                        Valve_class = "CL2500";
                    }
                    else if (End_Connection.ToUpper().Replace(" ", "").Contains("PN100"))
                    {
                        Valve_class = "PN100";
                    }
                    else if (End_Connection.ToUpper().Replace(" ", "").Contains("PN16"))
                    {
                        Valve_class = "PN16";
                    }
                    else if (End_Connection.ToUpper().Replace(" ", "").Contains("PN40"))
                    {
                        Valve_class = "PN40";
                    }
                    else if (End_Connection.ToUpper().Replace(" ", "").Contains("PN63"))
                    {
                        Valve_class = "PN63";
                    }
                    else if (End_Connection.ToUpper().Replace(" ", "").Contains("PN10"))
                    {
                        Valve_class = "PN10";
                    }
                    else if (End_Connection.ToUpper().Replace(" ", "").Contains("PN6"))
                    {
                        Valve_class = "PN6";
                    }
                    else if (End_Connection.ToUpper().Replace(" ", "").Contains("PN16"))
                    {
                        Valve_class = "PN16";
                    }
                    else if (End_Connection.ToUpper().Replace(" ", "").Contains("SCH") && End_Connection.ToUpper().Replace(" ", "").Contains("40"))
                    {
                        Valve_class = "SCHD40";
                    }
                    else if (End_Connection.ToUpper().Replace(" ", "").Contains("SCH") && End_Connection.ToUpper().Replace(" ", "").Contains("80"))
                    {
                        Valve_class = "SCHD80";
                    }
                    else
                    {
                        Valve_class = "";
                    }

                    //阀门端口
                    string Valve_end;
                    if (End_Connection.ToUpper().Contains("RF FLG/RF FLG"))
                    {
                        Valve_end = "RF";
                    }
                    else if (End_Connection.ToUpper().Contains("FF FLG/FF FLG"))
                    {
                        Valve_end = "FF";
                    }
                    else if (End_Connection.ToUpper().Contains("RTJ FLG/RTJ FLG"))
                    {
                        Valve_end = "RTJ";
                    }
                    else if (End_Connection.ToUpper().Contains("BWE"))
                    {
                        Valve_end = "BWE";
                    }
                    else if (End_Connection.ToUpper().Contains("BUTT"))
                    {
                        Valve_end = "BWE";
                    }
                    else
                    {
                        Valve_end = "N/A";
                    }

                    #endregion 第二步,对复合字段进行进一步拆分

                    #region 第三步,将各个字段分别转换为BOM number code

                    string code0 = ""; //special BOM alarm
                    string code1 = ""; //valve type
                    string code2 = ""; //valve size
                    string code3 = ""; //valve body material
                    string code4 = ""; //valve class
                    string code5 = ""; //valve seal type
                    string code6 = ""; //trim material
                    string code7 = ""; //flow characteristic
                    string code8 = ""; //shutoff
                    string code9 = ""; //packing
                    string code10 = ""; //Yoke boss
                    string code11 = ""; //travel
                    string code12 = ""; //optional, port type, reduced port?
                    string code13 = ""; //optional, port class
                    string code14 = ""; //optional, special bonnet
                    string code15 = ""; //optional, special body/bonnet bolt
                    string code16 = ""; //optional, speical packing flange/bolt/nut
                    string note1 = "";
                    string note2 = "";
                    string note3 = "";
                    string note4 = "";
                    string note5 = "";
                    string note6 = "";
                    string note7 = "";
                    string note8 = "";
                    string note9 = "";
                    string note10 = "";
                    string note11 = "";
                    string note12 = "";
                    string note13 = "";
                    string note14 = "";
                    string note15 = "";
                    string note16 = "";
                    switch (Valve_type)
                    {
                        default:
                            code1 = "(1111)-"; code0 = "Special !";
                            break;

                        case "Type 5400":
                            code1 = "5400-";
                            break;

                        case "Type 5100":
                            code1 = "5100-";
                            break;

                        case "Type 5364":
                            code1 = "5364-";
                            break;

                        case "Type 8100":
                            code1 = "8100-";
                            break;
                    }

                    if (Valve_type == "Type 5400")
                    {
                        /// <summary>
                        /// code 2
                        /// 目的 提取VALVE SIZE编码
                        /// 思路
                        /// 1. ref 第二步//判定阀门型号 和 尺寸
                        /// </summary>
                        switch (Valve_size)
                        {
                            default:
                                code2 = "(2)"; code0 = "Special !";
                                break;

                            case "NPS1":
                            case "DN25":
                                code2 = "1";
                                break;

                            case "NPS2":
                            case "DN50":
                                code2 = "2";
                                break;

                            case "NPS3":
                            case "DN80":
                                code2 = "3";
                                break;

                            case "NPS4":
                            case "DN100":
                                code2 = "4";
                                break;

                            case "NPS11/2":
                            case "DN40":
                                code2 = "5";
                                break;

                            case "NPS6":
                            case "DN150":
                                code2 = "6";
                                break;

                            case "NPS8":
                            case "DN200":
                                code2 = "8";
                                break;

                            case "NPS10X8":
                            case "DN250":
                                code2 = "A";
                                break;

                            case "NPS12":
                            case "DN300":
                                code2 = "C";
                                break;
                        }
                        /// <summary>
                        /// code 3
                        /// 目的 提取valve body编码
                        /// 思路
                        /// 1. 当端口为RF时, code3由Material确定;
                        /// </summary>
                        switch (Material.Trim())
                        {
                            default:
                                code3 = "(3)"; code0 = "Special !";
                                break;

                            case "WCC":
                            case "WCC Steel":
                                code3 = "W";
                                break;

                            case "CF8M":
                            case "CF8M SST":
                            case "CF8M Stainless Steel":
                                code3 = "S";
                                break;

                            case "LCC":
                            case "LCC Steel":
                                code3 = "L";
                                break;

                            case "CF3M":
                            case "CF3M SST":
                            case "CF3M Stainless Steel":
                                code3 = "T";
                                break;

                            case "CF3":
                            case "CF3 SST":
                            case "CF3 Stainless Steel":
                                code3 = "A";
                                break;

                            case "WC9":
                            case "WC9 Steel":
                                code3 = "H";
                                break;

                            case "WCC Steel 20B101":
                                code3 = "(3)"; code0 = "Special !"; note3 = "-3.NACE Requirement;";
                                break; //特例2.1 WCC NACE requirement
                        }

                        /// <summary>
                        /// code 4
                        /// 目的 提取Class编码
                        /// 思路
                        /// 1. 当端口为RF时, code4由valve class 确定;
                        /// 2. 当端口为BWE时, code4由 SCH确定;
                        /// 3.1 特例: 对于NPS 1, 1 1/2, 2,3 法兰端(RF), 当CL400时, 其Code 4与CL600时的CODE4相同, 此时Code4=C;
                        /// </summary>
                        if (Valve_end == "RF")
                        {
                            switch (Valve_class)
                            {
                                default:
                                    code4 = "(4)"; code0 = "Special !";
                                    break;

                                case "PN10":
                                    code4 = "1";
                                    break;

                                case "PN16":
                                    code4 = "2";
                                    break;

                                case "PN25":
                                    code4 = "3";
                                    break;

                                case "PN40":
                                    code4 = "4";
                                    break;

                                case "PN63":
                                    code4 = "5";
                                    break;

                                case "CL150":
                                    code4 = "A";
                                    break;

                                case "CL300":
                                    code4 = "B";
                                    break;

                                case "CL600":
                                    code4 = "C";
                                    break;
                            }
                            if (Valve_class == "CL400" && (code2 == "1" || code2 == "2" || code2 == "3" || code2 == "5")) //特例3.1
                            {
                                code4 = "C";
                            }
                        }
                        else if (Valve_end == "BWE")
                        {
                            switch (Valve_class)
                            {
                                default:
                                    code4 = "(4)"; code0 = "Special !";
                                    break;

                                case "SCHD40":
                                    code4 = "J";
                                    break;

                                case "SCHD80":
                                    code4 = "K";
                                    break;
                            }
                        }
                        else
                        {
                            code4 = "(4)";
                            code0 = "Special !";
                        }
                        /// <summary>
                        /// code 5
                        /// 目的 提取Plug Seal编码
                        /// 思路
                        /// 1. 确定Code5需要用到CODE8, 因此将其排在code8后面了;
                        /// </summary>

                        /// <summary>
                        /// code 6
                        /// 目的 提取Trim Material编码
                        /// 思路
                        /// 0. 先将Cage, plug, seat, Stem material 的字段去特殊符号和空格并转大写
                        /// 1. Trim material的选择依赖于Trim Number, 某一个trim number 就对应类一系列的cage, plug, seat ring , plug材料组合,如果Trim Number与材料不匹配,则认为spec有错误;
                        /// 2. 考虑design温度,如果温度超出了材料的适用范围,(参考工程师意见),则认为spec 有错误;
                        /// 3. IF语法为if((cage material判定)&&(plug material判定)&&(seatring material判定)&&(stem material判定)&&(design temperature判定)&&(trim number判定)&&(code2判定)), 其中code2 就是valve size, 各判定必须用()圈起来并用&&连接.
                        /// </summary>
                        Cage_Material = Cage_Material.ToString().Replace("\a", "").Replace(" ", "");
                        Plug_Material = Plug_Material.ToString().Replace("\a", "").Replace(" ", "");
                        Seatring_Material = Seatring_Material.Replace("\a", "").Replace(" ", "");
                        Stem_Material = Stem_Material.Replace("\a", "").Replace(" ", "");

                        if ((Cage_Material == ("S17400SST") || Cage_Material == ("CB7Cu-1SST")) && (Plug_Material == ("S41000SST") || Plug_Material == ("S41600SST")) && (Seatring_Material == ("S41600SST") || Seatring_Material == ("S41000SST")) && (Stem_Material == ("S31600SST")) && (Design_temperature <= 427 && Design_temperature >= 0) && Trim_Number.Contains("1") && (code2 != "1"))
                        {
                            code6 = "1";
                        }
                        else if ((Cage_Material == ("S31600SST") || Cage_Material == ("CF8MSST")) && (Plug_Material == ("316SST") || Plug_Material == ("S31600SST") || Plug_Material == ("CF8MSST")) && (Seatring_Material == ("S31600SST") || Seatring_Material == ("CF8MSST")) && (Stem_Material == ("S31600SST")) && (Design_temperature <= 150 && Design_temperature >= 0) && Trim_Number.Contains("2") && (code2 != "1"))
                        {
                            code6 = "2";
                        }
                        else if ((Cage_Material == ("S31600SST/ENC") || Cage_Material == ("316/ENC") || Cage_Material == ("CF8M/ENC")) && (Plug_Material == ("316SST") || Plug_Material == ("S31600SST") || Plug_Material == ("CF8MSST")) && (Seatring_Material == ("S31600SST") || Seatring_Material == ("CF8MSST")) && (Stem_Material == ("S31600SST")) && (Design_temperature <= 316 && Design_temperature >= 0) && Trim_Number.Contains("3") && (code2 != "1"))
                        {
                            code6 = "3";
                        }
                        else if ((Cage_Material == ("S31600SST/ENC") || Cage_Material == ("316/ENC") || Cage_Material == ("CF8M/ENC")) && (Plug_Material.Contains("S31600SST/CoCr-A") || Plug_Material.Contains("CF8M/CoCr-A")) && (Seatring_Material.Contains("S31600SST/CoCr-A") || Seatring_Material.Contains("CF8M/CoCr-A")) && (Stem_Material == ("S31600SST")) && (Design_temperature <= 343 && Design_temperature >= 0) && Trim_Number.Contains("4") && (code2 != "1"))
                        {
                            code6 = "4";
                        }
                        else if ((Cage_Material == ("R30006") || Cage_Material == ("ALLOY6")) && (Plug_Material.Contains("S31600SST/CoCr-A") || Plug_Material.Contains("CF8M/CoCr-A")) && (Seatring_Material.Contains("S31600SST/CoCr-A") || Seatring_Material.Contains("CF8M/CoCr-A")) && (Stem_Material == ("S31600SST")) && (Design_temperature <= 593 && Design_temperature >= 0) && Trim_Number.Contains("5") && (code2 != "1"))
                        {
                            code6 = "5";
                        }
                        else if ((Cage_Material == ("S17400SST") || Cage_Material == ("CB7Cu-1SST")) && (Plug_Material == ("316SST") || Plug_Material == ("S31600SST") || Plug_Material == ("CF8MSST")) && (Seatring_Material == ("S31600SST") || Seatring_Material == ("CF8MSST")) && (Stem_Material == ("S31600SST")) && Trim_Number.Contains("6") && (code2 == "1"))
                        {
                            code6 = "6";
                        }
                        else if ((Cage_Material == ("S17400SST") || Cage_Material == ("CB7Cu-1SST")) && (Plug_Material.Contains("S31600SST/CoCr-A") || Plug_Material.Contains("CF8M/CoCr-A")) && (Seatring_Material.Contains("S31600SST/CoCr-A") || Seatring_Material.Contains("CF8M/CoCr-A")) && (Stem_Material == ("S31600SST")) && Trim_Number.Contains("7") && (code2 == "1"))
                        {
                            code6 = "7";
                        }
                        else
                        {
                            code6 = "(6.Trim material issue)";
                            code0 = "Special !";
                            note6 = "-6.please check *trim material & *whether trim number matchs trim material combination * whether valve size matchs trim material combination *Temperature limit";
                        }

                        /// <summary>
                        /// code 7
                        /// 目的 提取流量特性编码
                        /// 思路
                        /// 1. 对于等百分比的流量特性, 其字段为Characteristic, 其值为Equal percent,因此将其值转变为大写后,如果包含"EQUAL PERCENT",则可以判定code7="E"
                        /// 2. 对于线性的流量特性, 其字段为Characteristic, 其值为Linear.因此将其值转变为大写后,如果包含"LINEAR",则可以判定code7="L"
                        /// 3. 对于CAV III, ONE STAGE的流量特性, 其字段为Characteristic, 其值为CAVITROL III, ONE STAGE,因此将其值转变为大写后,如果包含"CAVITROL III, ONE STAGE",则可以判定code7="C"
                        /// 4. 对于WHISPER I的流量特性, 其字段为Characteristic, 其值为WHISPER I,因此将其值转变为大写后,如果等于"WHISPER I",(因为要排除Whisper III),则可以判定code7="W"
                        /// 5. 对于WHISPER III的流量特性, 其字段为Characteristic, 其值为WHISPER III,因此将其值转变为大写后,如果包含"WHISPER III, ONE STAGE",则可以判定code7="V"
                        /// </summary>
                        if (Characteristic.ToUpper().Contains("EQUAL PERCENT"))
                        {
                            code7 = "E";
                        }
                        else if (Characteristic.ToUpper().Contains("LINEAR"))
                        {
                            code7 = "L";
                        }
                        else if (Characteristic.ToUpper().Contains("CAVITROL III, ONE STAGE"))
                        {
                            code7 = "C";
                        }
                        else if (Characteristic.ToUpper() == "WHISPER I")
                        {
                            code7 = "W";
                        }
                        else if (Characteristic.ToUpper().Contains("WHISPER III, ONE STAGE"))
                        {
                            code7 = "V";
                        }
                        else
                        {
                            code7 = "(7.Flow characteristic)";
                            code0 = "Special !";
                        }

                        /// <summary>
                        /// code 8
                        /// 目的 提取shut off编码
                        /// 思路
                        /// 1. 对于shut of II, 其字段为Shutoff, 其值为ANSI CL II 或者CLII,或者II,因此将其值转变为大写后,如果包含"CL II"或者包含"CLII"或者等于"II",则可以判定code8="2"
                        /// 2. 对于 III, IV, V, VI,TSO等级的shut off, 参考第一步
                        /// </summary>
                        if (Shutoff.ToUpper().Contains("CL II") || Shutoff.ToUpper().Contains("CLII") || Shutoff.ToUpper() == "II")
                        {
                            code8 = "2";
                        }
                        else if (Shutoff.ToUpper().Contains("CL III") || Shutoff.ToUpper().Contains("CLIII") || Shutoff.ToUpper() == "III")
                        {
                            code8 = "3";
                        }
                        else if (Shutoff.ToUpper().Contains("CL IV") || Shutoff.ToUpper().Contains("CLIV") || Shutoff.ToUpper() == "IV")
                        {
                            code8 = "4";
                        }
                        else if (Shutoff.ToUpper().Contains("CL V") || Shutoff.ToUpper().Contains("CLV") || Shutoff.ToUpper() == "V")
                        {
                            code8 = "5";
                        }
                        else if (Shutoff.ToUpper().Contains("CL VI") || Shutoff.ToUpper().Contains("CLVI") || Shutoff.ToUpper() == "VI")
                        {
                            code8 = "6";
                        }
                        else if (Shutoff.ToUpper().Contains("CL TSO") || Shutoff.ToUpper().Contains("CLTSO") || Shutoff.ToUpper() == "TSO")
                        {
                            code8 = "7";
                        }
                        else
                        {
                            code8 = "(8. Shutoff)";
                            code0 = "Special !";
                        }

                        /// <summary>
                        /// code 5
                        /// 目的 提取Plug Seal编码
                        /// 思路
                        /// 1. "Spring loaded seal ring": size limit: not for NPS1/1.5, Port limit: NA, Shutoff limit: IV,V,cage guided,temperature upto 232C Trim: NA;
                        /// 2. "PEEK": size limit: not for NPS1/1.5, Port limit: NA, Shutoff limit: V,cage guided,temperature rage 232-316C Trim: NA;
                        /// 2. "PEEK": size limit: for NPS 2/3/4, Port limit: NA, Shutoff limit: IV,cage guided,temperature rage 232-316C Trim: NA;
                        /// 3. "Single Graphite": size limit: not for NPS1/1.5, Port limit: NA, Shutoff limit: II,cage guided,temperature upto 593 Trim: not for CAV;
                        /// 4. "Multi-graphite": size limit: NPS6 /8, Port limit: >= 4 3/8, Shutoff limit: III,IV,cage guided,temperature upto 232-593 Trim: not for CAV;
                        /// 5. "No seal,cage guide": size limit: NPS1.5, Port limit: NA, Shutoff limit: IV,V,cage guided,temperature upto 427 Trim: not for CAV;
                        /// 6. "No seal, post guide": size limit: not for NPS1, Port limit: NA, Shutoff limit: IV,V,post guided,temperature upto 427 Trim: not for CAV;
                        /// </summary>
                        if (code2 == "1" && (code8 == "4" || code8 == "5") && (Design_temperature <= 427 && Design_temperature >= 0) && Guiding.ToUpper().Contains("POST") && !(Characteristic.ToUpper().Contains("CAV")))
                        {
                            code5 = "Z";
                        }
                        else if (code2 == "5" && (code8 == "4" || code8 == "5") && (Design_temperature <= 427 && Design_temperature >= 0) && Guiding.ToUpper().Contains("CAGE") && !(Characteristic.ToUpper().Contains("CAV")))
                        {
                            code5 = "N";
                        }
                        else if ((code2 != "1" && code2 != "5") && (code8 == "2") && (Design_temperature <= 593 && Design_temperature >= 0) && Guiding.ToUpper().Contains("CAGE") && !(Characteristic.ToUpper().Contains("CAV")))
                        {
                            code5 = "P";
                        }
                        else if ((code2 == "6" || code2 == "8") && (code8 == "3" || code8 == "4") && (Design_temperature <= 593 && Design_temperature > 232) && Guiding.ToUpper().Contains("CAGE") && !(Characteristic.ToUpper().Contains("CAV")))
                        {
                            code5 = "M";
                        }
                        else if ((code2 != "1" && code2 != "5") && (code8 == "5") && (Design_temperature <= 316 && Design_temperature > 232) && Guiding.ToUpper().Contains("CAGE") && !(Characteristic.ToUpper().Contains("CAV")))
                        {
                            code5 = "A";
                        }
                        else if ((code2 == "2" || code2 == "3" || code2 == "4") && (code8 == "4") && (Design_temperature <= 316 && Design_temperature > 232) && Guiding.ToUpper().Contains("CAGE") && !(Characteristic.ToUpper().Contains("CAV")))
                        {
                            code5 = "A";
                        }
                        else if ((code2 != "1" && code2 != "5") && (code8 == "3" || code8 == "4") && (Design_temperature <= 232 && Design_temperature > 0) && Guiding.ToUpper().Contains("CAGE"))
                        {
                            code5 = "T";
                        }
                        else
                        {
                            code5 = "(5. Plug seal)";
                            code0 = "Special !";
                        }

                        /// <summary>
                        /// code 9
                        /// 目的 提取Packing type编码
                        /// 思路
                        /// 1. 对于single PTFE Packing, 其字段为Packing, 其值为single PTFE,因此将其值转变为大写后,如果包含"SINGLE PTFE",则可以判定code9="P"
                        /// 2. 对于 Single Graphite, Double PTFE, Double Graphite, 参考第一步
                        /// </summary>
                        if (Packing.ToUpper().Contains("SINGLE PTFE") && (Design_temperature <= 232 && Design_temperature >= -46))
                        {
                            code9 = "P";
                        }
                        else if (Packing.ToUpper().Contains("DOUBLE PTFE") && (Design_temperature <= 232 && Design_temperature >= -46))
                        {
                            code9 = "B";
                        }
                        else if (Packing.ToUpper().Contains("SINGLE GRAPHITE") && (Design_temperature <= 538 && Design_temperature >= -198))
                        {
                            code9 = "G";
                        }
                        else if (Packing.ToUpper().Contains("DOUBLE GRAPHITE") && (Design_temperature <= 538 && Design_temperature >= -198))
                        {
                            code9 = "D";
                        }
                        else
                        {
                            code9 = "(9. Packing type)";
                            code0 = "Special !";
                        }
                        /// <summary>
                        /// code 10
                        /// 目的 提取Yoke boss编码
                        /// 思路
                        /// 1. 这个好简单啊, 看Boss size就好了;
                        /// </summary>
                        if (Boss_Size.Contains("2 1/8"))
                        {
                            code10 = "1";
                        }
                        else if (Boss_Size.Contains("2 13/16"))
                        {
                            code10 = "2";
                        }
                        else if (Boss_Size.Contains("3 9/16"))
                        {
                            code10 = "3";
                        }
                        else if (Boss_Size.Contains("5H"))
                        {
                            code10 = "4";
                        }
                        else
                        {
                            code10 = "(10. Yokeboss)";
                            code0 = "Special !";
                        }
                        /// <summary>
                        /// code 11
                        /// 目的 提取Travel编码
                        /// 思路
                        /// 1. 这个好简单啊, 看Travel就好了;
                        /// </summary>
                        Travel = Travel.Replace("/a", "").Replace(" ", "").ToUpper();
                        switch (Travel)
                        {
                            default:
                                code11 = "(11.Travel)"; code0 = "Special !";
                                break;

                            case "3/4INCH":
                            case "19MM":
                                code11 = "1";
                                break;

                            case "11/2INCH":
                            case "38MM":
                                code11 = "2";
                                break;

                            case "2INCH":
                            case "51MM":
                                code11 = "3";
                                break;

                            case "3INCH":
                            case "76MM":
                                code11 = "4";
                                break;

                            case "4INCH":
                            case "101.6MM":
                            case "102MM":
                                code11 = "5";
                                break;

                            case "51/2INCH":
                            case "140MM":
                                code11 = "6";
                                break;
                        }
                        /// <summary>
                        /// code 12
                        /// 目的 reduce port编码
                        /// 思路
                        /// 1. 这个暂时不选;8/1/2018
                        /// </summary>
                        Port_Size = Port_Size.Replace("/a", "").Replace(" ", "").ToUpper();
                        if ((code2 == "1") && (Port_Size.Contains("22")))
                        {
                            code12 = "T";
                        }
                        else if ((code2 == "5") && (Port_Size.Contains("33") || Port_Size.Contains("15/16IN")))
                        {
                            code12 = "T";
                        }
                        else if ((code2 == "2") && (Port_Size.Contains("47.6") || Port_Size.Contains("48") || Port_Size.Contains("17/8IN")))
                        {
                            code12 = "T";
                        }
                        else if ((code2 == "3") && (Port_Size.Contains("73") || Port_Size.Contains("27/8")))
                        {
                            code12 = "T";
                        }
                        else if ((code2 == "4") && (Port_Size.Contains("87.3") || Port_Size.Contains("87") || Port_Size.Contains("37/16IN")))
                        {
                            code12 = "T";
                        }
                        else if ((code2 == "6") && (Port_Size.Contains("177.8") || Port_Size.Contains("178") || Port_Size.Contains("7IN")))
                        {
                            code12 = "T";
                        }
                        else if ((code2 == "8") && (Port_Size.Contains("203.2") || Port_Size.Contains("203") || Port_Size.Contains("8IN")))
                        {
                            code12 = "T";
                        }
                        else if ((code2 == "C") && (Port_Size.Contains("279") || Port_Size.Contains("11IN")))
                        {
                            code12 = "T";
                        }
                        else if ((code2 == "A") && (Port_Size.Contains("203.2") || Port_Size.Contains("203") || Port_Size.Contains("8IN")))
                        {
                            code12 = "R";
                        }
                        else
                        {
                            code12 = "(12. reduce port)";
                            code0 = "Special !";
                        }
                        /// <summary>
                        /// code 13
                        /// 目的 阀口缩颈编码
                        /// 思路
                        /// 1. 这个暂时不选;8/1/2018
                        /// </summary>
                        if (code12 == "T")
                        {
                            code13 = "1";
                        }
                        else
                        {
                            code13 = "(13. reduce valve end size)";
                            code0 = "Special !";
                        }
                        /// <summary>
                        /// code 14
                        /// 目的 Bonnet style编码
                        /// 思路
                        /// 1. 也很简单
                        /// </summary>
                        Bonnet_Style = Bonnet_Style.Replace("/a", "").Replace(" ", "").ToUpper();
                        if (Bonnet_Style.Contains("PLAIN"))
                        {
                            code14 = "1";
                        }
                        else if (Bonnet_Style.Contains("BELLOWS"))
                        {
                            code14 = "5";
                        }
                        else if (Bonnet_Style.Contains("CRYO"))
                        {
                            code14 = "6";
                        }
                        else if (Bonnet_Style.Contains("1"))
                        {
                            code14 = "2";
                        }
                        else if (Bonnet_Style.Contains("2"))
                        {
                            code14 = "3";
                        }
                        else if (Bonnet_Style.Contains("3"))
                        {
                            code14 = "4";
                        }
                        else
                        {
                            code14 = "(14. bonnet style)";
                            code0 = "Special !";
                        }
                        /// <summary>
                        /// code 15
                        /// 目的 body bonnet bolting编码
                        /// 思路
                        /// 1. 也很简单
                        /// </summary>
                        Bolt_Bonnet = Bolt_Bonnet.Replace("/a", "").Replace(" ", "").ToUpper();
                        if (Bolt_Bonnet.Contains("B7") && Bolt_Bonnet.Contains("2H") && Bolt_Bonnet.Contains("NCF2"))
                        {
                            code15 = "1";
                        }
                        else if (Bolt_Bonnet.Contains("B7") && Bolt_Bonnet.Contains("7"))
                        {
                            code15 = "2";
                        }
                        else if (Bolt_Bonnet.Contains("B8MCL2") && Bolt_Bonnet.Contains("8M"))
                        {
                            code15 = "3";
                        }
                        else if (Bolt_Bonnet.Contains("B16") && Bolt_Bonnet.Contains("7"))
                        {
                            code15 = "4";
                        }
                        else
                        {
                            code15 = "(15. body/bonnet bolting)";
                            code0 = "Special !";
                        }
                        /// <summary>
                        /// code 16
                        /// 目的 Packing bolting编码
                        /// 思路
                        /// 1. 也很简单
                        /// </summary>
                        Bolt_Flange_Packing = Bolt_Flange_Packing.Replace("/a", "").Replace(" ", "").ToUpper();
                        if (Bolt_Flange_Packing.Contains("SSTPKG") && Bolt_Flange_Packing.Contains("SSTSTUDS&NUTS"))
                        {
                            code16 = "1";
                        }
                        else if (Bolt_Flange_Packing.Contains("STLPKG") && Bolt_Flange_Packing.Contains("SSTSTUDS&NUTS"))
                        {
                            code16 = "2";
                        }
                        else
                        {
                            code16 = "(16. packing flang/bolting)";
                            code0 = "Special !";
                        }

                        output_5400 += code0.ToString() + "<Item>" + Item_number.Trim() + "<Bom>" + code1.ToString() + code2.ToString() + code3.ToString() + code4.ToString() + code5.ToString() + code6.ToString() + code7.ToString() + code8.ToString() + code9.ToString() + code10.ToString() + code11.ToString() + code12.ToString() + code13.ToString() + code14.ToString() + code15.ToString() + code16.ToString() + "\r\n" + note1.ToString() + note2.ToString() + note3.ToString() + note4.ToString() + note5.ToString() + note6.ToString() + "\r\n" + "\r\n";
                    }
                    else { }

                    #endregion 第三步,将各个字段分别转换为BOM number code

                    //Console.WriteLine("Item: {0}, *** {1}****{2}\n", Item, Bolt_Bonnet, Bolt_Flange_Packing); //加上这句调试时间就变长了,为什么????
                }

                #endregion 逐item转换

                doc.Close();
                File.WriteAllText(txtPath, output_5400);
                Console.WriteLine("Converted!");
                Console.ReadKey();
            }
            catch
            {
                Console.WriteLine("An error occured. Please check the file path to your word document, and whether the word document is valid.");
                Console.ReadKey();
            }

            app.Quit();
            //finally
            //{
            //object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
            //APP.QUIT(REF SAVECHANGES, REF MISSING, REF MISSING);

            //app.Quit();
            //}
            Console.ReadKey();
        }

        #region Class, for main codes

        /// <summary>
        /// 对word转txt后，提取各字段的方法
        /// 目的 提取字符串中的特定字符
        /// 思路
        /// 1. 明确起始字符（串）位置
        /// 2. 明确结尾字符（串）位置
        /// 3. 用Substring函数返回上述两者之间的字符串
        /// 4. 如果无返回值，则返回空值
        /// </summary>
        /// <param name="strSource"></param>
        /// <param name="strStart"></param>
        /// <param name="strEnd"></param>
        /// <returns></returns>
        public static string getBetween(string strSource, string strStart, string strEnd) //字符串中查找的函数,配合string sex = getBetween("aaa" + text1[n], "aaa", "bbb");一起用
        {
            int Start, End;
            if (strSource.Contains(strStart) && strSource.Contains(strEnd))
            {
                Start = strSource.IndexOf(strStart, 0) + strStart.Length;
                End = strSource.IndexOf(strEnd, Start);
                return strSource.Substring(Start, End - Start);
            }
            else
            {
                return "";
            }
        }

        /// <summary>
        /// 对设计温度，设计压力的处理
        /// 目的是得到温度和压力的数值(double Value)，单位(string Letter)
        /// 思路
        /// 设计温度和设计压力在spec上的格式都是+/-123 abc(g)，也包括word转text带入的一些特殊字符，用\a替换        ///
        /// 1.去空格和特殊字符，将字符串变成+/-123abc(g)
        /// 2.描述取字母的正则表达式
        /// 3.将字母都替换掉，并转换为double, 得到+/-123，注意这一步中要去括号，因为一些压力单位是带括号的，例如MPa(g)
        /// 4.将第三步得到的+/-123替换掉，剩下的就是abc
        /// 5.return
        /// </summary>
        /// <param name="Sentence"></param>
        /// <param name="Pattern"></param>
        /// <returns></returns>
        public static string GetLetter(string Sentence, string Pattern) //提取字符串+/-123abc的带符号的字母abc;
        {
            Sentence = Sentence.Replace("\a", "").Replace(" ", ""); //将字符串格式转换为+/ -123abc
            Pattern = "[a-zA-Z]"; //取字母的正则表达式;
            string Number = Regex.Replace(Sentence, Pattern, ""); // 将字母都替换成空格，得到+/-123
            //Double value = Convert.ToDouble(Number);
            string Letter = Sentence.Replace(Number, ""); //将+/-123abc中的+/-123部分替换掉,剩下的就是字母abc;
            return Letter;
        }

        public static double GetNumber(string Sentence, string Pattern) //提取字符串+/-123abc的带符号的数字+/-123;
        {
            Sentence = Sentence.Replace("\a", "").Replace(" ", ""); //将字符串转换为+/-123abc
            Pattern = "[a-zA-Z]"; //取字母的正则表达式;
            string Number = Regex.Replace(Sentence.Replace("(", "").Replace(")", ""), Pattern, ""); // 去除字符串中的括号，并将字母都替换成空格，得到+/-123
            Double Value = Convert.ToDouble(Number);
            //string Letter = Sentence.Replace(Number, ""); //将+/-123abc中的+/-123部分替换掉,剩下的就是字母abc;
            return Value;
        }

        #endregion Class, for main codes
    }
}
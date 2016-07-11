using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Threading;
using SerialComm;
using RealTime;
using Test7;
using GetData;
using HeatBalance;
using System.Reflection;
using System.Drawing.Drawing2D;
using System.IO;
using System.Drawing.Imaging;
using System.Windows.Forms.DataVisualization.Charting;

namespace Test7
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            //thread = new Thread();
        }
        public static string Tinside;
        public static ATS ats = new ATS();
        public static GDAC gdac = new GDAC();
        public static HBC hbc = new HBC();
        public static RTIM rti = new RTIM();
        Thread thread;
        Color backColor = Color.Black;//指定绘制曲线图的背景色
        public void RealtimeShow_Load(object sender, EventArgs e)
        {
            ats = new ATS();
            gdac = new GDAC();
            hbc = new HBC();
            rti = new RTIM();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            ats.AutoTest();
            thread.Start();
            gdac.getdata();
            hbc.HeatBalance1();
            hbc.HeatBalance2();
            rti.realtimeimagemaker(chart1.Height,Color.Black);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            thread.Abort();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Tinside = textBox1.Text;
        }

        public void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //在窗体即将关闭之前中止线程 
            //    
        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

    }
}
//记录原始数据
namespace SerialComm
{
    public class ATS
    {
        StringBuilder sb;
        public static double[,] orgdata = {{65.645816,  53.886809,	30.196343,	27.46359,	6.889639,	14.914445,	5.061891,	3.2864528},
                                           {65.380627,	53.836263,	30.202275,	27.478527,	6.927845,	14.888123,	4.578617,	2.859075},
                                           {65.100596,	53.783736,	30.188434,	27.48251,	7.031067,	14.937237,	5.034804,	3.0833456},
                                           {64.823533,	53.688593,	30.189423,	27.465582,	6.982766,	15.038727,	4.389028,	2.4416801},
                                           {64.561313,	53.582547,	30.179536,	27.46359,	7.037579,	14.96668,	4.792881,	2.66826},
                                           {64.317894,	53.462626,	30.179536,	27.468569,	6.924914,	14.931268,	4.986745,	3.2405925},
                                           {64.121971,	53.336759,	30.146909,	27.495455,	7.071878,	14.936017,	4.880978,	2.8920397},
                                           {63.973544,  53.178186,  30.165694,	27.495455,	6.956391,	14.949448,	5.138139,	3.2835007},
                                           {63.915163,	53.094936,	30.132079,	27.486493,	6.933815,	14.903997,	5.35799,	3.6987371},
                                           {63.931985,	52.985917,	30.115272,	27.504417,	6.993946,	14.991236,	5.118543,	3.3362494},
                                           {64.034894,	52.867978,	30.133068,	27.504417,	6.982332,	14.981468,	4.805113,	2.9857286},
                                           {64.230817,	52.807522,	30.136034,	27.492467,	6.932729,	14.922585,	5.034,	    3.1592831},
                                           {64.471268,	52.778781,	30.138011,	27.483506,	7.044417,	14.976176,	5.473978,	3.5274529},
                                           {64.745362,	52.790674,	30.149875,	27.480518,	6.942606,	14.921771,	4.704446,	2.9415187},
                                           {65.045183,  52.849148,  30.158774,	27.478527,	6.99807	,   14.87442,	4.77045,	2.7157241},
                                           {65.349953,	52.944291,	30.160751,	27.477531,	6.817025,	14.940087,	4.989644,	3.2282126},
                                           {65.640868,	53.047364,	30.16174,	27.477531,	6.975494,	14.874962,	4.878925,	3.016662},
                                           {65.91991,	53.167284,  30.169649	,27.480518,	6.905594,	14.9196,	4.415523,	2.4704074},
                                           {66.169267,	53.295134,	30.182502,	27.478527,  6.884863,	14.887173,	5.156635,	3.1408725},
                                           {66.356284,	53.434876,	30.200298,	27.480518,	6.93121,	14.924485,	4.711006,	2.760672},
                                           {66.457214,	53.574618,	30.211174,	27.479523,	6.820933,	14.920278,	4.839501,	2.977896},
                                           {66.498774,  53.697512,  30.224026,	27.483506,	6.862938,	14.853526,	5.22067,	3.3798799},
                                           {66.463151,	53.81446,	30.235891,  27.471557,	6.927736,	14.884596,	4.840072,	2.804763},
                                           {66.3454,	53.862032,	30.268517,	27.466578,	6.936854,	14.833988,	4.754748,	3.0445241},
                                           {66.158382,	53.90663,	30.272472,	27.492467,	6.891592,	14.901148,	5.024688,	3.2736049},
                                           {65.917931,	53.951229,	30.237868,	27.498442,	6.868799,	14.919193,	4.892427,	2.9827924},
                                           {65.651753,	53.942309	,30.248743	,27.47554,	7.089895,	14.922585,	4.80361,	3.0553166},
                                           {65.23069826,53.39141533,30.1846257,	27.48169856,6.946775074,14.92278589,4.913150074,3.046490507}};
        public void AutoTest()
        {
            sb = new StringBuilder();
            Microsoft.Office.Interop.Excel.Application excel0 = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook0 = excel0.Workbooks.Add(true);
            Worksheet worksheet0 = (Worksheet)workbook0.Worksheets["sheet1"];
            //Worksheet worksheet0 =(Worksheet)workbook0.Worksheets.Add(Type.Missing,workbook0.Worksheet[1], 1, Type.Missing);
            Microsoft.Office.Interop.Excel.Application excel3 = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook3 = excel3.Workbooks.Add(true);
            Worksheet worksheet3 = (Worksheet)workbook3.Worksheets["sheet1"];
            //Worksheet worksheet3 =(Worksheet)workbook1.Worksheets.Add(Type.Missing,workbook1.Worksheet[1], 1, Type.Missing);
            worksheet0.Cells[1, 1] = " <油路进口> (C)";
            worksheet0.Cells[1, 2] = " <油路出口> (C)";
            worksheet0.Cells[1, 3] = " <水路进口> (C)";
            worksheet0.Cells[1, 4] = " <水路出口> (C)";
            worksheet0.Cells[1, 5] = " <油流量>m3/h";
            worksheet0.Cells[1, 6] = " <水流量>m3/h";
            worksheet0.Cells[1, 7] = " <总压降>kPa";
            worksheet0.Cells[1, 8] = " <管束压降>kPa";
            worksheet3.Cells[1, 1] = " <油流量>m3/h";
            worksheet3.Cells[1, 2] = " <油路入口>℃";
            worksheet3.Cells[1, 3] = " <油路出口>℃";
            worksheet3.Cells[1, 4] = " <水流量>m3/h ";
            worksheet3.Cells[1, 5] = " <水路进口> (C) ";
            worksheet3.Cells[1, 6] = " <水路出口> (C)";
            worksheet3.Cells[1, 7] = " 壳程热负荷kW";
            worksheet3.Cells[1, 8] = " 管程热负荷kW";
            worksheet3.Cells[1, 9] = " 热平衡偏差%";
            worksheet3.Cells[1, 10] = " 对数平均温差 ";
            worksheet3.Cells[1, 11] = " 总压降kPa ";
            worksheet3.Cells[1, 12] = " 管束压降kPa";
            worksheet3.Cells[1, 13] = " 总传热系数W/(m2.k)";
            worksheet3.Cells[1, 14] = " 管程传热系数W/(m2.k)";

            for (int i = 0; i < 25; i++)
            {
                for (int j = 0; j < 8; j++)
                {
                    worksheet0.Cells[i + 2, j + 1] = orgdata[i, j];
                }
            }
            excel0.Visible = false;
            excel0.DisplayAlerts = false;//不显示提示框
            workbook0.Close(true, @"E:\学习\毕设\软件\原始数据.xls", null);
            worksheet0 = null;
            workbook0 = null;
            excel0.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel0);
            excel0 = null;
            System.GC.Collect();
            excel3.Visible = false;
            excel3.DisplayAlerts = false;//不显示提示框
            workbook3.Close(true, @"E:\学习\毕设\软件\有效数据.xls", null);
            worksheet3 = null;
            workbook3 = null;
            excel3.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel3);
            excel3 = null;
            System.GC.Collect();
        }
    }
}

//采集比较、记录数据
namespace GetData
{
    public class GDAC
    {
        double[] Toil, Tw;
        double T1, T2; //温度均值
        int startnumber;
        public static double T301 = 0;
        public static double T302 = 0;
        public static double T303 = 0;
        public static double T304 = 0;
        public static double T305 = 0;
        public static double T306 = 0;
        public static double T307 = 0;
        public static double T308 = 0;
        public void getdata()
        {

            bool stableornot1 = false;
            bool stableornot2 = false;
            Toil = new double[3];
            Tw = new double[3];
            T1 = new double();
            T2 = new double();
            int a = 0;
            int b = 0;
            double T3011 = 0;
            double T3021 = 0;
            double T3031 = 0;
            double T3041 = 0;
            double T3051 = 0;
            double T3061 = 0;
            double T3071 = 0;
            double T3081 = 0;
            while (stableornot1 == false)
            {
                for (a = 0; a < 18; a++)
                {
                    //读取最后的三个温度值   
                    Toil[0] = ATS.orgdata[a, 1];
                    Toil[1] = ATS.orgdata[a + 1, 1];
                    Toil[2] = ATS.orgdata[a + 2, 1];
                    T1 = Toil[0] + Toil[2] + Toil[2];
                    T1 = T1 / 3;
                    if (Math.Abs((T1 - Toil[1])) / T1 < 0.05 && Math.Abs((T1 - Toil[2])) / T1 < 0.05 && Math.Abs((T1 - Toil[0])) / T1 < 0.05)
                    {
                        stableornot1 = true;
                        break;
                    }
                }
            }
            while (stableornot2 == false)
            {
                for (b = 0; b < 18; b++)
                {
                    Tw[0] = ATS.orgdata[b + 1, 3];
                    Tw[1] = ATS.orgdata[b + 2, 3];
                    Tw[2] = ATS.orgdata[b + 3, 3];
                    T2 = Tw[0] + Tw[1] + Tw[1];
                    T2 = T1 / 2;
                    if (Math.Abs((T2 - Tw[1])) / T2 < 0.05 && Math.Abs((T2 - Tw[2])) / T2 < 0.05 && Math.Abs((T2 - Tw[0])) / T2 < 0.05)
                    {
                        stableornot2 = true;
                        break;
                    }
                }
            }
            if (stableornot1 == true && stableornot2 == true)
            {
                //记录稳定后的数据平均值
                startnumber = a;
                MessageBox.Show("系统已达到稳定。");
                for (int i = startnumber; i < 25; i++)
                {

                    T3011 += ATS.orgdata[i, 0];
                    T301 = T3011 / (25 - startnumber);
                    T3021 += ATS.orgdata[i, 1];
                    T302 = T3021 / (25 - startnumber);
                    T3031 += ATS.orgdata[i, 2];
                    T303 = T3031 / (25 - startnumber);
                    T3041 += ATS.orgdata[i, 3];
                    T304 = T3041 / (25 - startnumber);
                    T3051 += ATS.orgdata[i, 4];
                    T305 = T3051 / (25 - startnumber);
                    T3061 += ATS.orgdata[i, 5];
                    T306 = T3061 / (25 - startnumber);
                    T3071 += ATS.orgdata[i, 6];
                    T307 = T3071 / (25 - startnumber);
                    T3081 += ATS.orgdata[i, 7];
                    T308 = T3081 / (25 - startnumber);
                }
            }

        }
    }
}

//热平衡核算
namespace HeatBalance
{
    public class HBC
    {
        int b = 2;
        public void HeatBalance1()
        {
            Microsoft.Office.Interop.Excel.Application excel1 = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook1 = excel1.Workbooks.Open(@"E:\学习\毕设\软件\水的物性参数.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Microsoft.Office.Interop.Excel.Application excel2 = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook2 = excel2.Workbooks.Open(@"E:\学习\毕设\软件\热平衡算表.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Worksheet worksheet1 = (Worksheet)workbook1.Worksheets["sheet1"];
            Worksheet worksheet2 = (Worksheet)workbook2.Worksheets["sheet1"];
            int iColums = excel1.Columns.Count;//列数      
            int iRows = excel1.Rows.Count;//行数          
            //定义二维数组存储 Excel 表中读取的数据          
            string[,] waterdata = new string[iRows, iColums];
            for (int i = 0; i < 50; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    //将Excel表中的数据存储到数组    
                    Range rang3 = (Range)worksheet1.Cells[i + 1, j + 1];
                    waterdata[i, j] = Convert.ToString(rang3.Value2);
                }
            }
            //根据输入的室内温度,选取相应睡得物理性质,并填入热平衡计算表中进行计算           
            for (int i = 0; i < 50; i++)
            {
                if (Form1.Tinside == waterdata[i, 0])
                {
                    worksheet2.Cells[6, 2] = worksheet1.Cells[i, 2];
                    worksheet2.Cells[7, 2] = worksheet1.Cells[i, 3];
                    break;
                }
            }
            worksheet2.Cells[1, 1] = GDAC.T301;
            worksheet2.Cells[1, 3] = GDAC.T302;
            worksheet2.Cells[1, 5] = GDAC.T303;
            worksheet2.Cells[1, 7] = GDAC.T304;
            worksheet2.Cells[1, 9] = GDAC.T305;
            worksheet2.Cells[1, 11] = GDAC.T306;
            worksheet2.Cells[1, 13] = GDAC.T307;
            worksheet2.Cells[1, 15] = GDAC.T308;
            excel1.Visible = false;
            excel1.DisplayAlerts = false;//不显示提示框
            workbook1.Close(true, @"E:\学习\毕设\软件\水的物性参数.xls", null);
            worksheet1 = null;
            workbook1 = null;
            excel1.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel1);
            excel1 = null;
            System.GC.Collect();
            excel2.Visible = false;
            excel2.DisplayAlerts = false;//不显示提示框
            workbook2.Close(true, @"E:\学习\毕设\软件\热平衡算表.xls", null);
            worksheet2 = null;
            workbook2 = null;
            excel2.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel2);
            excel2 = null;
            System.GC.Collect();
        }
        public void HeatBalance2()
        {
            Microsoft.Office.Interop.Excel.Application excel2 = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook2 = excel2.Workbooks.Open(@"E:\学习\毕设\软件\热平衡算表.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Microsoft.Office.Interop.Excel.Application excel3 = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook3 = excel3.Workbooks.Open(@"E:\学习\毕设\软件\有效数据.xls", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Worksheet worksheet2 = (Worksheet)workbook2.Worksheets["sheet1"];
            Worksheet worksheet3 = (Worksheet)workbook3.Worksheets["sheet1"];
            Range rang = (Range)worksheet2.Cells[36, 9];
            MessageBox.Show("热平衡偏差为：\n" + Convert.ToString(rang.Value2));
            if (Math.Abs(Convert.ToDouble(rang.Value2)) < 6)
            {
                for (int j = 1; j < 15; j++)
                {
                    worksheet3.Cells[b, j] = worksheet2.Cells[36, j];
                }
                b++;
                MessageBox.Show("本组数据有效。");
                //电磁阀开度增加一个单位
            }
            else
            {
                DialogResult da = MessageBox.Show("热平衡偏差太大,本组数据不可用！");
            }
            excel3.Visible = false;
            excel3.DisplayAlerts = false;//不显示提示框
            workbook3.Close(true, @"E:\学习\毕设\软件\有效数据.xls", null);
            worksheet3 = null;
            workbook3 = null;
            excel3.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel3);
            excel3 = null;
            excel2.Visible = false;
            excel2.DisplayAlerts = false;//不显示提示框
            workbook2.Close(true, @"E:\学习\毕设\软件\热平衡算表.xls", null);
            worksheet2 = null;
            workbook2 = null;
            excel2.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel2);
            excel2 = null;
            System.GC.Collect();
        }
    }
}

//连线
namespace RealTime
{
    public class RTIM
    {
        private int width;//要生成的曲线图的宽度 
        private int height;//要生成的曲线图的高度 
        private Bitmap currentImage;//当前要绘制的图片 
        private Color backColor;//图片背景色 
        private static int count = ATS.orgdata.GetLength(0); //取得记录数量 
        System.Windows.Forms.DataVisualization.Charting.Series series = new System.Windows.Forms.DataVisualization.Charting.Series();
        private  System.Windows.Forms.DataVisualization.Charting.Chart chart1;
        public void realtimeimagemaker(int height,Color backColor)
        {
            this.height = height;
            this.backColor = backColor;
            //初始化曲线上的所有点坐标 
            int wd = 80 + 20 * (count - 1);//记算图表宽度 
            if (wd < 1200) wd = 1200;//设置最小宽度为1200        
            series.BorderWidth = 3;          
            series.ShadowOffset = 2;
            for (int i = 0; i < count-2; i++)
            {

                chart1.Series["水路进口温度"].Points.AddY(ATS.orgdata[i,0]);
                chart1.Series["水路出口温度"].Points.AddY(ATS.orgdata[i, 1]);
                chart1.Series["油路进口温度"].Points.AddY(ATS.orgdata[i, 2]);
                chart1.Series["油路出口温度"].Points.AddY(ATS.orgdata[i, 3]);
                chart1.Series["总压降"].Points.AddY(ATS.orgdata[i, 4]);
                chart1.Series["管束压降"].Points.AddY(ATS.orgdata[i, 5]);
                chart1.Series["水流量"].Points.AddY(ATS.orgdata[i, 6]);
                chart1.Series["油流量"].Points.AddY(ATS.orgdata[i, 7]);
            }
            // Set series chart type 
            chart1.Series["水路进口温度"].ChartType = SeriesChartType.Line;
            chart1.Series["水路出口温度"].ChartType = SeriesChartType.Line;
            chart1.Series["油路进口温度"].ChartType = SeriesChartType.Line;
            chart1.Series["油路出口温度"].ChartType = SeriesChartType.Line;
            chart1.Series["总压降"].ChartType = SeriesChartType.Line;
            chart1.Series["管束压降"].ChartType = SeriesChartType.Line;
            chart1.Series["水流量"].ChartType = SeriesChartType.Line;
            chart1.Series["油流量"].ChartType = SeriesChartType.Line;
            // Set point labels 
            chart1.Series["水路出口温度"].IsValueShownAsLabel = true;
            chart1.Series["Series2"].IsValueShownAsLabel = true;
            chart1.Series["Series3"].IsValueShownAsLabel = true;
            chart1.Series["Series4"].IsValueShownAsLabel = true;
            chart1.Series["Series5"].IsValueShownAsLabel = true;
            chart1.Series["Series6"].IsValueShownAsLabel = true;
            chart1.Series["Series7"].IsValueShownAsLabel = true;
            chart1.Series["Series8"].IsValueShownAsLabel = true;
            // Enable X axis margin 
            chart1.ChartAreas["ChartArea1"].AxisX.IsMarginVisible = true;
            chart1.Series["水路出口温度"]["ShowMarkerLines"] = "True";
            chart1.Series["Series2"]["ShowMarkerLines"] = "True";
            chart1.Series["Series3"]["ShowMarkerLines"] = "True";
            chart1.Series["Series4"]["ShowMarkerLines"] = "True";
            chart1.Series["Series5"]["ShowMarkerLines"] = "True";
            chart1.Series["Series6"]["ShowMarkerLines"] = "True";
            chart1.Series["Series7"]["ShowMarkerLines"] = "True"; 
            chart1.Series["Series8"]["ShowMarkerLines"] = "True";
        }
    }
}
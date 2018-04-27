using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports;

namespace csv
{
    public partial class Common : Form
    {
        string str;
        int wendu, shidu, jiedian;
        DataRow newRow;//创建新的单元行
        //实例化一个SerialPort
        private SerialPort ComDevice = new SerialPort();
        OpenFileDialog op = new OpenFileDialog();//实例化打开对画框。
        int k;
        bool isfrist = true;
        public DataTable data = new DataTable();
        public string path; //保存打开的路径
        public Common()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

     



        private void button1_Click(object sender, EventArgs e)
        {
              
            if (op.ShowDialog() == DialogResult.OK)//选择的文件名有效  
            {
                path = op.FileName;
                //创建读这个流的对象，第一个参数是文件流  
                data = OpenCSV(op.FileName);//取出datatable
            }

        }

        /// <summary>
        /// 将CSV文件的数据读取到DataTable中
        /// </summary>
        /// <param name="fileName">CSV文件路径</param>
        /// <returns>返回读取了CSV数据的DataTable</returns>
        public static DataTable OpenCSV(string filePath)
        {

            DataTable dt = new DataTable();

            StreamReader sr = new StreamReader(filePath, Encoding.Default);

            //记录每次读取的一行记录
            string strLine = "";
            //记录每行记录中的各字段内容
            string[] aryLine = null;
            string[] tableHead = null;
            //标示列数
            int columnCount = 0;
            //标示是否是读取的第一行
            bool IsFirst = true;
            //逐行读取CSV中的数据
            while ((strLine = sr.ReadLine()) != null)
            {
                //strLine = Common.ConvertStringUTF8(strLine, encoding);
                //strLine = Common.ConvertStringUTF8(strLine);

                if (IsFirst == true)
                {
                    tableHead = strLine.Split(',');
                    IsFirst = false;
                    columnCount = tableHead.Length;
                    //创建列
                    for (int i = 0; i < columnCount; i++)
                    {
                        DataColumn dc = new DataColumn(tableHead[i]);
                        dt.Columns.Add(dc);
                    }
                }
                else
                {
                    aryLine = strLine.Split(',');
                    DataRow dr = dt.NewRow();
                    for (int j = 0; j < columnCount; j++)
                    {
                        dr[j] = aryLine[j];
                    }
                    dt.Rows.Add(dr);
                }
            }
            if (aryLine != null && aryLine.Length > 0)
            {
                dt.DefaultView.Sort = tableHead[0] + " " + "asc";
            }

            sr.Close();
            return dt;
        }

        /// <summary>
        /// 将DataTable中数据写入到CSV文件中
        /// </summary>
        /// <param name="dt">提供保存数据的DataTable</param>
        /// <param name="fileName">CSV的文件路径</param>
        public static void SaveCSV(DataTable dt, string fullPath)
        {
            FileInfo fi = new FileInfo(fullPath);
            if (!fi.Directory.Exists)
            {
                fi.Directory.Create();
            }
            FileStream fs = new FileStream(fullPath, System.IO.FileMode.Create, System.IO.FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.UTF8);
            string data = "";
            //写出列名称
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                data += dt.Columns[i].ColumnName.ToString();
                if (i < dt.Columns.Count - 1)
                {
                    data += ",";
                }
            }
            sw.WriteLine(data);
            //写出各行数据
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                data = "";
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    string str = dt.Rows[i][j].ToString();
                    str = str.Replace("\"", "\"\"");//替换英文冒号 英文冒号需要换成两个冒号
                    if (str.Contains(',') || str.Contains('"')
                        || str.Contains('\r') || str.Contains('\n')) //含逗号 冒号 换行符的需要放到引号中
                    {
                        str = string.Format("\"{0}\"", str);
                    }

                    data += str;
                    if (j < dt.Columns.Count - 1)
                    {
                        data += ",";
                    }
                }
                sw.WriteLine(data);
            }
            sw.Close();
            fs.Close();
            DialogResult result = MessageBox.Show("文件保存成功！");
            //if (result == DialogResult.OK)
            //{
            //    System.Diagnostics.Process.Start("explorer.exe", fullPath);
            //}
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveCSV(data, path);
        }

        private void bt_open_Click(object sender, EventArgs e)
        {
            if (ComDevice.IsOpen == false)
            {
                //串口号
                //获取串口号
                string chuankou = comboBox1.Text;
                ComDevice.PortName = chuankou;
                //波特率
                ComDevice.BaudRate = Convert.ToInt32("115200");
                //设置奇偶校验检查协议
                ComDevice.Parity = Parity.None;
                //每个字节的标准数据位长度
                ComDevice.DataBits = 8;
                //设置每个字节的标准停止位数
                ComDevice.StopBits = StopBits.One;
                try
                {
                    ComDevice.Open();
                    //绑定事件
                    if(isfrist)
                    ComDevice.DataReceived += new SerialDataReceivedEventHandler(Com_DataReceived);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                bt_open.Text = "关闭串口";
            }
            else
            {
                try
                {
                    ComDevice.Close();
                    isfrist = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                bt_open.Text = "打开串口";
            }
        }


        /// <summary>
        /// 接收数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Com_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            str += ComDevice.ReadExisting();//字符串方式读
            if (str.Length == 8)
            {
                k = int.Parse(str);
                detail(k);
                str = "";
            }
            /* int wendu, shidu, jiedian;
             wendu =k/10000;//取前两位
             shidu = k&10000/100;//取中间两位
             jiedian = k%10;//取最后一位
             str = DateTime.Now.ToLocalTime().ToString() + " 节点：" + jiedian + " 湿度：" + shidu + " 温度" + wendu + "\n";
             richTextBox1.AppendText(str);*/
            //changeData(int.Parse(jiedian), int.Parse(shidu), int.Parse(wendu));
        }

        /// <summary>
        /// 添加行到DataTable
        /// </summary>
        /// <param name="id">节点号</param>
        /// <param name="shidu">湿度</param>
        /// <param name="wendu">温度</param>
        public void changeData(int id, int shidu, int wendu)
        {

            if (id == 1)
            {
                label5.Text = shidu.ToString();
                label3.Text = wendu.ToString();
            }
            else if (id == 2)
            {
                label6.Text = shidu.ToString();
                label8.Text = wendu.ToString();
            }
           
            newRow = data.NewRow();//把data的行格式给newRow

            newRow["节点"] = id;
            newRow["时间"] = DateTime.Now.ToLocalTime().ToString();
            newRow["温度"] = wendu;
            newRow["湿度"] = shidu;

            data.Rows.Add(newRow);

        }

        public void detail(int k)
        {
            wendu = k / 100000;//取前两位
            shidu = k % 100000 / 1000;//取中间两位
            jiedian = k % 10;//取最后一位
            richTextBox1.Text += DateTime.Now.ToLocalTime().ToString() + " 节点：" + jiedian + " 湿度：" + shidu + " 温度:" + wendu + "\n";
            changeData(jiedian, shidu, wendu);
        }
    }
}

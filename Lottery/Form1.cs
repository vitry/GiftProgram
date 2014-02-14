using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Font = System.Drawing.Font;
using Point = System.Drawing.Point;

namespace myLottery
{
    public partial class Form1 : Form
    {
        private ArrayList _controls;
        private Boolean _isLock = true;
        private ArrayList _jiangxiangArr; //奖项列表
        private ArrayList _liwaiArr; //例外，被排除的人员
        private ArrayList _renyuanArr; //人员列表

        private Boolean _stop = true;
        private Boolean _bgName = true;
        private int _cIndex;
        private String _dataName = "";
        private Application _excel; //
        private Form _f;
        private Worksheet _prizeWs;
        private long _lastActiveTime;
        private Worksheet _excludeListWs;
        private String _path = ""; //程序运行的路径
        private Worksheet _peopleWs;
        private Worksheet _logWs;
        private int _screenHeight; //电脑屏幕的高度
        private int _screenWidth; //电脑屏幕的宽度
        private Thread _tCj;
        private int _type; //几等奖
        private const XlFileFormat Version = XlFileFormat.xlExcel8; //2003版本
        private Workbook _wb; //data.xls文件

        public Form1()
        {
            InitializeComponent();
            Text = "出口易2013年会抽奖现场";
        }

        public override sealed string Text
        {
            get { return base.Text; }
            set { base.Text = value; }
        }

        //独立进程随机抽奖
        private void Run()
        {
            try
            {
                while (true)
                {
                    //返回0到人数总数的整数值
                    if (_stop)
                        break;
                    if (!_isLock)
                        _cIndex = new Random().Next(_renyuanArr.Count);
                    Thread.Sleep(100);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                SetStyle(ControlStyles.DoubleBuffer, true); //   设置双缓冲，防止图像抖动
                SetStyle(ControlStyles.AllPaintingInWmPaint, true); //   忽略系统消息，防止图像闪烁
                //this.SuspendLayout();
                _lastActiveTime = DateTime.Now.Ticks;
                //初始化设置
                //取得电脑屏幕的宽度和高度
                _screenWidth = Screen.GetBounds(this).Width;
                _screenHeight = Screen.GetBounds(this).Height;

                //程序运行路径
                _path = System.Windows.Forms.Application.StartupPath + "\\";
                //Excel文件完整路径
                _dataName = _path + "data.xls";
                //
                //this.BackgroundImage = Image.FromFile(path + "bg.jpg");
                BackgroundImageLayout = ImageLayout.Stretch;
                _bgName = true;

                BackColor = Color.FromArgb(196, 14, 10);
                String color = "";
                try
                {
                    using (var sr = new StreamReader(_path + "setting.txt"))
                    {
                        String line;
                        while ((line = sr.ReadLine()) != null)
                        {
                            if (line.StartsWith("color"))
                                color = line;
                        }
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("未找到文件setting.txt");
                }
                if (!color.Equals(""))
                {
                    color = color.Substring(6);
                    String[] arr = color.Split(',');
                    if (arr.Length == 2)
                        BackColor = Color.FromArgb(Convert.ToInt32(arr[0]), Convert.ToInt32(arr[1]),
                            Convert.ToInt32(arr[2]));
                }

                //设置背景透明
                labelMsg.AutoSize = true;
                labelMsg.Text = "";
                labelMsg.Font = new Font("LiSu", 16F, FontStyle.Bold, GraphicsUnit.Point, 134);
                labelMain.Text = "";
                labelMain.AutoSize = true;
                labelMain.Size = new Size(400, 60);
                labelMain.Location = new Point(Convert.ToInt32(_screenWidth/2) - 200,
                    Convert.ToInt32(_screenHeight/2) - 30);

                Location = new Point(Convert.ToInt32(_screenWidth/2) - Convert.ToInt32(Width/2),
                    Convert.ToInt32(_screenHeight/2) - Convert.ToInt32(Height/2));

                _renyuanArr = new ArrayList();
                _jiangxiangArr = new ArrayList();
                _liwaiArr = new ArrayList();

                //处理Excel表数据
                _excel = new Application {Visible = false, DisplayAlerts = false};
                _wb = _excel.Workbooks.Open(_dataName);

                Range range,range2 ,range3 ,range4;

                _peopleWs = (Worksheet) _wb.Worksheets[1]; //人员表
                _prizeWs = (Worksheet) _wb.Worksheets[2]; //奖项表
                _logWs = (Worksheet) _wb.Worksheets[3]; //日志表
                _excludeListWs = (Worksheet) _wb.Worksheets[4]; //例外表

                Worksheet s = _excludeListWs;
                for (int i = 2; i < s.UsedRange.Rows.Count + 1; i++)
                {
                    range = s.Range[s.Cells[i, 1], s.Cells[i, 1]];
                    if (range.Value2 != null && range.Value2.ToString() != "")
                        _liwaiArr.Add(range.Value2.ToString());
                    else
                        break;
                }

                s = _peopleWs;
                for (var i = 2; i < s.UsedRange.Rows.Count + 1; i++)
                {
                    range = s.Range[s.Cells[i, 1], s.Cells[i, 1]];
                    range2 = s.Range[s.Cells[i, 2], s.Cells[i, 2]];
                    range3 = s.Range[s.Cells[i, 3], s.Cells[i, 3]];
                    range4 = s.Range[s.Cells[i, 4], s.Cells[i, 4]];

                    if (range.Value2.ToString() != "" && range2.Value2.ToString() != "" &&
                        !_liwaiArr.Contains(range.Value2.ToString()))
                    {
                        var renyuan = new People
                        {
                            Number = range.Value2.ToString(),
                            Name = range2.Value2.ToString()
                        };
                        if (range3.Value2 != null)
                            renyuan.Department = range3.Value2.ToString();
                        if (range4.Value2 != null)
                            renyuan.OP = Convert.ToInt32(range4.Value2.ToString());
                        _renyuanArr.Add(renyuan);
                    }
                }

                //打乱顺序
                int num = 0; //动态随机数种子
                var temp = new ArrayList();
                while (_renyuanArr.Count > 0)
                {
                    int listIndex = new Random(new Random().Next(1000)*num*215).Next(_renyuanArr.Count);
                    temp.Add(_renyuanArr[listIndex]);
                    _renyuanArr.RemoveAt(listIndex);
                    num++;
                }
                _renyuanArr = new ArrayList(temp);
                temp.Clear();


                s = _prizeWs;
                for (int i = 2; i < s.UsedRange.Rows.Count + 1; i++)
                {
                    range = s.Range[s.Cells[i, 1], s.Cells[i, 1]];
                    range2 = s.Range[s.Cells[i, 2], s.Cells[i, 2]];
                    range3 = s.Range[s.Cells[i, 3], s.Cells[i, 3]];
                    range4 = s.Range[s.Cells[i, 4], s.Cells[i, 4]];
                    Range range5 = s.Range[s.Cells[i, 5], s.Cells[i, 5]];
                    Range range6 = s.Range[s.Cells[i, 6], s.Cells[i, 6]];

                    if (range.Value2.ToString() != "" && range2.Value2.ToString() != "")
                    {
                        _jiangxiangArr.Add(new Prize
                        {
                            TypeName = range.Value2.ToString(),
                            Count = Convert.ToInt32(range2.Value2.ToString()),
                            Current = Convert.ToInt32(range3.Value2.ToString()),
                            Yicichouqu = Convert.ToInt32(range4.Value2.ToString()),
                            Picpath = range5.Value2.ToString(),
                            JpIcon = range6.Value2.ToString(),
                        });
                    }
                    else
                        break;
                }
                _type = _jiangxiangArr.Count - 1;

                int level = 0;

                int startLocationY = Convert.ToInt32(_screenHeight/2) + 150;
                int startLocationX = Convert.ToInt32(_screenWidth/2 - 113*6/2);
                for (int i = 0; i < _jiangxiangArr.Count; i++)
                {
                    var jiangxiang = (Prize) _jiangxiangArr[i];
                    var btn = new PictureBox();
                    if (i > 0 && i%6 == 0)
                        level++;
                    btn.Location = new Point(startLocationX + (i - 6*level)*113, startLocationY + level*38);
                    btn.Name = "btn" + i;
                    btn.Size = new Size(110, 34);
                    btn.Image = Image.FromFile(_path + jiangxiang.Picpath);
                    btn.SizeMode = PictureBoxSizeMode.StretchImage;
                    btn.BackColor = Color.Transparent;
                    btn.Click += btn_Click;
                    btn.Cursor = Cursors.Hand;
                    Controls.Add(btn);
                    //


                    var img = new PictureBox
                    {
                        Name = "img" + i,
                        Size = new Size(100, 80),
                        Image = Image.FromFile(_path + jiangxiang.JpIcon),
                        BorderStyle = BorderStyle.FixedSingle,
                        SizeMode = PictureBoxSizeMode.StretchImage,
                        BackColor = Color.Transparent
                    };

                    img.Click += btn_Click;
                    img.Cursor = Cursors.Hand;
                    if (i < 6)
                        img.Location = new Point(30, i * 110 + 50);
                    else
                        img.Location = new Point(_screenWidth - 140, (i - 6) * 110 + 50);

                    Controls.Add(img);
                }

                FullScreen();
                //
                _controls = new ArrayList();
                foreach (Control con in Controls)
                {
                    if (con.Name.StartsWith("btn") || con.Name.StartsWith("img"))
                        _controls.Add(con);
                }
                ResumeLayout(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Form1_Closed(object sender, FormClosedEventArgs e)
        {
            try
            {
                _excel.Quit();
                _stop = true;
                if (_tCj != null)
                    _tCj.Abort();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (_renyuanArr.Count == 0)
                    return;
                SuspendLayout();
                if (e.KeyValue == 32 || e.KeyValue == 33 || e.KeyValue == 34)
                {
                    if (_f != null)
                        _f.Close();

                    ShowCon(false);
                    if (_bgName)
                    {
                        BackgroundImage = Image.FromFile(_path + "bg2.jpg");
                        BackgroundImageLayout = ImageLayout.Stretch;
                        _bgName = false;
                    }

                    var prize = (Prize) _jiangxiangArr[_type];

                    //如果已抽取数量大于或等于总数量将不抽取
                    if (prize.Current >= prize.Count)
                        return;

                    if (_stop)
                    {
                        //开始抽奖
                        try
                        {
                            if (DateTime.Now.Ticks - _lastActiveTime > 50000000)
                            {
                                _lastActiveTime = DateTime.Now.Ticks;

                                labelMain.Text = "";
                                _stop = false;
                                _tCj = new Thread(Run);
                                _tCj.Start();

                                timer1.Interval = 100;
                                timer1.Start();

                                labelMain.Font = new Font("LiSu", 54F, FontStyle.Bold, GraphicsUnit.Point, 134);
                                //
                                _isLock = false;
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                    }
                    else
                    {
                        //停止抽奖
                        try
                        {
                            if (DateTime.Now.Ticks - _lastActiveTime > 20000000)
                            {
                                _lastActiveTime = DateTime.Now.Ticks;

                                labelMain.Text = "";
                                _isLock = true;
                                Extract(prize, prize.Yicichouqu);
                                labelMsg.Text = prize.TypeName + " 总数量 " + prize.Count + " 已抽取 " +
                                                prize.Current + " 每次抽取 " + prize.Yicichouqu;
                                timer1.Stop();
                                _stop = true;
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                    }
                }
                //全屏或推出全屏 F11
                if (e.KeyValue == 122)
                    FullScreen();

                if (e.KeyValue == 27)
                {
                    labelMain.Text = "";
                    labelMain.AutoSize = true;
                    if (!_bgName)
                    {
                        var resources = new ComponentResourceManager(typeof (Form1));
                        BackgroundImage = ((Image) (resources.GetObject("$this.BackgroundImage")));
                        BackgroundImageLayout = ImageLayout.Stretch;
                        _bgName = true;
                    }
                    ShowCon(true);
                }
                ResumeLayout(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Form1_DoubleClick(object sender, EventArgs e)
        {
            //mFullScreen();
        }

        private void ShowCon(bool show)
        {
            try
            {
                var c = (Control) _controls[0];
                if (c.Visible != show)
                {
                    foreach (Control con in _controls)
                        con.Visible = show;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void FullScreen()
        {
            try
            {
                SuspendLayout();
                if (WindowState == FormWindowState.Maximized)
                {
                    WindowState = FormWindowState.Normal;
                    FormBorderStyle = FormBorderStyle.Sizable;
                }
                else
                {
                    FormBorderStyle = FormBorderStyle.None;
                    WindowState = FormWindowState.Maximized;
                }
                SuspendLayout();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        //抽取指定的人数
        private void Extract(Prize prize, int i)
        {
            try
            {
                if (_renyuanArr.Count < 1)
                {
                    timer1.Stop();
                    return;
                }
                String str = "";
                int num = 1;
                int j = i;
                while (j > 0)
                {
                    _cIndex = new Random(new Random().Next(num*1835*(j + 1000))).Next(_renyuanArr.Count);
                    var renyuan = (People) _renyuanArr[_cIndex];
                    //Console.WriteLine("cIndex: " + cIndex + " type: " + type + "   renyuan.getOp():" + renyuan.getOp());
                    //中奖人员权限限制
                    //if (type >= renyuan.getOp())
                    if (_type >= -1)
                    {
                        //将中奖人员工号写入日志，格式是时间+奖项+工号
                        String now = DateTime.Now.ToLocalTime().ToString(CultureInfo.InvariantCulture);
                        int r = _logWs.UsedRange.Rows.Count;
                        _logWs.Cells[r + 1, 1] = now;
                        _logWs.Cells[r + 1, 2] = prize.TypeName;
                        _logWs.Cells[r + 1, 3] = "'" + renyuan.Number;
                        _logWs.Cells[r + 1, 4] = renyuan.Name;
                        _logWs.Cells[r + 1, 5] = renyuan.Department;

                        String tmp = FillNumber(renyuan.Number) + " " + FillName(renyuan.Name) + " " +
                                     renyuan.Department;
                        //更新奖项记录
                        _prizeWs.Cells[_type + 2, 3] = prize.Current + 1;
                        if (i == 1)
                            str = tmp;
                        else if (i <= 10)
                        {
                            str = str + tmp + "\n";
                        }
                        else
                        {
                            str = str + tmp;
                            if ((j - 1)%2 == 0)
                                str = str + "\n";
                        }

                        //将已中奖工号写入列外Sheet
                        _excludeListWs.Cells[_excludeListWs.UsedRange.Rows.Count + 1, 1] = "'" + renyuan.Number;


                        _wb.SaveAs(_dataName, Version, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing);

                        if (prize.Count - prize.Current > 0)
                            prize.Current =prize.Current + 1;
                        _jiangxiangArr[_type] = prize;
                        //将已中奖工号从renyuanArr列表中移除
                        _renyuanArr.RemoveAt(_cIndex);
                        //

                        j--;

                        num++;

                        if (i == 1)
                        {
                            labelMain.Font = new Font("LiSu", 46F, FontStyle.Bold, GraphicsUnit.Point, 134);
                        }
                        else if (i <= 5 && i > 1)
                        {
                            labelMain.Font = new Font("LiSu", 46F, FontStyle.Bold, GraphicsUnit.Point, 134);
                        }
                        else if (i <= 10 && i > 5)
                        {
                            labelMain.Font = new Font("LiSu", 46F, FontStyle.Bold, GraphicsUnit.Point, 134);
                        }
                        else
                        {
                            labelMain.Font = new Font("LiSu", 20F, FontStyle.Bold, GraphicsUnit.Point, 134);
                        }

                        //如果抽奖数量已经达到就退出循环
                        if (prize.Count == prize.Current)
                            break;
                    }
                    else
                        Thread.Sleep(10);
                }

                labelMain.TextAlign = ContentAlignment.TopLeft;
                labelMain.Text = str;
                labelMain.Location = new Point(Convert.ToInt32(_screenWidth/2 - labelMain.Width/2),
                    Convert.ToInt32(_screenHeight/2) - labelMain.Height/2);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private String FillName(String data)
        {
            if (data.Length < 3)
                return data + "　";
            return data;
        }

        private String FillNumber(String data)
        {
            if (data.Length < 5)
                return data + new String(' ', 5 - data.Length);
            return data;
        }

        //自动刷新
        private void Timer_Tick(object sender, EventArgs e)
        {
            ReFresh();
        }

        private void ReFresh()
        {
            try
            {
                //this.SuspendLayout();
                if (_renyuanArr.Count < 1)
                {
                    timer1.Stop();
                    return;
                }
                var people = (People) _renyuanArr[_cIndex];
                labelMain.Text = string.Format("{0} {1} {2}"
                    , people.Number.Substring(0, 1) + "XX"
                    , people.Name.Substring(0, 1) + "XX"
                    , "XX部");
                labelMain.Location = new Point(Convert.ToInt32(_screenWidth/2 - labelMain.Width/2),
                    Convert.ToInt32(_screenHeight/2) - labelMain.Height/2);
                labelMain.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btn_Click(object sender, EventArgs e)
        {
            try
            {
                if (!_stop)
                    return;

                labelMain.Text = "";
                labelMain.AutoSize = true;

                var pic = (PictureBox) sender;
                _type = Convert.ToInt32(pic.Name.Substring(3, pic.Name.Length - 3));
                var prize = (Prize) _jiangxiangArr[_type];
                labelMsg.Text = prize.TypeName + " 总数量 " + prize.Count + " 已抽取 " +
                                prize.Current + " 每次抽取 " + prize.Yicichouqu;
                labelMsg.Location = new Point(Convert.ToInt32(_screenWidth/2) - Convert.ToInt32(labelMsg.Width/2),
                    Convert.ToInt32(_screenHeight/2) + 320);

                if (_f != null)
                    _f.Close();

                _f = new Form2(prize.JpIcon);
                _f.Show();
                Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            _stop = true;
            Close();
        }

        //动态增加抽奖人数，临时调整中奖人数
        private void AddUser(int i)
        {
            try
            {
                var prize = (Prize) _jiangxiangArr[_type];
                prize.Count=prize.Count + i;
                _jiangxiangArr[_type] = prize;
                _prizeWs.Cells[_type + 2, 2] = prize.Count;
                _wb.SaveAs(_dataName, Version, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                prize = (Prize) _jiangxiangArr[_type];
                labelMsg.Text = string.Format("{0}总数量 {1} 已抽取 {2}", prize.TypeName, prize.Count,
                    prize.Current);
                    
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btn_AddUser1(object sender, EventArgs e)
        {
            AddUser(1);
        }

        private void btn_AddUser2(object sender, EventArgs e)
        {
            AddUser(2);
        }

        private void btn_AddUser5(object sender, EventArgs e)
        {
            AddUser(5);
        }

        private void btn_AddUse10(object sender, EventArgs e)
        {
            AddUser(10);
        }

        private void btn_ShowWinner(object sender, EventArgs e)
        {
            try
            {
                if (_f != null)
                    _f.Close();

                ShowCon(false);
                if (_bgName)
                {
                    BackgroundImage = Image.FromFile(_path + "bg2.jpg");
                    BackgroundImageLayout = ImageLayout.Stretch;
                    _bgName = false;
                }

                var jiangxiang = (Prize) _jiangxiangArr[_type];

                String str = "";
                var temp = (Worksheet) _wb.Worksheets[3];
                int j = 0;

                for (int i = 2; i < temp.UsedRange.Rows.Count + 1; i++)
                {
                    Range range2 = temp.Range[temp.Cells[i, 2], temp.Cells[i, 2]];
                    Range range3 = temp.Range[temp.Cells[i, 3], temp.Cells[i, 3]];
                    Range range4 = temp.Range[temp.Cells[i, 4], temp.Cells[i, 4]];
                    if (range2.Value2.ToString() == jiangxiang.TypeName)
                    {
                        if (range3.Value2.ToString() != "" && range4.Value2.ToString() != "")
                        {
                            j++;
                            if (_type == 0)
                            {
                                str = str + "\n" + FillNumber(range3.Value2.ToString()) + " " +
                                      FillName(range4.Value2.ToString());
                            }
                            else if (_type == 1 || _type == 2)
                            {
                                if ((j - 1)%2 == 0)
                                    str = str + "\n" + FillNumber(range3.Value2.ToString()) + " " +
                                          FillName(range4.Value2.ToString());
                                else
                                    str = str + "  " + FillNumber(range3.Value2.ToString()) + " " +
                                          FillName(range4.Value2.ToString());
                            }
                            else
                            {
                                if ((j - 1)%3 == 0)
                                    str = str + "\n" + FillNumber(range3.Value2.ToString()) + " " +
                                          FillName(range4.Value2.ToString());
                                else
                                    str = str + "  " + FillNumber(range3.Value2.ToString()) + " " +
                                          FillName(range4.Value2.ToString());
                            }
                        }
                    }
                }
                if (_type == 0)
                {
                    labelMain.Font = new Font("LiSu", 60F, FontStyle.Bold, GraphicsUnit.Point, 134);
                }
                else if (_type == 1 || _type == 2)
                {
                    labelMain.Font = new Font("LiSu", 46F, FontStyle.Bold, GraphicsUnit.Point, 134);
                }
                else
                {
                    labelMain.Font = new Font("LiSu", 24F, FontStyle.Bold, GraphicsUnit.Point, 134);
                }
                labelMain.TextAlign = ContentAlignment.TopLeft;
                labelMain.Text = str;
                labelMain.Location = new Point(Convert.ToInt32(_screenWidth/2 - labelMain.Width/2),
                    Convert.ToInt32(_screenHeight/2) - labelMain.Height/2);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }

    //人员名单
    public class People
    {
        public string Department { get; set; }
        public string Number { get; set; }
        public int OP { get; set; }
        public string Name { get; set; }

    }

    //奖项设置
    public class Prize
    {
        public int Count { get; set; }
        public int Current { get; set; } //总数量、已抽取、一次抽取
        public string JpIcon { get; set; }//奖项名称、图片地址
        public string Picpath { get; set; } //奖项名称、图片地址
        public string TypeName { get; set; } //奖项名称、图片地址
        public int Yicichouqu { get; set; } //总数量、已抽取、一次抽取

    }
}
using CefSharp;
using CefSharp.WinForms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace YB_RTC_Grab
{
    public partial class Main_Form : Form
    {
        private ChromiumWebBrowser chromeBrowser;
        private bool m_aeroEnabled;
        private bool __isClose;
        private int __secho;
        private int __display_length = 20;
        private int __result_count_json;
        private int __total_page;
        private int __i = 0;
        private int __index = 1;
        private JObject __jo;
        private JToken __conn_id;
        private bool __isLogin = false;
        private bool __isStart = false;
        private bool __isBreak = false;
        private string __player_last_username = "";
        private string __playerlist_cn;
        private string __playerlist_ea;
        private string __playerlist_qq;
        private string __playerlist_wc;
        private string __player_id;
        private string __player_ldd;
        private string __start_time;
        private string __end_time;
        private string __get_time;
        private string __url = "";
        private bool __isInsert = false;
        private string __brand_code = "YB";
        private string __brand_color = "#EC6506";
        private int __count = 0;
        Form __mainFormHandler;

        // Deposit
        private int __index_deposit = 1;
        private int __count_deposit = 0;
        private bool __isInsert_deposit = false;
        private bool __isInsertDetect_deposit = false;
        private JObject __jo_deposit;
        private bool __detectInsert_deposit = false;
        private int __send = 0;

        // Drag Header to Move
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        // ----- Drag Header to Move

        // Form Shadow
        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn
        (
            int nLeftRect,
            int nTopRect,
            int nRightRect,
            int nBottomRect,
            int nWidthEllipse,
            int nHeightEllipse
        );
        [DllImport("dwmapi.dll")]
        public static extern int DwmExtendFrameIntoClientArea(IntPtr hWnd, ref MARGINS pMarInset);
        [DllImport("dwmapi.dll")]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
        [DllImport("dwmapi.dll")]
        public static extern int DwmIsCompositionEnabled(ref int pfEnabled);
        private const int CS_DROPSHADOW = 0x00020000;
        private const int WM_NCPAINT = 0x0085;
        private const int WM_ACTIVATEAPP = 0x001C;
        private const int WM_NCHITTEST = 0x84;
        private const int HTCLIENT = 0x1;
        private const int HTCAPTION = 0x2;
        private const int WS_MINIMIZEBOX = 0x20000;
        private const int CS_DBLCLKS = 0x8;
        public struct MARGINS
        {
            public int leftWidth;
            public int rightWidth;
            public int topHeight;
            public int bottomHeight;
        }
        protected override CreateParams CreateParams
        {
            get
            {
                m_aeroEnabled = CheckAeroEnabled();

                CreateParams cp = base.CreateParams;
                if (!m_aeroEnabled)
                    cp.ClassStyle |= CS_DROPSHADOW;

                cp.Style |= WS_MINIMIZEBOX;
                cp.ClassStyle |= CS_DBLCLKS;
                return cp;
            }
        }
        private bool CheckAeroEnabled()
        {
            if (Environment.OSVersion.Version.Major >= 6)
            {
                int enabled = 0;
                DwmIsCompositionEnabled(ref enabled);
                return (enabled == 1) ? true : false;
            }
            return false;
        }
        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case WM_NCPAINT:
                    if (m_aeroEnabled)
                    {
                        var v = 2;
                        DwmSetWindowAttribute(Handle, 2, ref v, 4);
                        MARGINS margins = new MARGINS()
                        {
                            bottomHeight = 1,
                            leftWidth = 0,
                            rightWidth = 0,
                            topHeight = 0
                        };
                        DwmExtendFrameIntoClientArea(Handle, ref margins);

                    }
                    break;
                default:
                    break;
            }
            base.WndProc(ref m);

            if (m.Msg == WM_NCHITTEST && (int)m.Result == HTCLIENT)
                m.Result = (IntPtr)HTCAPTION;
        }
        // ----- Form Shadow

        public Main_Form()
        {
            InitializeComponent();
            
            timer_landing.Start();
        }

        // Drag to Move
        private void panel_header_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void label_title_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void panel2_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void pictureBox_loader_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void label_brand_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void label_player_last_registered_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void panel_landing_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void pictureBox_landing_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void pictureBox_header_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        // ----- Drag to Move

        // Click Close
        private void pictureBox_close_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Exit the program?", "YB RTC Grab", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                __isClose = true;
                Environment.Exit(0);
            }
        }

        // Click Minimize
        private void pictureBox_minimize_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        // Form Closing
        private void Main_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!__isClose)
            {
                DialogResult dr = MessageBox.Show("Exit the program?", "YB RTC Grab", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr == DialogResult.No)
                {
                    e.Cancel = true;
                }
                else
                {
                    Environment.Exit(0);
                }
            }

            Environment.Exit(0);
        }
        
        [DllImport("winmm.dll")]
        public static extern int waveOutGetVolume(IntPtr h, out uint dwVolume);

        // Mute Sounds
        [DllImport("winmm.dll")]
        public static extern int waveOutSetVolume(IntPtr h, uint dwVolume);

        // Form Load
        private void Main_Form_Load(object sender, EventArgs e)
        {
            int NewVolume = ((ushort.MaxValue / 10) * 100);
            uint NewVolumeAllChannels = (((uint)NewVolume & 0x0000ffff) | ((uint)NewVolume << 16));
            waveOutSetVolume(IntPtr.Zero, NewVolumeAllChannels);

            InitializeChromium();
        }

        // CefSharp Initialize
        private void InitializeChromium()
        {
            CefSettings settings = new CefSettings();

            settings.CachePath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + @"\CEF";
            Cef.Initialize(settings);
            chromeBrowser = new ChromiumWebBrowser("http://103.4.104.8/page/manager/login.jsp");
            panel_cefsharp.Controls.Add(chromeBrowser);
            chromeBrowser.AddressChanged += ChromiumBrowserAddressChanged;
        }

        static int LineNumber([System.Runtime.CompilerServices.CallerLineNumber] int lineNumber = 0)
        {
            return lineNumber;
        }
        
        // CefSharp Address Changed
        private void ChromiumBrowserAddressChanged(object sender, AddressChangedEventArgs e)
        {
            __url = e.Address.ToString();
            if (e.Address.ToString().Equals("http://103.4.104.8/page/manager/login.jsp"))
            {
                if (__isStart)
                {
                    Invoke(new Action(() =>
                    {
                        label_brand.Visible = false;
                        pictureBox_loader.Visible = false;
                        label_player_last_registered.Visible = false;
                        label_page_count.Visible = false;
                        label_currentrecord.Visible = false;
                        __mainFormHandler = Application.OpenForms[0];
                        __mainFormHandler.Size = new Size(466, 468);
                    
                        string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                        //SendITSupport("The application have been logout, please re-login again.");
                        SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>The application have been logout, please re-login again.</b></body></html>");
                        __send = 0;
                    }));
                }

                __isLogin = false;
                __isStart = false;

                Invoke(new Action(() =>
                {
                    chromeBrowser.FrameLoadEnd += (sender_, args) =>
                    {
                        if (args.Frame.IsMain)
                        {
                            Invoke(new Action(() =>
                            {
                                if (!__isStart)
                                {
                                    timer.Stop();
                                    args.Frame.ExecuteJavaScriptAsync("document.getElementById('username').value = 'ybrtcgrab';");
                                    args.Frame.ExecuteJavaScriptAsync("document.getElementById('password').value = 'rg123888';");
                                    //args.Frame.ExecuteJavaScriptAsync("document.getElementById('username').value = 'testrain';");
                                    //args.Frame.ExecuteJavaScriptAsync("document.getElementById('password').value = 'rain12345';");
                                    __isStart = false;
                                    panel_cefsharp.Visible = true;
                                    label_player_last_registered.Text = "-";
                                    label_brand.Visible = false;
                                    pictureBox_loader.Visible = false;
                                    label_player_last_registered.Visible = false;
                                }
                            }));
                        }
                    };
                }));
            }
            
            if (e.Address.ToString().Equals("http://103.4.104.8/page/manager/member/search.jsp") || e.Address.ToString().Equals("http://103.4.104.8/page/manager/dashboard.jsp"))
            {
                Invoke(new Action(async () =>
                {
                    label_brand.Visible = true;
                    pictureBox_loader.Visible = true;
                    label_player_last_registered.Visible = true;
                    label_page_count.Visible = true;
                    label_currentrecord.Visible = true;
                    __mainFormHandler = Application.OpenForms[0];
                    __mainFormHandler.Size = new Size(466, 168);
                    
                    __isLogin = true;
                    
                    if (!__isStart)
                    {
                        __isStart = true;
                        panel_cefsharp.Visible = false;
                        label_brand.Visible = true;
                        pictureBox_loader.Visible = true;
                        label_player_last_registered.Visible = true;
                        ___PlayerLastRegistered();
                        await ___GetPlayerListsRequest();
                        ___GetPlayerListsRequest_Deposit();
                    }
                }));
            }
        }

        private async void timer_TickAsync(object sender, EventArgs e)
        {
            timer.Stop();
            await ___GetPlayerListsRequest();

            if (__isInsert_deposit)
            {
                __isInsert_deposit = false;
                ___GetPlayerListsRequest_Deposit();
            }
        }

        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern IntPtr FindWindowByCaption(IntPtr ZeroOnly, string lpWindowName);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);

        const UInt32 WM_CLOSE = 0x0010;

        void ___CloseMessageBox()
        {
            IntPtr windowPtr = FindWindowByCaption(IntPtr.Zero, "JavaScript Alert - http://103.4.104.8");

            if (windowPtr == IntPtr.Zero)
            {
                return;
            }

            SendMessage(windowPtr, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        }

        private void timer_close_message_box_Tick(object sender, EventArgs e)
        {
            ___CloseMessageBox();
        }

        // ----- Functions
        private async Task ___GetPlayerListsRequest()
        {
            try
            {
                var cookieManager = Cef.GetGlobalCookieManager();
                var visitor = new CookieCollector();
                cookieManager.VisitUrlCookies(__url, true, visitor);
                var cookies = await visitor.Task;
                var cookie = CookieCollector.GetCookieHeader(cookies);
                WebClient wc = new WebClient();
                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

                byte[] result = await wc.DownloadDataTaskAsync("http://103.4.104.8/manager/member/searchMember?userId=&userName=&email=&lastDepositSince=&lastBetTimeSince=&noLoginSince=&loginIp=&vipLevel=-1&phoneNumber=&registeredDateStart=&registeredDateEnd=&birthOfDateStart=&birthOfDateEnd=&searchType=1&affiliateCode=All&pageNumber=1&pageSize=5000&sortCondition=1&sortName=sign_up_time&sortOrder=1&searchText=");
                string responsebody = Encoding.UTF8.GetString(result);
                var deserializeObject = JsonConvert.DeserializeObject(responsebody);
                __jo = JObject.Parse(deserializeObject.ToString());
                await ___PlayerListAsync();
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    await ___GetPlayerListsRequest();
                }
            }
        }

        private async Task ___PlayerListAsync()
        {
            List<string> player_info = new List<string>();
            
            for (int i = 0; i < 5000; i++)
            {
                JToken username = __jo.SelectToken("$.aaData[" + i + "].userId").ToString();
                
                if (username.ToString() != Properties.Settings.Default.______last_registered_player)
                {
                    if (i == 0 && __index == 1)
                    {
                        __player_last_username = username.ToString();
                    }

                    await ___PlayerListContactNumberEmailAsync(username.ToString());

                    JToken date_time_register = __jo.SelectToken("$.aaData[" + i + "].createTime").ToString();
                    JToken name = __jo.SelectToken("$.aaData[" + i + "].userName").ToString();
                    JToken email = __jo.SelectToken("$.aaData[" + i + "].email").ToString();
                    JToken cn = __jo.SelectToken("$.aaData[" + i + "].phoneNumber").ToString();
                    JToken ldd = __jo.SelectToken("$.aaData[" + i + "].lastDepositTime").ToString();

                    if (!String.IsNullOrEmpty(date_time_register.ToString()) && !String.IsNullOrEmpty(ldd.ToString()))
                    {
                        DateTime date_time_register_replace = new DateTime(1970, 1, 1, 0, 0, 0, 0).AddSeconds(Math.Round(Convert.ToDouble(date_time_register.ToString()) / 1000d)).ToLocalTime();
                        DateTime ldd_replace = new DateTime(1970, 1, 1, 0, 0, 0, 0).AddSeconds(Math.Round(Convert.ToDouble(ldd.ToString()) / 1000d)).ToLocalTime();

                        player_info.Add(username + "*|*" + name + "*|*" + date_time_register_replace.ToString("yyyy-MM-dd HH:mm:ss") + "*|*" + ldd_replace.ToString("yyyy-MM-dd HH:mm:ss") + "*|*" + cn + "*|*" + email + "*|*" + __playerlist_qq + "*|*" + __playerlist_wc);
                    }
                    else
                    {
                        if (date_time_register.ToString() != "" && ldd.ToString() == "")
                        {
                            DateTime date_time_register_replace = new DateTime(1970, 1, 1, 0, 0, 0, 0).AddSeconds(Math.Round(Convert.ToDouble(date_time_register.ToString()) / 1000d)).ToLocalTime();

                            player_info.Add(username + "*|*" + name + "*|*" + date_time_register_replace.ToString("yyyy-MM-dd HH:mm:ss") + "*|*" + "" + "*|*" + cn + "*|*" + email + "*|*" + __playerlist_qq + "*|*" + __playerlist_wc);
                        }
                        else
                        {
                            DateTime ldd_replace = new DateTime(1970, 1, 1, 0, 0, 0, 0).AddSeconds(Math.Round(Convert.ToDouble(ldd.ToString()) / 1000d)).ToLocalTime();

                            player_info.Add(username + "*|*" + name + "*|*" + "" + "*|*" + ldd_replace.ToString("yyyy-MM-dd HH:mm:ss") + "*|*" + cn + "*|*" + email + "*|*" + __playerlist_qq + "*|*" + __playerlist_wc);
                        }
                    }

                    __playerlist_qq = "";
                    __playerlist_wc = "";
                }
                else
                {
                    if (player_info.Count != 0)
                    {
                        player_info.Reverse();
                        string __player_info_get = String.Join(",", player_info);
                        string[] values = __player_info_get.Split(',');
                        foreach (string value in values)
                        {
                            string[] values_inner = value.Split(new string[] { "*|*" }, StringSplitOptions.None);
                            int count = 0;
                            string _username = "";
                            string _name = "";
                            string _date_register = "";
                            string _date_deposit = "";
                            string _cn = "";
                            string _email = "";
                            string _agent = "";
                            string _qq = "";
                            string _wc = "";

                            foreach (string value_inner in values_inner)
                            {
                                count++;

                                // Username
                                if (count == 1)
                                {
                                    _username = value_inner;
                                }
                                // Name
                                else if (count == 2)
                                {
                                    _name = value_inner;
                                }
                                // Register Date
                                else if (count == 3)
                                {
                                    _date_register = value_inner;
                                }
                                // Last Deposit Date
                                else if (count == 4)
                                {
                                    _date_deposit = value_inner;
                                }
                                // Contact Number
                                else if (count == 5)
                                {
                                    _cn = value_inner;
                                }
                                // Email
                                else if (count == 6)
                                {
                                    _email = value_inner;
                                }
                                // QQ
                                else if (count == 7)
                                {
                                    _qq = value_inner;
                                }
                                // WeChat
                                else if (count == 8)
                                {
                                    _wc = value_inner;
                                }
                            }

                            // ----- Insert Data
                            //using (StreamWriter file = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\rtcgrab_yb.txt", true, Encoding.UTF8))
                            //{
                            //    file.WriteLine(_username + "*|*" + _name + "*|*" + _date_register + "*|*" + _date_deposit + "*|*" + _cn + "*|*" + _email + "*|*" + _agent + "*|*" + _qq + "*|*" + _wc);
                            //}
                            using (StreamWriter file = new StreamWriter(Path.GetTempPath() + @"\rtcgrab_yb.txt", true, Encoding.UTF8))
                            {
                                file.WriteLine(_username + "*|*" + _name + "*|*" + _date_register + "*|*" + _date_deposit + "*|*" + _cn + "*|*" + _email + "*|*" + _agent + "*|*" + _qq + "*|*" + _wc);
                            }

                            Thread t = new Thread(delegate () { ___InsertData(_username, _name, _date_register, _date_deposit, _cn, _email, _agent, _qq, _wc, __brand_code); });
                            t.Start();

                            __count = 0;
                        }
                    }

                    if (!String.IsNullOrEmpty(__player_last_username.Trim()))
                    {
                        ___SavePlayerLastRegistered(__player_last_username);
                        label_player_last_registered.Text = "Last Registered: " + Properties.Settings.Default.______last_registered_player;
                    }
                    
                    player_info.Clear();
                    timer.Start();
                    break;

                }
            }
        }

        private async Task ___PlayerListContactNumberEmailAsync(string username)
        {
            try
            {
                var cookieManager = Cef.GetGlobalCookieManager();
                var visitor = new CookieCollector();
                cookieManager.VisitUrlCookies(__url, true, visitor);
                var cookies = await visitor.Task;
                var cookie = CookieCollector.GetCookieHeader(cookies);
                WebClient wc = new WebClient();
                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

                byte[] result = await wc.DownloadDataTaskAsync("http://103.4.104.8/manager/member/getProfileOverview?userId=" + username);
                string responsebody = Encoding.UTF8.GetString(result);
                var deserializeObject = JsonConvert.DeserializeObject(responsebody);
                JObject jo_deposit = JObject.Parse(deserializeObject.ToString());
                JToken _qq = jo_deposit.SelectToken("$.qqId").ToString();
                JToken _wc = jo_deposit.SelectToken("$.wechatId").ToString();

                __playerlist_qq = _qq.ToString();
                __playerlist_wc = _wc.ToString();
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    await ___PlayerListContactNumberEmailAsync(username);
                }
            }
        }

        private void ___InsertData(string username, string name, string date_register, string date_deposit, string contact, string email, string agent, string qq, string wc, string brand_code)
        {
            try
            {
                string password = username + date_register + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["username"] = username,
                        ["name"] = name,
                        ["date_register"] = date_register,
                        ["date_deposit"] = date_deposit,
                        ["contact"] = contact,
                        ["email"] = email,
                        ["agent"] = agent,
                        ["qq"] = qq,
                        ["wc"] = wc,
                        ["brand_code"] = brand_code,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://zeus.ssimakati.com:8080/API/sendRTC", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);
                }
            }
            catch (Exception err)
            {
                __count++;
                if (__count == 5)
                {
                    string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                    SendITSupport("There's a problem to the server, please re-open the application.");
                    SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                    __send = 0;

                    __isClose = false;
                    Environment.Exit(0);
                }
                else
                {
                    ____InsertData2(username, name, date_register, date_deposit, contact, email, agent, qq, wc, brand_code);
                }
            }
        }

        private void ____InsertData2(string username, string name, string date_register, string date_deposit, string contact, string email, string agent, string qq, string wc, string brand_code)
        {
            try
            {
                string password = username + date_register + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["username"] = username,
                        ["name"] = name,
                        ["date_register"] = date_register,
                        ["date_deposit"] = date_deposit,
                        ["contact"] = contact,
                        ["email"] = email,
                        ["agent"] = agent,
                        ["qq"] = qq,
                        ["wc"] = wc,
                        ["brand_code"] = brand_code,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://zeus2.ssitex.com:8080/API/sendRTC", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __count++;
                    if (__count == 5)
                    {
                        string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                        __send = 0;

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___InsertData(username, name, date_register, date_deposit, contact, email, agent, qq, wc, brand_code);
                    }
                }
            }
        }

        public static long ___ConvertToTS(DateTime datetime)
        {
            DateTime sTime = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);

            return (long)(datetime - sTime).TotalMilliseconds;
        }

        private void ___PlayerLastRegistered()
        {
            if (Properties.Settings.Default.______last_registered_player == "" && Properties.Settings.Default.______last_registered_player_deposit == "")
            {
                ___GetLastRegisteredPlayer();
            }

            label_player_last_registered.Text = "Last Registered: " + Properties.Settings.Default.______last_registered_player;
        }

        private void ___SavePlayerLastRegistered(string username)
        {
            Properties.Settings.Default.______last_registered_player = username;
            Properties.Settings.Default.Save();
        }

        private void timer_landing_Tick(object sender, EventArgs e)
        {
            panel_landing.Visible = false;
            timer_landing.Stop();
        }

        // Deposit
        private async void ___GetPlayerListsRequest_Deposit()
        {
            try
            {
                var cookieManager = Cef.GetGlobalCookieManager();
                var visitor = new CookieCollector();
                cookieManager.VisitUrlCookies(__url, true, visitor);
                var cookies = await visitor.Task;
                var cookie = CookieCollector.GetCookieHeader(cookies);
                WebClient wc = new WebClient();
                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

                byte[] result = await wc.DownloadDataTaskAsync("http://103.4.104.8/manager/member/searchMember?userId=&userName=&email=&lastDepositSince=&lastBetTimeSince=&noLoginSince=&loginIp=&vipLevel=-1&phoneNumber=&registeredDateStart=&registeredDateEnd=&birthOfDateStart=&birthOfDateEnd=&searchType=1&affiliateCode=All&pageNumber=1&pageSize=5000&sortCondition=1&sortName=sign_up_time&sortOrder=1&searchText=");
                string responsebody = Encoding.UTF8.GetString(result);
                var deserializeObject = JsonConvert.DeserializeObject(responsebody);
                __jo_deposit = JObject.Parse(deserializeObject.ToString());
                ___PlayerListAsync_Deposit();
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    ___GetPlayerListsRequest_Deposit();
                }
            }
        }

        private void ___PlayerListAsync_Deposit()
        {
            List<string> player_info = new List<string>();
            string path = @"\rtcgrab_yb_deposit.txt";
            for (int i = 0; i < 5000; i++)
            {
                if (!File.Exists(Path.GetTempPath() + path))
                {
                    using (StreamWriter file = new StreamWriter(Path.GetTempPath() + path, true, Encoding.UTF8))
                    {
                        file.WriteLine("test123*|*");
                        file.Close();
                    }
                }

                JToken username = __jo_deposit.SelectToken("$.aaData[" + i + "].userId").ToString();

                if (username.ToString() == Properties.Settings.Default.______last_registered_player)
                {
                    __detectInsert_deposit = true;
                }

                bool isInsert = false;

                if (__detectInsert_deposit)
                {
                    using (StreamReader sr = File.OpenText(Path.GetTempPath() + path))
                    {
                        string s = String.Empty;
                        while ((s = sr.ReadLine()) != null)
                        {
                            Application.DoEvents();

                            if (s == username.ToString())
                            {
                                isInsert = true;
                                break;
                            }
                            else
                            {
                                isInsert = false;
                            }
                        }
                        sr.Close();
                    }
                }

                if (username.ToString() != Properties.Settings.Default.______last_registered_player_deposit)
                {
                    if (__detectInsert_deposit)
                    {
                        JToken ldd = __jo_deposit.SelectToken("$.aaData[" + i + "].lastDepositTime").ToString();
                        if (!String.IsNullOrEmpty(ldd.ToString()))
                        {
                            DateTime ldd_replace = new DateTime(1970, 1, 1, 0, 0, 0, 0).AddSeconds(Math.Round(Convert.ToDouble(ldd.ToString()) / 1000d)).ToLocalTime();

                            if (!isInsert)
                            {
                                player_info.Add(username + "*|*" + ldd_replace.ToString("yyyy-MM-dd HH:mm:ss"));

                                using (StreamWriter file = new StreamWriter(Path.GetTempPath() + path, true, Encoding.UTF8))
                                {
                                    file.WriteLine(username);
                                    file.Close();
                                }
                            }
                        }
                    }
                }
                else
                {
                    if (player_info.Count != 0)
                    {
                        player_info.Reverse();
                        string __player_info_deposit_get = String.Join(",", player_info);
                        string[] values = __player_info_deposit_get.Split(',');
                        foreach (string value in values)
                        {
                            string[] values_inner = value.Split(new string[] { "*|*" }, StringSplitOptions.None);
                            int count = 0;
                            string _username = "";
                            string _date_deposit = "";

                            foreach (string value_inner in values_inner)
                            {
                                count++;

                                // Username
                                if (count == 1)
                                {
                                    _username = value_inner;
                                }
                                // Last Deposit Date
                                else if (count == 2)
                                {
                                    _date_deposit = value_inner;
                                }
                            }

                            Thread t = new Thread(delegate () { ___InsertData_Deposit(_username, _date_deposit, __brand_code); });
                            t.Start();

                            __count_deposit = 0;
                        }
                    }

                    player_info.Clear();
                    __detectInsert_deposit = false;

                    break;
                }
            }

            ___DepositLastRegistered();
            __isInsert_deposit = true;
        }

        private void ___InsertData_Deposit(string username, string last_deposit_date, string brand)
        {
            try
            {
                string password = username + last_deposit_date + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["username"] = username,
                        ["date_deposit"] = last_deposit_date,
                        ["brand_code"] = brand,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://zeus.ssimakati.com:8080/API/sendRTCdep", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __count_deposit++;
                    if (__count_deposit == 5)
                    {
                        string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                        __send = 0;

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___InsertData2_Deposit(username, last_deposit_date, brand);
                    }
                }
            }
        }

        private void ___InsertData2_Deposit(string username, string last_deposit_date, string brand)
        {
            try
            {
                string password = username + last_deposit_date + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["username"] = username,
                        ["date_deposit"] = last_deposit_date,
                        ["brand_code"] = brand,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://zeus2.ssitex.com:8080/API/sendRTCdep", "POST", data);
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __count_deposit++;
                    if (__count_deposit == 5)
                    {
                        string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                        __send = 0;

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___InsertData_Deposit(username, last_deposit_date, brand);
                    }
                }
            }
        }

        private void ___DepositLastRegistered()
        {            
            string path = Path.GetTempPath() + @"\rtcgrab_yb_deposit.txt";
            if (label_player_last_registered.Text != "-" && label_player_last_registered.Text.Trim() != "")
            {
                if (Properties.Settings.Default.______detect_deposit == "")
                {
                    DateTime today = DateTime.Now;
                    DateTime date = today.AddDays(1);
                    Properties.Settings.Default.______detect_deposit = date.ToString("yyyy-MM-dd 23");
                    Properties.Settings.Default.Save();
                }
                else
                {
                    DateTime today = DateTime.Now;
                    if (Properties.Settings.Default.______detect_deposit == today.ToString("yyyy-MM-dd HH"))
                    {
                        Properties.Settings.Default.______detect_deposit = "";
                        Properties.Settings.Default.______last_registered_player_deposit = label_player_last_registered.Text.Replace("Last Registered: ", "");
                        Properties.Settings.Default.Save();

                        if (File.Exists(path))
                        {
                            File.Delete(path);
                        }
                    }
                    else
                    {
                        string start_datetime = today.ToString("yyyy-MM-dd HH");
                        DateTime start = DateTime.ParseExact(start_datetime, "yyyy-MM-dd HH", CultureInfo.InvariantCulture);

                        string end_datetime = Properties.Settings.Default.______detect_deposit;
                        DateTime end = DateTime.ParseExact(end_datetime, "yyyy-MM-dd HH", CultureInfo.InvariantCulture);

                        if (start > end)
                        {
                            Properties.Settings.Default.______detect_deposit = "";
                            Properties.Settings.Default.______last_registered_player_deposit = label_player_last_registered.Text.Replace("Last Registered: ", "");
                            Properties.Settings.Default.Save();

                            if (File.Exists(path))
                            {
                                File.Delete(path);
                            }
                        }
                    }
                }
            }
        }
                
        private void SendEmail(string get_message)
        {
            try
            {
                int port = 587;
                string host = "smtp.gmail.com";
                string username = "drake@18tech.com";
                string password = "@ccess123418tech";
                string mailFrom = "noreply@mail.com";
                string mailTo = "drake@18tech.com";
                string mailTitle = "YB RTC Grab";
                string mailMessage = get_message;

                using (SmtpClient client = new SmtpClient())
                {
                    MailAddress from = new MailAddress(mailFrom);
                    MailMessage message = new MailMessage
                    {
                        From = from
                    };
                    message.To.Add(mailTo);
                    message.Subject = mailTitle;
                    message.Body = mailMessage;
                    message.IsBodyHtml = true;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    client.Host = host;
                    client.Port = port;
                    client.EnableSsl = true;
                    client.Credentials = new NetworkCredential
                    {
                        UserName = username,
                        Password = password
                    };
                    client.Send(message);
                }
            }
            catch (Exception err)
            {
                __send++;
                if (__send <= 5)
                {
                    SendEmail(get_message);
                }
                else
                {
                    MessageBox.Show(err.ToString());
                }
            }
        }

        private void SendITSupport(string message)
        {
            try
            {
                string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                string urlString = "https://api.telegram.org/bot{0}/sendMessage?chat_id={1}&text={2}";
                string apiToken = "730388860:AAEwto3A-XT5UBTEEe3wBUQ5edxde8z508Q";
                string chatId = "@rtc_grab_it_support";
                string text = "Brand:%20-----" + __brand_code + "-----%0AIP:%20192.168.10.252%0ALocation:%20Robinsons%20Summit%20Office%0ADate%20and%20Time:%20[" + datetime + "]%0AMessage:%20" + message + "";
                urlString = String.Format(urlString, apiToken, chatId, text);
                WebRequest request = WebRequest.Create(urlString);
                Stream rs = request.GetResponse().GetResponseStream();
                StreamReader reader = new StreamReader(rs);
                string line = "";
                StringBuilder sb = new StringBuilder();
                while (line != null)
                {
                    line = reader.ReadLine();
                    if (line != null)
                        sb.Append(line);
                }
            }
            catch (Exception err)
            {
                __send++;
                if (__send <= 5)
                {
                    SendITSupport(message);
                }
                else
                {
                    MessageBox.Show(err.ToString());
                }
            }
        }

        private void ___GetLastRegisteredPlayer()
        {
            try
            {
                string password = __brand_code + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["brand_code"] = __brand_code,
                        ["token"] = token
                    };

                    var result = wb.UploadValues("http://zeus.ssimakati.com:8080/API/lastRTCrecord", "POST", data);
                    string responsebody = Encoding.UTF8.GetString(result);
                    var deserializeObject = JsonConvert.DeserializeObject(responsebody);
                    JObject jo = JObject.Parse(deserializeObject.ToString());
                    JToken plr = jo.SelectToken("$.msg");

                    Properties.Settings.Default.______last_registered_player = plr.ToString();
                    Properties.Settings.Default.______last_registered_player_deposit = plr.ToString();
                    Properties.Settings.Default.Save();
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __count++;
                    if (__count == 5)
                    {
                        string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                        __send = 0;

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___GetLastRegisteredPlayer2();
                    }
                }
            }
        }

        private void ___GetLastRegisteredPlayer2()
        {
            try
            {
                string password = __brand_code + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["brand_code"] = __brand_code,
                        ["token"] = token
                    };

                    var result = wb.UploadValues("http://zeus2.ssitex.com:8080/API/lastRTCrecord", "POST", data);
                    string responsebody = Encoding.UTF8.GetString(result);
                    var deserializeObject = JsonConvert.DeserializeObject(responsebody);
                    JObject jo = JObject.Parse(deserializeObject.ToString());
                    JToken plr = jo.SelectToken("$.msg");

                    Properties.Settings.Default.______last_registered_player = plr.ToString();
                    Properties.Settings.Default.______last_registered_player_deposit = plr.ToString();
                    Properties.Settings.Default.Save();
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __count++;
                    if (__count == 5)
                    {
                        string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                        __send = 0;

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___GetLastRegisteredPlayer();
                    }
                }
            }
        }

        private void timer_flush_memory_Tick(object sender, EventArgs e)
        {
            FlushMemory();
        }

        public static void FlushMemory()
        {
            Process prs = Process.GetCurrentProcess();
            try
            {
                prs.MinWorkingSet = (IntPtr)(300000);
            }
            catch (Exception err)
            {
                // leave blank
            }
        }

        private double __total_records_mb;
        private double __display_length_mb = 5000;
        private int __total_page_mb;
        private JObject __jo_mb;
        private int __result_count_json_mb;
        private bool __inserted_in_excel_mb = true;
        private bool __detect_mb = false;
        private int __i_mb = 0;
        private int __ii_mb = 0;
        private int __pages_count_display_mb = 0;
        private int __test_gettotal_count_record_mb;
        private int __get_ii_mb = 1;
        private int __get_ii_display_mb = 1;
        private int __pages_count_mb = 0;
        private string __shared_path = "\\\\192.168.10.22\\ssi-reporting\\";
        private string __file_name = "";
        private string __task_id = "";
        StringBuilder __csv_mb = new StringBuilder();
        StringBuilder __csv_memberrregister_custom_mb = new StringBuilder();
        
        private async Task __GetMABListsAsync()
        {
            try
            {
                var cookieManager = Cef.GetGlobalCookieManager();
                var visitor = new CookieCollector();
                cookieManager.VisitUrlCookies(__url, true, visitor);
                var cookies = await visitor.Task;
                var cookie = CookieCollector.GetCookieHeader(cookies);
                WebClient wc = new WebClient();
                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

                byte[] result = await wc.DownloadDataTaskAsync("http://103.4.104.8/manager/member/searchMember?userId=&userName=&email=&lastDepositSince=&lastBetTimeSince=&noLoginSince=&loginIp=&vipLevel=-1&phoneNumber=&registeredDateStart=&registeredDateEnd=&birthOfDateStart=&birthOfDateEnd=&searchType=1&affiliateCode=All&pageNumber=1&pageSize=80000&sortCondition=1&sortName=sign_up_time&sortOrder=1&searchText=");
                string responsebody = Encoding.UTF8.GetString(result);
                var deserializeObject = JsonConvert.DeserializeObject(responsebody);
                __jo_mb = JObject.Parse(deserializeObject.ToString());
                JToken count = __jo_mb.SelectToken("$.aaData");
                __total_records_mb = count.Count();
                ___MABPlayerList();
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    await ___GetPlayerListsRequest();
                }
            }
        }
        
        private void ___MABPlayerList()
        {
            List<string> player_info = new List<string>();

            for (int i = 0; i < __total_records_mb; i++)
            {
                Application.DoEvents();
                
                JToken username = __jo_mb.SelectToken("$.aaData[" + i + "].userId").ToString();
                JToken mab = __jo_mb.SelectToken("$.aaData[" + i + "].totalBalance").ToString();

                if (__get_ii_mb == 1)
                {
                    var header = string.Format("{0},{1},{2}", "Brand", "Username", "Main Account Balance");
                    __csv_mb.AppendLine(header);
                }

                var newLine = string.Format("{0},{1},{2}", __brand_code, "\"" + username + "\"", "\"" + mab + "\"");
                __csv_mb.AppendLine(newLine);

                label_currentrecord.Text = (__get_ii_display_mb).ToString("N0") + " of " + Convert.ToInt32(__total_records_mb).ToString("N0");
                label_currentrecord.Invalidate();
                label_currentrecord.Update();

                __get_ii_mb++;
                __get_ii_display_mb++;
            }
            
            __PlayerListInsertDoneMAB();
        }

        private void __PlayerListInsertDoneMAB()
        {
            try
            {
                string _current_datetime = DateTime.Now.ToString("yyyy-MM-ddHHmm");
                __file_name = __brand_code + "_" + _current_datetime;
                string _folder_path_result = "C:\\Projects\\zeus\\uploads\\Balance\\" + __brand_code + "_" + _current_datetime + ".txt";
                string _folder_path_result_xlsx = "C:\\Projects\\zeus\\uploads\\Balance\\" + __brand_code + "_" + _current_datetime + ".xlsx";

                if (File.Exists(_folder_path_result))
                {
                    File.Delete(_folder_path_result);
                }

                if (File.Exists(_folder_path_result_xlsx))
                {
                    File.Delete(_folder_path_result_xlsx);
                }

                __csv_mb.ToString().Reverse();
                File.WriteAllText(_folder_path_result, __csv_mb.ToString(), Encoding.UTF8);

                Excel.Application app = new Excel.Application();
                Excel.Workbook wb = app.Workbooks.Open(_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Excel.Worksheet worksheet = wb.ActiveSheet;
                worksheet.Activate();
                worksheet.Application.ActiveWindow.SplitRow = 1;
                worksheet.Application.ActiveWindow.FreezePanes = true;
                Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                firstRow.AutoFilter(1,
                                    Type.Missing,
                                    Excel.XlAutoFilterOperator.xlAnd,
                                    Type.Missing,
                                    true);
                Excel.Range usedRange = worksheet.UsedRange;
                Excel.Range rows = usedRange.Rows;
                int count = 0;
                foreach (Excel.Range row in rows)
                {
                    if (count == 0)
                    {
                        Excel.Range firstCell = row.Cells[1];

                        string firstCellValue = firstCell.Value as String;

                        if (!string.IsNullOrEmpty(firstCellValue))
                        {
                            row.Interior.Color = Color.FromArgb(236, 101, 6);
                            row.Font.Color = Color.FromArgb(255, 255, 255);
                        }

                        break;
                    }

                    count++;
                }
                int i;
                for (i = 1; i <= 3; i++)
                {
                    worksheet.Columns[i].ColumnWidth = 22;
                }
                wb.SaveAs(_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wb.Close();
                app.Quit();
                Marshal.ReleaseComObject(app);

                if (File.Exists(_folder_path_result))
                {
                    File.Delete(_folder_path_result);
                }

                __csv_mb.Clear();
                __total_records_mb = 0;
                __display_length_mb = 5000;
                __total_page_mb = 0;
                __result_count_json_mb = 0;
                __inserted_in_excel_mb = true;
                __detect_mb = false;
                __i_mb = 0;
                __ii_mb = 0;
                __pages_count_display_mb = 0;
                __test_gettotal_count_record_mb = 0;
                __get_ii_mb = 1;
                __get_ii_display_mb = 1;
                __pages_count_mb = 0;
                __csv_memberrregister_custom_mb.Clear();
                label_currentrecord.Text = "";
                label_page_count.Text = "";

                // send
                ___SetTaskStatus(__task_id, __file_name);
            }
            catch (Exception err)
            {
                __count++;
                if (__count == 5)
                {
                    string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                    SendITSupport("There's a problem to the server, please re-open the application.");
                    SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                    __send = 0;

                    __isClose = false;
                    Environment.Exit(0);
                }
                else
                {
                    MessageBox.Show(err.ToString());
                    //___GetTaskStatusAsync();
                }
            }
        }

        private void timer_mb_detect_Tick(object sender, EventArgs e)
        {
            ___GetTaskStatusAsync();
        }

        private async void ___GetTaskStatusAsync()
        {
            try
            {
                timer_mb_detect.Stop();
                string password = __brand_code + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["brand_code"] = __brand_code,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://zeus.ssimakati.com:8080/API/getBalanceTaskStatus", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);
                    var deserializeObject = JsonConvert.DeserializeObject(responseInString);
                    JObject jo_mb = JObject.Parse(deserializeObject.ToString());
                    JToken status = jo_mb.SelectToken("$.status");
                    JToken task_id = jo_mb.SelectToken("$.task_id");
                    __task_id = task_id.ToString();

                    if (status.ToString() == "1")
                    {
                        if (__url != "http://103.4.104.8/page/manager/login.jsp")
                        {
                            // start
                            timer_mb_detect.Stop();
                            __UpdateTaskStatus();
                            await __GetMABListsAsync();
                        }
                        else
                        {
                            timer_mb_detect.Start();
                        }
                    }
                    else
                    {
                        timer_mb_detect.Start();
                    }
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __count++;
                    if (__count == 5)
                    {
                        string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                        __send = 0;

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___GetTaskStatus2Async();
                    }
                }
            }
        }
        
        private async void ___GetTaskStatus2Async()
        {
            try
            {
                timer_mb_detect.Stop();
                string password = __brand_code + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["brand_code"] = __brand_code,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://zeus2.ssimakati.com:8080/API/getBalanceTaskStatus", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);
                    var deserializeObject = JsonConvert.DeserializeObject(responseInString);
                    JObject jo_mb = JObject.Parse(deserializeObject.ToString());
                    JToken status = jo_mb.SelectToken("$.status");
                    JToken task_id = jo_mb.SelectToken("$.task_id");
                    __task_id = task_id.ToString();

                    if (status.ToString() == "1")
                    {
                        if (__url != "http://103.4.104.8/page/manager/login.jsp")
                        {
                            // start
                            timer_mb_detect.Stop();
                            __UpdateTaskStatus();
                            await __GetMABListsAsync();
                        }
                        else
                        {
                            timer_mb_detect.Start();
                        }
                    }
                    else
                    {
                        timer_mb_detect.Start();
                    }
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __count++;
                    if (__count == 5)
                    {
                        string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                        __send = 0;

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___GetTaskStatusAsync();
                    }
                }
            }
        }

        private void ___SetTaskStatus(string task_id, string file_name)
        {
            try
            {
                string password = file_name + task_id + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["brand_code"] = __brand_code,
                        ["task_id"] = task_id,
                        ["filename"] = file_name,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://zeus.ssimakati.com:8080/API/setBalanceTaskStatus", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);

                    __file_name = "";
                    __task_id = "";
                    timer_mb_detect.Start();
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __count++;
                    if (__count == 5)
                    {
                        string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                        __send = 0;

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___SetTaskStatus2(task_id, file_name);
                    }
                }
            }
        }

        private void ___SetTaskStatus2(string task_id, string file_name)
        {
            try
            {
                string password = file_name + task_id + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["task_id"] = task_id,
                        ["filename"] = file_name,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://zeus2.ssimakati.com:8080/API/setBalanceTaskStatus", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);

                    __file_name = "";
                    __task_id = "";
                    timer_mb_detect.Start();
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __count++;
                    if (__count == 5)
                    {
                        string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                        __send = 0;

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___SetTaskStatus(task_id, file_name);
                    }
                }
            }
        }

        private void __UpdateTaskStatus()
        {
            try
            {
                string password = __brand_code + __task_id + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["brand_code"] = __brand_code,
                        ["task_id"] = __task_id,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://zeus.ssimakati.com:8080/API/updBalanceTaskStatus", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __count++;
                    if (__count == 5)
                    {
                        string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                        __send = 0;

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        __UpdateTaskStatus2();
                    }
                }
            }
        }

        private void __UpdateTaskStatus2()
        {
            try
            {
                string password = __brand_code + __task_id + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["brand_code"] = __brand_code,
                        ["task_id"] = __task_id,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://zeus2.ssimakati.com:8080/API/updBalanceTaskStatus", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __count++;
                    if (__count == 5)
                    {
                        string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendEmail("<html><body>Brand: <font color='" + __brand_color + "'>-----" + __brand_code + "-----</font><br/>IP: 192.168.10.252<br/>Location: Robinsons Summit Office<br/>Date and Time: [" + datetime + "]<br/>Line Number: " + LineNumber() + "<br/>Message: <b>" + err.ToString() + "</b></body></html>");
                        __send = 0;

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        __UpdateTaskStatus();
                    }
                }
            }
        }
    }
}
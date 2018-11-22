using CefSharp;
using CefSharp.WinForms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

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
        private bool __isStart = false;
        private bool __isBreak = false;
        private string __player_last_username = "";
        private string __playerlist_cn;
        private string __playerlist_ea;
        private string __player_id;
        private string __player_ldd;
        private string __start_time;
        private string __end_time;
        private string __get_time;
        private string __url = "";
        List<string> __player_info = new List<string>();
        private bool __isInsert = false;
        private string __brand_code = "YB";
        private int __count = 0;

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
        private void label_status_MouseDown(object sender, MouseEventArgs e)
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
        // ----- Drag to Move

        // Click Close
        private void pictureBox_close_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Exit the program?", "YB", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
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
                DialogResult dr = MessageBox.Show("Exit the program?", "YB", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
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

        // Form Load
        private void Main_Form_Load(object sender, EventArgs e)
        {
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

        // CefSharp Address Changed
        private void ChromiumBrowserAddressChanged(object sender, AddressChangedEventArgs e)
        {
            __url = e.Address.ToString();
            if (e.Address.ToString().Equals("http://103.4.104.8/page/manager/login.jsp"))
            {
                System.Media.SoundPlayer player = new System.Media.SoundPlayer(Properties.Resources.rtc_grab);
                bool isPlaying = false;
                if (__isStart)
                {
                    player.PlayLooping();
                    isPlaying = true;
                }
                
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
                                    args.Frame.ExecuteJavaScriptAsync("document.getElementById('username').value = 'testrain';");
                                    args.Frame.ExecuteJavaScriptAsync("document.getElementById('password').value = 'rain12345';");
                                    __isStart = false;
                                    panel_cefsharp.Visible = true;
                                    timer.Stop();
                                    label_status.Text = "-";
                                    label_player_last_registered.Text = "-";
                                    label_brand.Visible = false;
                                    pictureBox_loader.Visible = false;
                                    label_status.Visible = false;
                                    label_player_last_registered.Visible = false;

                                    if (isPlaying)
                                    {
                                        DialogResult dr = MessageBox.Show("You've been logout please login again.", "YB", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        if (dr == DialogResult.OK)
                                        {
                                            player.Stop();
                                        }
                                    }
                                }
                            }));
                        }
                    };
                }));
            }

            if (e.Address.ToString().Equals("http://103.4.104.8/page/manager/member/search.jsp"))
            {
                Invoke(new Action(() =>
                {
                    if (!__isStart)
                    {
                        __isStart = true;
                        panel_cefsharp.Visible = false;
                        label_brand.Visible = true;
                        pictureBox_loader.Visible = true;
                        label_status.Visible = true;
                        label_player_last_registered.Visible = true;
                        label_status.Text = "...";
                        ___PlayerLastRegistered();
                        ___GetPlayerListsRequestAsync(__index.ToString());
                    }
                }));
            }
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            label_status.Font = new Font("Microsoft Sans Serif", 12, FontStyle.Bold);
            label_status.Location = new Point(1, 70);
            DateTime start = DateTime.ParseExact(__start_time, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
            DateTime end = DateTime.ParseExact(__end_time, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

            TimeSpan difference = end - start;
            int hrs = difference.Hours;
            int mins = difference.Minutes;
            int secs = difference.Seconds;

            TimeSpan spinTime = new TimeSpan(hrs, mins, secs);

            TimeSpan delta = DateTime.Now - start;
            TimeSpan timeRemaining = spinTime - delta;

            if (timeRemaining.Minutes != 0)
            {
                if (timeRemaining.Seconds == 0)
                {
                    label_status.Text = timeRemaining.Minutes + ":" + timeRemaining.Seconds + "0";
                }
                else
                {
                    if (timeRemaining.Seconds.ToString().Length == 1)
                    {
                        label_status.Text = timeRemaining.Minutes + ":0" + timeRemaining.Seconds;
                    }
                    else
                    {
                        label_status.Text = timeRemaining.Minutes + ":" + timeRemaining.Seconds;
                    }
                }

                label_status.Visible = true;
            }
            else
            {
                if (label_status.Text == "1")
                {
                    timer.Stop();
                    label_status.Font = new Font("Microsoft Sans Serif", 8, FontStyle.Regular);
                    label_status.Location = new Point(0, 65);
                    label_status.Text = "...";
                    ___GetPlayerListsRequestAsync(__index.ToString());
                }
                else
                {
                    label_status.Text = timeRemaining.Seconds.ToString();
                    label_status.Visible = true;
                }
            }

            if (label_status.Text.Contains("-"))
            {
                System.Media.SoundPlayer player = new System.Media.SoundPlayer(Properties.Resources.rtc_grab);
                player.PlayLooping();

                DialogResult dr = MessageBox.Show("There's a problem to the server. Please call IT Support, thank you!", "YB", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (dr == DialogResult.OK)
                {
                    player.Stop();
                }
                
                __isClose = false;
                Environment.Exit(0);
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
        private async void ___GetPlayerListsRequestAsync(string index)
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


                byte[] result = await wc.DownloadDataTaskAsync("http://103.4.104.8/manager/member/searchMember?userId=&userName=&email=&lastDepositSince=&lastBetTimeSince=&noLoginSince=&loginIp=&vipLevel=-1&phoneNumber=&registeredDateStart=&registeredDateEnd=&birthOfDateStart=&birthOfDateEnd=&searchType=1&affiliateCode=All&pageNumber=" + __index + "&pageSize=10&sortCondition=1&sortName=sign_up_time&sortOrder=1&searchText=");
                string responsebody = Encoding.UTF8.GetString(result);
                var deserializeObject = JsonConvert.DeserializeObject(responsebody);
                __jo = JObject.Parse(deserializeObject.ToString());
                //MessageBox.Show(__jo.ToString());
                ___PlayerListAsync();
            }
            catch (Exception err)
            {
                ___GetPlayerListsRequestAsync(__index.ToString());
            }
        }

        private void ___PlayerListAsync()
        {
            for (int i = 0; i < 10; i++)
            {
                JToken username = __jo.SelectToken("$.aaData[" + i + "].userId").ToString();
                if (username.ToString() != Properties.Settings.Default.______last_registered_player)
                {
                    if (i == 0 && __index == 1)
                    {
                        __player_last_username = username.ToString();
                    }

                    JToken date_time_register = __jo.SelectToken("$.aaData[" + i + "].createTime").ToString();
                    JToken name = __jo.SelectToken("$.aaData[" + i + "].userName").ToString();
                    JToken email = __jo.SelectToken("$.aaData[" + i + "].email").ToString();
                    JToken cn = __jo.SelectToken("$.aaData[" + i + "].phoneNumber").ToString();
                    JToken ldd = __jo.SelectToken("$.aaData[" + i + "].lastDepositTime").ToString();
                    
                    if (!String.IsNullOrEmpty(date_time_register.ToString()) && !String.IsNullOrEmpty(ldd.ToString()))
                    {
                        DateTime date_time_register_replace = new DateTime(1970, 1, 1, 0, 0, 0, 0).AddSeconds(Math.Round(Convert.ToDouble(date_time_register.ToString()) / 1000d)).ToLocalTime();
                        DateTime ldd_replace = new DateTime(1970, 1, 1, 0, 0, 0, 0).AddSeconds(Math.Round(Convert.ToDouble(ldd.ToString()) / 1000d)).ToLocalTime();

                        __player_info.Add(username + "*|*" + name + "*|*" + date_time_register_replace.ToString("yyyy-MM-dd HH:mm:ss") + "*|*" + ldd_replace.ToString("yyyy-MM-dd HH:mm:ss") + "*|*" + cn + "*|*" + email);
                    }
                    else
                    {
                        if (date_time_register.ToString() != "" && ldd.ToString() == "")
                        {
                            DateTime date_time_register_replace = new DateTime(1970, 1, 1, 0, 0, 0, 0).AddSeconds(Math.Round(Convert.ToDouble(date_time_register.ToString()) / 1000d)).ToLocalTime();
                            
                            __player_info.Add(username + "*|*" + name + "*|*" + date_time_register_replace.ToString("yyyy-MM-dd HH:mm:ss") + "*|*" + "" + "*|*" + cn + "*|*" + email);
                        }
                        else
                        {
                            DateTime ldd_replace = new DateTime(1970, 1, 1, 0, 0, 0, 0).AddSeconds(Math.Round(Convert.ToDouble(ldd.ToString()) / 1000d)).ToLocalTime();
                            
                            __player_info.Add(username + "*|*" + name + "*|*" + "" + "*|*" + ldd_replace.ToString("yyyy-MM-dd HH:mm:ss") + "*|*" + cn + "*|*" + email);
                        }
                    }

                    if (i == 9)
                    {
                        __index++;
                        ___GetPlayerListsRequestAsync(__index.ToString());
                    }
                }
                else
                {
                    if (!String.IsNullOrEmpty(__player_last_username.Trim()))
                    {
                        ___SavePlayerLastRegistered(__player_last_username);
                        label_player_last_registered.Text = "Last Registered: " + Properties.Settings.Default.______last_registered_player;
                    }
                    //__player_info.Reverse();
                    //MessageBox.Show(String.Join("," + Environment.NewLine, __player_info));
                    
                    // send to api 
                    if (__player_info.Count != 0)
                    {
                        __player_info.Reverse();
                        string player_info_get = String.Join(",", __player_info);
                        string[] values = player_info_get.Split(',');
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
                                // Agent
                                else if (count == 7)
                                {
                                    _agent = value_inner;
                                }
                            }

                            // ----- Insert Data
                            using (StreamWriter file = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\test_yb.txt", true, Encoding.UTF8))
                            {
                                file.WriteLine(_username + "*|*" + _name + "*|*" + _date_register + "*|*" + _date_deposit + "*|*" + _cn + "*|*" + _email + "*|*" + _agent + "*|*" + __brand_code);
                            }
                            using (StreamWriter file = new StreamWriter(Path.GetTempPath() + @"\test_yb.txt", true, Encoding.UTF8))
                            {
                                file.WriteLine(_username + "*|*" + _name + "*|*" + _date_register + "*|*" + _date_deposit + "*|*" + _cn + "*|*" + _email + "*|*" + _agent + "*|*" + __brand_code);
                            }
                            ___InsertData(_username, _name, _date_register, _date_deposit, _cn, _email, _agent, __brand_code);
                            __count = 0;
                        }

                        __player_info.Clear();
                    }
                    __index = 1;
                    timer.Start();
                    __start_time = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                    __end_time = DateTime.Now.AddSeconds(302).ToString("dd/MM/yyyy HH:mm:ss");
                    break;

                }
            }
        }

        private void ___InsertData(string username, string name, string date_register, string date_deposit, string contact, string email, string agent, string brand_code)
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
                        ["brand_code"] = brand_code,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://zeus.ssitex.com:8080/API/sendRTC", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);
                }
            }
            catch (Exception err)
            {
                __count++;
                if (__count == 5)
                {
                    System.Media.SoundPlayer player = new System.Media.SoundPlayer(Properties.Resources.rtc_grab);
                    player.PlayLooping();

                    DialogResult dr = MessageBox.Show("There's a problem to the server. Please call IT Support, thank you!", "YB", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (dr == DialogResult.OK)
                    {
                        player.Stop();
                    }

                    __isClose = false;
                    Environment.Exit(0);
                }
                else
                {
                    ____InsertData2(username, name, date_register, date_deposit, contact, email, agent, brand_code);
                }
            }
        }

        private void ____InsertData2(string username, string name, string date_register, string date_deposit, string contact, string email, string agent, string brand_code)
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
                        ["brand_code"] = brand_code,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://zeus2.ssitex.com:8080/API/sendRTC", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);
                }
            }
            catch (Exception err)
            {
                __count++;
                if (__count == 5)
                {
                    System.Media.SoundPlayer player = new System.Media.SoundPlayer(Properties.Resources.rtc_grab);
                    player.PlayLooping();

                    DialogResult dr = MessageBox.Show("There's a problem to the server. Please call IT Support, thank you!", "YB", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (dr == DialogResult.OK)
                    {
                        player.Stop();
                    }

                    __isClose = false;
                    Environment.Exit(0);
                }
                else
                {
                    ___InsertData(username, name, date_register, date_deposit, contact, email, agent, brand_code);
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
            // handle last registered player
            if (Properties.Settings.Default.______last_registered_player == "")
            {
                //MessageBox.Show("ghghg");
                Properties.Settings.Default.______last_registered_player = "a824349234";
                Properties.Settings.Default.Save();
                // handle request
            }

            //Properties.Settings.Default.______last_registered_player = "a824349234";
            //Properties.Settings.Default.Save();

            label_player_last_registered.Text = "Last Registered: " + Properties.Settings.Default.______last_registered_player;
            // todo
        }

        private void ___SavePlayerLastRegistered(string username)
        {
            Properties.Settings.Default.______last_registered_player = username;
            Properties.Settings.Default.Save();
        }
    }
}
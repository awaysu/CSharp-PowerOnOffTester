/*
 * Created by SharpDevelop.
 * User: jason_su
 * Date: 2014/6/23
 * Time: 下午 03:10
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Net.NetworkInformation;
using System.Diagnostics;
using System.Net;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Text;
using System.IO;
using System.ComponentModel;
using System.IO.Ports;
using System.Threading;

namespace PowerOnOffTester
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class chkStopPowerOff : Form
	{
		
		[DllImport("Ws2_32.dll", EntryPoint = "inet_addr", CharSet = CharSet.Ansi)]
		private static extern Int32 inet_addr(string ipaddr);
		[DllImport("Iphlpapi.dll", EntryPoint = "SendARP", CharSet = CharSet.Ansi)]
		private static extern int SendARP(Int32 DestIP, Int32 SrcIP, ref Int64 MacAddr, ref Int32 PhyAddrLen);        
		
		//[DllImport("inpout32.dll")]
        //private static extern void Out32(ushort PortAddress, short Data);
	
		
		List<CheckBox> arrCheck = new List<CheckBox>();		
		List<TextBox> arrIPText = new List<TextBox>();
		List<TextBox> arrMacText = new List<TextBox>();
		List<TextBox> arrPassText = new List<TextBox>();
		List<TextBox> arrFailText = new List<TextBox>();
        List<TextBox> arrTelFailText = new List<TextBox>();
        List<TextBox> arrFail2Text = new List<TextBox>();
		List<PictureBox> arrStatus = new List<PictureBox>();
        List<PictureBox> arrRes = new List<PictureBox>();

        Boolean powerStatusOn = false;
        Boolean startFirst = false;
        String configFile = Application.StartupPath + "\\PowerOnOffTester.ini";
        SerialPort serialPort = new SerialPort();
        System.IO.Ports.SerialPort sp;
        bool openFlag = false;
        bool testbutton = false;

        int[] ping1Fail = new int[15];
        int[] ping2Fail = new int[15];
        int[] telnetFail = new int[15];

        bool[] ping1OK = new bool[15];

		public chkStopPowerOff()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//	

		}

        void setDateLogName()
        {
            DateTime myDate = DateTime.Now;
            string myDateString = myDate.ToString("MMddHHmmss");
            String newpath = Application.StartupPath + "\\log\\";

            if (!Directory.Exists(newpath))
            {
                Directory.CreateDirectory(newpath);
            }

            txtScriptPath.Text = newpath + "\\script.sh";
            txtLogPath.Text = newpath + "\\boot_test_" + myDateString + ".log";
            txtLogPath.SelectionStart = txtLogPath.Text.Length;
            //txtLogPath.ScrollToCaret();
        }

			
		void ChkStopPowerOffLoad(object sender, EventArgs e)
		{
			arrCheck.Add(chkBox01);
			arrCheck.Add(chkBox02);
			arrCheck.Add(chkBox03);
			arrCheck.Add(chkBox04);
			arrCheck.Add(chkBox05);
			arrCheck.Add(chkBox06);
			arrCheck.Add(chkBox07);
			arrCheck.Add(chkBox08);
			arrCheck.Add(chkBox09);
			arrCheck.Add(chkBox10);
			arrCheck.Add(chkBox11);
			arrCheck.Add(chkBox12);
			arrCheck.Add(chkBox13);
			arrCheck.Add(chkBox14);
			arrCheck.Add(chkBox15);				
			
			arrIPText.Add(txtIP01);
			arrIPText.Add(txtIP02);
			arrIPText.Add(txtIP03);
			arrIPText.Add(txtIP04);
			arrIPText.Add(txtIP05);
			arrIPText.Add(txtIP06);
			arrIPText.Add(txtIP07);
			arrIPText.Add(txtIP08);
			arrIPText.Add(txtIP09);
			arrIPText.Add(txtIP10);
			arrIPText.Add(txtIP11);
			arrIPText.Add(txtIP12);
			arrIPText.Add(txtIP13);
			arrIPText.Add(txtIP14);
			arrIPText.Add(txtIP15);			
			
			arrMacText.Add(txtMac01);
			arrMacText.Add(txtMac02);
			arrMacText.Add(txtMac03);
			arrMacText.Add(txtMac04);
			arrMacText.Add(txtMac05);
			arrMacText.Add(txtMac06);
			arrMacText.Add(txtMac07);
			arrMacText.Add(txtMac08);
			arrMacText.Add(txtMac09);
			arrMacText.Add(txtMac10);
			arrMacText.Add(txtMac11);
			arrMacText.Add(txtMac12);
			arrMacText.Add(txtMac13);
			arrMacText.Add(txtMac14);
			arrMacText.Add(txtMac15);	
			
			arrPassText.Add(txtPass01);
			arrPassText.Add(txtPass02);
			arrPassText.Add(txtPass03);
			arrPassText.Add(txtPass04);
			arrPassText.Add(txtPass05);
			arrPassText.Add(txtPass06);
			arrPassText.Add(txtPass07);
			arrPassText.Add(txtPass08);
			arrPassText.Add(txtPass09);
			arrPassText.Add(txtPass10);
			arrPassText.Add(txtPass11);
			arrPassText.Add(txtPass12);
			arrPassText.Add(txtPass13);
			arrPassText.Add(txtPass14);
			arrPassText.Add(txtPass15);				
			
			arrFailText.Add(txtFail01);
			arrFailText.Add(txtFail02);
			arrFailText.Add(txtFail03);
			arrFailText.Add(txtFail04);
			arrFailText.Add(txtFail05);
			arrFailText.Add(txtFail06);
			arrFailText.Add(txtFail07);
			arrFailText.Add(txtFail08);
			arrFailText.Add(txtFail09);
			arrFailText.Add(txtFail10);
			arrFailText.Add(txtFail11);
			arrFailText.Add(txtFail12);
			arrFailText.Add(txtFail13);
			arrFailText.Add(txtFail14);
			arrFailText.Add(txtFail15);	
			
			arrStatus.Add(picStatus01);
			arrStatus.Add(picStatus02);
			arrStatus.Add(picStatus03);
			arrStatus.Add(picStatus04);
			arrStatus.Add(picStatus05);
			arrStatus.Add(picStatus06);
			arrStatus.Add(picStatus07);
			arrStatus.Add(picStatus08);
			arrStatus.Add(picStatus09);
			arrStatus.Add(picStatus10);
			arrStatus.Add(picStatus11);
			arrStatus.Add(picStatus12);
			arrStatus.Add(picStatus13);
			arrStatus.Add(picStatus14);
			arrStatus.Add(picStatus15);

            arrRes.Add(resBox01);
            arrRes.Add(resBox02);
            arrRes.Add(resBox03);
            arrRes.Add(resBox04);
            arrRes.Add(resBox05);
            arrRes.Add(resBox06);
            arrRes.Add(resBox07);
            arrRes.Add(resBox08);
            arrRes.Add(resBox09);
            arrRes.Add(resBox10);
            arrRes.Add(resBox11);
            arrRes.Add(resBox12);
            arrRes.Add(resBox13);
            arrRes.Add(resBox14);
            arrRes.Add(resBox15);

            arrFail2Text.Add(txtFail201);
            arrFail2Text.Add(txtFail202);
            arrFail2Text.Add(txtFail203);
            arrFail2Text.Add(txtFail204);
            arrFail2Text.Add(txtFail205);
            arrFail2Text.Add(txtFail206);
            arrFail2Text.Add(txtFail207);
            arrFail2Text.Add(txtFail208);
            arrFail2Text.Add(txtFail209);
            arrFail2Text.Add(txtFail210);
            arrFail2Text.Add(txtFail211);
            arrFail2Text.Add(txtFail212);
            arrFail2Text.Add(txtFail213);
            arrFail2Text.Add(txtFail214);
            arrFail2Text.Add(txtFail215);

            arrTelFailText.Add(txtTelFail01);
            arrTelFailText.Add(txtTelFail02);
            arrTelFailText.Add(txtTelFail03);
            arrTelFailText.Add(txtTelFail04);
            arrTelFailText.Add(txtTelFail05);
            arrTelFailText.Add(txtTelFail06);
            arrTelFailText.Add(txtTelFail07);
            arrTelFailText.Add(txtTelFail08);
            arrTelFailText.Add(txtTelFail09);
            arrTelFailText.Add(txtTelFail10);
            arrTelFailText.Add(txtTelFail11);
            arrTelFailText.Add(txtTelFail12);
            arrTelFailText.Add(txtTelFail13);
            arrTelFailText.Add(txtTelFail14);
            arrTelFailText.Add(txtTelFail15);	

			resetItem();
            setDateLogName();

            loadConfigFile();

            initComPort();
		}
       
        public void initComPort()
        {

            int baud = 9600;
            System.IO.Ports.Parity parity = System.IO.Ports.Parity.None;
            System.IO.Ports.StopBits stopbits = System.IO.Ports.StopBits.One;
            int databits = 8;
            //System.IO.Ports.Handshake handshake = System.IO.Ports.Handshake.None;
            comboBox1.Items.Clear();
            openFlag = false;

            for (int ii = 1; ii <= 24; ii++)
            {
                try
                {
                    sp = new System.IO.Ports.SerialPort("COM" + ii.ToString(), baud, parity, databits, stopbits);
                    sp.Open();
                    if (sp.IsOpen)
                    {
                        comboBox1.Items.Add("COM" + ii.ToString());
                    }
                }
                catch (Exception)
                {
                }
                finally
                {
                    sp.Close();
                }
            }


            if (comboBox1.Items.Count > 0)
                comboBox1.Text = comboBox1.Items[0].ToString();
        }

		
		public void resetItem()
		{
			for (int i =0; i != arrIPText.Count;i++)
			{
				//arrIPText[i].Text = "192.168.1.1";
				arrMacText[i].Text = "00-00-00-00-00-00";
				arrPassText[i].Text = "0";
				
                arrFailText[i].Text = "0";
                arrTelFailText[i].Text = "0";
                arrFail2Text[i].Text = "0";

                ping1Fail[i] = 0;
                ping2Fail[i] = 0;
                telnetFail[i] = 0;
                ping1OK[i] = false;

				arrStatus[i].BackColor = System.Drawing.Color.LightGreen;
                arrRes[i].BackColor = System.Drawing.Color.LightGreen;
			}
		}
		
		public PingReply Ping(string url_or_IP, int timeout)	
		{
			PingReply reply = null;
			
			using (Ping ping = new Ping())
			{
				try
				{
					reply = ping.Send(url_or_IP, timeout);
					//MessageBox.Show(reply.Status.ToString());
				}
				catch (Exception ex)
				{
					//Console.WriteLine("Error ({0})", ex.InnerException.Message);
                    MessageBox.Show("Ping Function catch fail!!");
				}								
			}
			
			return reply;
		}		

	    private string GetMacAddress(string RemoteIP)
		{
			StringBuilder macAddress = new StringBuilder();
			try
			{
			    Int32 remote = inet_addr(RemoteIP);
			    Int64 macInfo = new Int64();
			    Int32 length = 6;
                try
                {
                    SendARP(remote, 0, ref macInfo, ref length);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("GetMacAddress Function catch fail!!");
                }		
			
			    string temp = Convert.ToString(macInfo, 16).PadLeft(12, '0').ToUpper();
			
			    int x = 12;
			    for (int i = 0; i < 6; i++)
			    {
			        if (i == 5)
			        {
			            macAddress.Append(temp.Substring(x - 2, 2));
			        }
			        else
			        {
			            macAddress.Append(temp.Substring(x - 2, 2) + "-");
			        }
			        x -= 2;
			    }
			}
			catch
			{
			}
			
			return macAddress.ToString();
		}
    

		private const int MODE_STOP =	0;
		private const int MODE_ONTIME =	1;
		private const int MODE_CHECK =	2;
		private const int MODE_OFFTIME = 3;
				
		int save_on_time = 0;
		int save_off_time = 0;
		int now_mode = MODE_STOP;
		int now_on_time = 0;
		int now_off_time = 0;
				

		void BtnStartClick(object sender, EventArgs e)
		{
            if (openFlag == false)
            {
                MessageBox.Show("Please first select com port!", "Warning");
                return;
            }


			DateTime dt = DateTime.Now;

			save_on_time = Convert.ToInt32(txtOnTime.Text);
			save_off_time = Convert.ToInt32(txtOffTime.Text);
			now_on_time = save_on_time;
			now_off_time = save_off_time;
			now_mode = MODE_ONTIME;
			
			timer1.Enabled = true;	
			btnStart.Enabled = false;
			btnStop.Enabled = true;
			btnPause.Enabled = true;
			
			resetItem();
								
			picStart.Visible = true;
			picStop.Visible = false;

            startFirst = true;

			powerOnFunction();		

            //if (System.IO.File.Exists(txtLogPath.Text))
            //    System.IO.File.Delete(txtLogPath.Text);
		
            setDateLogName();
            saveLogFunction(txtProject.Text + " Start Time : " + dt.ToString() + "\r\n\r\n", false);
			
		}		
		
		void BtnStopClick(object sender, EventArgs e)
		{
            String str = "";

			timer1.Enabled = false;		
			btnStart.Enabled = true;
			btnStop.Enabled = false;
			btnPause.Enabled = false;
			
			now_mode = MODE_STOP;
			
			resetOnOffTime();
			
			picStart.Visible = false;
			picStop.Visible = true;



			for (int i = 0; i != arrCheck.Count; i++)
			{
                if (arrCheck[i].Checked == true)
                {
                    str += arrIPText[i].Text + 
                        " (" + arrMacText[i].Text  + 
                        ") - pass count : " + arrPassText[i].Text + 
                        " , 1st ping fail count : " + arrFailText[i].Text +
                        " , telnet fail count : " + arrTelFailText[i].Text + 
                        " , 2nd ping fail count : " + arrFail2Text[i].Text +                        
                        "\r\n";
                }
            }

            saveLogFunction(str, true);
							
		}
		
		void BtnPauseClick(object sender, EventArgs e)
		{
			if (now_mode == MODE_STOP)
				now_mode = MODE_ONTIME;

            if (now_mode == MODE_ONTIME || now_mode == MODE_CHECK)
                powerOnFunction();
			
			if (timer1.Enabled == true)
			{
				timer1.Enabled = false;
				btnStart.Enabled = true;
				btnStop.Enabled = false;
				btnPause.Enabled = true;	

				picStart.Visible = true;
				picStop.Visible = false;				
			}
			else
			{
				timer1.Enabled = true;	
				btnStart.Enabled = false;
				btnStop.Enabled = true;
				btnPause.Enabled = true;
				
				picStart.Visible = false;
				picStop.Visible = true;				
			}
		}
		
		void BtnVerifyIPClick(object sender, EventArgs e)
		{
            testbutton = true;
            verifyIPFunction(1);
            testbutton = false;

        }
				
		void BtnPowerOnClick(object sender, EventArgs e)
		{
			powerOnFunction();
		}
		
		void BtnPowerOffClick(object sender, EventArgs e)
		{
			powerOffFunction();
		}
		
		void Timer1Tick(object sender, EventArgs e)
		{
			timer1.Enabled = false;
			
			if (now_mode == MODE_ONTIME)
			{				
				txtOnTime.BackColor =  System.Drawing.Color.Yellow;
				txtOnTime.Text = now_on_time.ToString() + "/" + save_on_time.ToString();
				now_on_time--;
								
				if (now_on_time < 0)
				{
					now_mode = MODE_CHECK;
					now_on_time = save_on_time;
					txtOnTime.Text	= save_on_time.ToString();
					txtOnTime.BackColor =  System.Drawing.SystemColors.Window;
				}								
			}
			else if (now_mode == MODE_CHECK)
			{				
				resetOnOffTime();

                if (startFirst == true)
                {
                    verifyMacFunction();
                    startFirst = false;
                }

                bool runVerifyIP = true;
                if (chkTelnet.Checked == true)
                {
                    runVerifyIP = verifyIPFunction(0);

                    if (runVerifyIP == true)
                    {
                        telnetFunction();

                        int t = Convert.ToInt32(txtScriptTime.Text) * 1000;
                        Application.DoEvents();
                        System.Threading.Thread.Sleep(t);
                    }
                }

                if (runVerifyIP == true)
                    verifyIPFunction(1);
			
				if (now_mode == MODE_CHECK)
				{
					powerOffFunction();				
					now_mode = MODE_OFFTIME;
				}
			}
			else if (now_mode == MODE_OFFTIME)
			{				
				txtOffTime.BackColor =  System.Drawing.Color.Yellow;
				txtOffTime.Text = now_off_time.ToString() + "/" + save_off_time.ToString();
				now_off_time--;
				
				if (now_off_time < 0)
				{
					now_mode = MODE_ONTIME;
					now_off_time = save_off_time;
					txtOffTime.Text = save_off_time.ToString();
					txtOffTime.BackColor =  System.Drawing.SystemColors.Window;
					
					powerOnFunction();
				}								
			}
			
			if (now_mode == MODE_STOP)
			{
				txtOnTime.Text	= save_on_time.ToString();
				txtOffTime.Text = save_off_time.ToString();				
			}
			else
				timer1.Enabled = true;

						
		}
		
		void resetOnOffTime()
		{
			txtOnTime.Text	= save_on_time.ToString();
			txtOffTime.Text = save_off_time.ToString();
			txtOnTime.BackColor =  System.Drawing.SystemColors.Window;
			txtOffTime.BackColor =  System.Drawing.SystemColors.Window;
		}
		
		void verifyMacFunction()
		{
			int value = 0 ;
			for (int i = 0; i != arrCheck.Count; i++)
			{
				if (arrCheck[i].Checked == true)
				{
					Color bkcolor = arrMacText[i].BackColor;
					arrMacText[i].BackColor = System.Drawing.Color.Yellow;

					Application.DoEvents();
					System.Threading.Thread.Sleep(20);
					
					arrMacText[i].Text = GetMacAddress(arrIPText[i].Text);
										
                    /*
					if (arrMacText[i].Text.Trim() != "00-00-00-00-00-00" && 
					    arrMacText[i].Text.Length == 17)
					{
						value = Convert.ToInt32(arrPassText[i].Text) + 1;
						arrPassText[i].Text = value.ToString();
					}*/				
										
					arrMacText[i].BackColor = bkcolor;
				}				
			}	
		}

        void telnetFunction()
        {
            for (int i = 0; i != arrCheck.Count; i++)
            {
                if (arrCheck[i].Checked == true && ping1OK[i] == true)
                {
                    Application.DoEvents();
                    System.Threading.Thread.Sleep(20);

                    arrStatus[i].BackColor = System.Drawing.Color.LightPink;

                    if (runTelectScript(arrIPText[i].Text) == false)
                    {
                        telnetFail[i]++;
                        arrTelFailText[i].Text = telnetFail[i].ToString();                        
                        arrRes[i].BackColor = System.Drawing.Color.Red;
                    }

                    Application.DoEvents();
                    System.Threading.Thread.Sleep(200);
                    arrStatus[i].BackColor = System.Drawing.Color.LightGreen;
                }
            }
        }

        bool verifyIPFunction(int id)
		{
			bool result = false;
			int value = 0;

			for (int i = 0; i != arrCheck.Count; i++)
			{
                if (chkTelnet.Checked == false)
                    ping1OK[i] = true;

				if (arrCheck[i].Checked == true)
				{	
					arrStatus[i].BackColor = System.Drawing.Color.Yellow;
					result = false;
					
					Application.DoEvents();
					System.Threading.Thread.Sleep(20);

					DateTime dt = DateTime.Now;
											
					for (int j = 0; j != Convert.ToInt32(txtRetry.Text); j++)
					{
                        if (id == 0)
                            ping1OK[i] = false;

                        Application.DoEvents();
                        System.Threading.Thread.Sleep(20);

						PingReply reply = Ping(arrIPText[i].Text, Convert.ToInt32(txtTimeout.Text));

                        Application.DoEvents();
                        System.Threading.Thread.Sleep(20);
													
						if (reply.Status == IPStatus.Success)
						{
                            if (id == 0)
                            {
                                ping1OK[i] = true;
                            }
                            else if (id == 1 && ping1OK[i] == true)
                            {
                                value = Convert.ToInt32(arrPassText[i].Text) + 1;
                                arrPassText[i].Text = value.ToString();
                            }

							arrStatus[i].BackColor = System.Drawing.Color.LightGreen;
							result = true;
							break;
						}
                        else
						{
							saveLogFunction(dt.ToString() + " pin " + arrIPText[i].Text + "\r\n", true);
							saveLogFunction("result : " + reply.Status.ToString() + "\r\n", true);
                            txtLog.SelectionStart = txtLog.Text.Length;
                            txtLog.ScrollToCaret();
												
							arrStatus[i].BackColor = System.Drawing.Color.Red;                            
							result = false;
						}
					}

                    Application.DoEvents();
                    System.Threading.Thread.Sleep(20);
											
					if (result == false)
					{
                        arrRes[i].BackColor = System.Drawing.Color.Red;

                        saveLogFunction("=========================================\r\n", true);

                        if (id == 0)
                        {
                            ping1Fail[i]++;
                            arrFailText[i].Text = ping1Fail[i].ToString();	
                        }
                        else
                        {
                            ping2Fail[i]++;
                            arrFail2Text[i].Text = ping2Fail[i].ToString();
                        }
	
						
						if (chkFailStop.Checked == true)
						{
							timer1.Enabled = false;
							btnStart.Enabled = true;
							btnStop.Enabled = false;
							btnPause.Enabled = true;
							
							resetOnOffTime();
							
							picStart.Visible = false;
							picStop.Visible = true;								
							
							now_mode = MODE_STOP;

                            if (chkStopOff.Checked == true)
								powerOffFunction();
							
							//MsgBox "Found ping " + txtIPlst(T).Text + " failure in " + Format(Now, "yyyy/m/d hh:mm:ss")
							MessageBox.Show("Found ping " + arrIPText[i].Text + " failure in " + dt.ToString());
							break;
						}
					}
				}
                				
				if (testbutton == false && now_mode == MODE_STOP)
					break;		
			}
            return result;
		}		

		void powerOnFunction()
		{			
			//Out32(888, 0);
            ComSetHigh();
            powerStatusOn = true;
		}
		
		void powerOffFunction()
		{
			//Out32(888, 255);
            ComSetLow();
            powerStatusOn = false;
		}			
		
		void writeFile(string file, string content, bool append)
		{

			try
   			{
		        using (StreamWriter sw = new StreamWriter(file, append))
		        {
		            sw.Write(content);
                    sw.Close();
		        }
			}
		    catch (IOException ex)
		    {
		    }			
		}

        String readFile(string file, string content, bool append)
        {
            String str = "";

            try
            {
                using (StreamReader sr = new StreamReader(file))
                {
                    str = sr.ReadToEnd();
                    sr.Close();
                }
            }
            catch (IOException ex)
            {
            }

            return str;
        }
		
		void saveLogFunction(string content, bool append)
		{		
			if (append == false)
			{
				txtLog.Text = "";
			}

            writeFile(txtLogPath.Text, content, append);			
			txtLog.Text += content;
            txtLog.SelectionStart = txtLog.Text.Length;
            txtLog.ScrollToCaret();
            txtLog.Focus();
		}
	
		void BtnOpenClick(object sender, EventArgs e)
		{
			Process notePad = new Process();
			notePad.StartInfo.FileName = "notepad.exe ";
            notePad.StartInfo.Arguments = txtLogPath.Text;
			notePad.Start();
		}


        private void loadConfigFile()
        {
            if (System.IO.File.Exists(configFile))
            {
                try
                {
                    using (StreamReader sr = new StreamReader(configFile))
                    {
                        txtProject.Text = sr.ReadLine();
                        save_on_time = Convert.ToInt32(sr.ReadLine());
                        save_off_time = Convert.ToInt32(sr.ReadLine());
                        txtTimeout.Text = sr.ReadLine();
                        txtRetry.Text = sr.ReadLine();

                        txtOnTime.Text = save_on_time.ToString();
                        txtOffTime.Text = save_off_time.ToString();

                        for (int i = 0; i != arrCheck.Count; i++)
                        {
                            String check = sr.ReadLine();
                            if (check.CompareTo("True") == 0)
                                arrCheck[i].Checked = true;
                            else
                                arrCheck[i].Checked = false;
                            arrIPText[i].Text = sr.ReadLine();
                        }

                        sr.Close();
                    }
                }
                catch (IOException ex)
                {
                }
            }
        }

        private void chkStopPowerOff_FormClosed(object sender, FormClosedEventArgs e)
        {

            if (save_on_time > 0)
                txtOnTime.Text = save_on_time.ToString();

            if (save_off_time > 0)
                txtOffTime.Text = save_off_time.ToString();	

            String save = txtProject.Text + "\r\n" +
                          txtOnTime.Text + "\r\n" +
                          txtOffTime.Text + "\r\n" +
                          txtTimeout.Text + "\r\n" +
                          txtRetry.Text + "\r\n";

            for (int i = 0; i != arrCheck.Count; i++)
            {
                save += arrCheck[i].Checked.ToString() + "\r\n";
                save += arrIPText[i].Text + "\r\n";
            }

            writeFile(configFile, save, false);


            if (openFlag == true)
            {
                serialPort.Close();
            }

        }


        private void btnComOpen_Click(object sender, EventArgs e)
        {
            if (openFlag == false)
            {
                serialPort.PortName = comboBox1.Text;
                serialPort.BaudRate = 9600;            // baud rate = 9600
                serialPort.Parity = Parity.None;       // Parity = none
                serialPort.StopBits = StopBits.One;    // stop bits = one
                serialPort.DataBits = 8;               // data bits = 8

                serialPort.Open();
                serialPort.DiscardInBuffer();       // RX
                serialPort.DiscardOutBuffer();      // TX
                openFlag = true;
                txtLog.Text += comboBox1.Text + " Open!\r\n";
                comboBox1.Enabled = false;
                btnComOpen.Enabled = false;

                this.Text += " - " + comboBox1.Text;
            }
        }

        private void btnComClose_Click(object sender, EventArgs e)
        {
            if (openFlag == true)
            {
                serialPort.Close();
                openFlag = false;
                txtLog.Text += comboBox1.Text + " Close!\r\n";
                comboBox1.Enabled = true;
                btnComOpen.Enabled = true;
            }
        }

        void ComSetHigh()
        {
            if (openFlag == true)
            {
                serialPort.RtsEnable = true;
                serialPort.DtrEnable = true;
                //txtLog.Text += comboBox1.Text + " RTS & DTR High!\r\n";
            }
        }

        void ComSetLow()
        {
            if (openFlag == true)
            {
                serialPort.RtsEnable = false;
                serialPort.DtrEnable = false;
                //txtLog.Text += comboBox1.Text + " RTS & DTR Low!\r\n";
            }
        }

        private void btnScriptOpen_Click(object sender, EventArgs e)
        {
            //runTelectScript("192.168.1.179");

            Process notePad = new Process();
            notePad.StartInfo.FileName = "notepad.exe ";
            notePad.StartInfo.Arguments = txtScriptPath.Text;
            notePad.Start();
        }

        bool runTelectScript(String ip)
        {
            string line;
            bool connect = false;

            TelnetConnection tc = new TelnetConnection();
            connect = tc.TelnetInit(ip, 23);

            if (connect == false)
            {
                txtLog.Text += ip + " telnet fail!\r\n";
                return false;
            }

            string s = tc.Read();
            string prompt = s.TrimEnd();
            prompt = s.Substring(prompt.Length - 1, 1);
            if (prompt != "$" && prompt != ">" && prompt != "#" && prompt != ":")
            {
                txtLog.Text += ip + " telnet fail!\r\n";
                return false;
            }

            txtLog.Text += "telnet " + ip + " ...\r\n";

            if (File.Exists(txtScriptPath.Text) == true)
            {
                // Read the file and display it line by line.
                System.IO.StreamReader file =
                    new System.IO.StreamReader(txtScriptPath.Text);
                while ((line = file.ReadLine()) != null)
                {
                    txtLog.Text += " " + line + "\r\n";
                    tc.WriteLine(line);
                    System.Threading.Thread.Sleep(1000);
                }
                file.Close();
            }
            else
                txtLog.Text += "Can't find " + txtScriptPath.Text;
            
            //tc.WriteLine("exit");
            return true;
        }


        
	}
}

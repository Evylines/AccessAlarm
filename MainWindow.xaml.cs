using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace AccessAlarm
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        [DllImport(@"User32", SetLastError = true, EntryPoint = "RegisterPowerSettingNotification",
            CallingConvention = CallingConvention.StdCall)]

        private static extern IntPtr RegisterPowerSettingNotification(IntPtr hRecipient, ref Guid PowerSettingGuid,
            Int32 Flags);

        internal struct POWERBROADCAST_SETTING
        {
            public Guid PowerSetting;
            public uint DataLength;
            public byte Data;
        }
        Guid GUID_LIDOPEN_POWERSTATE = new Guid(0x99FF10E7, 0x23B1, 0x4C07, 0xA9, 0xD1, 0x5C, 0x32, 0x06, 0xD7, 0x41, 0xB4);
        Guid GUID_LIDSWITCH_STATE_CHANGE = new Guid(0xBA3E0F4D, 0xB817, 0x4094, 0xA2, 0xD1, 0xD5, 0x63, 0x79, 0xE6, 0xA0, 0xF3);
        const int DEVICE_NOTIFY_WINDOW_HANDLE = 0x00000000;
        const int WM_POWERBROADCAST = 0x0218;
        const int PBT_POWERSETTINGCHANGE = 0x8013;

        private bool? _previousLidState = null;
        // property variables
        private int m_TimeToCapture_milliseconds = 100;
        private int m_Width = 320;
        private int m_Height = 240;
        private int mCapHwnd;
        private bool bStopped = true;
        IntPtr handle;

        private ulong m_FrameNumber = 0;

        public MainWindow()
        {
            InitializeComponent();
            this.SourceInitialized += MainWindow_SourceInitialized;
        }

        ~MainWindow()
        {
            Stop();
        }

        /// <summary>
        /// Starts the video capture
        /// </summary>
        /// <param name="FrameNumber">the frame number to start at. 
        /// Set to 0 to let the control allocate the frame number</param>
        public void Start(ulong FrameNum)
        {
            try
            {
                // for safety, call stop, just in case we are already running
                //this.Stop();
                var i = 0;
                // setup a capture window
                mCapHwnd = capCreateCaptureWindowA("WebCap", 0, 0, 0, m_Width, m_Height, handle.ToInt32(), 0);
                Thread.Sleep(0);
                
                i = SendMessage(mCapHwnd, WM_CAP_CONNECT, 0, 0);
                SendMessage(mCapHwnd, WM_CAP_SET_PREVIEW, 0, 0);

                // set the frame number
                m_FrameNumber = FrameNum;
                this.ImageCaptured += MainWindow_ImageCaptured;

                //// set the timer information
                //this.timer1.Interval = m_TimeToCapture_milliseconds;
                //bStopped = false;
                //this.timer1.Start();
            }

            catch (Exception excep)
            {
                MessageBox.Show("An error ocurred while starting the video capture. Check that your webcamera is connected properly and turned on.\r\n\n" + excep.Message);
                this.Stop();
            }
        }
        public delegate void WebCamEventHandler(object source, WebcamEventArgs e);
        public class WebcamEventArgs : System.EventArgs
        {
            private BitmapSource m_Image;
            private ulong m_FrameNumber = 0;

            public WebcamEventArgs()
            {
            }

            /// <summary>
            ///  WebCamImage
            ///  This is the image returned by the web camera capture
            /// </summary>
            public BitmapSource WebCamImage
            {
                get
                { return m_Image; }

                set
                { m_Image = value; }
            }

            /// <summary>
            /// FrameNumber
            /// Holds the sequence number of the frame capture
            /// </summary>
            public ulong FrameNumber
            {
                get
                { return m_FrameNumber; }

                set
                { m_FrameNumber = value; }
            }
        }
        // fired when a new image is captured
        public event WebCamEventHandler ImageCaptured;
        private WebcamEventArgs x = new WebcamEventArgs();

        private IDataObject tempObj;
        private BitmapSource tempImg;

        /// <summary>
        /// Capture the next frame from the video feed
        /// </summary>
        private void _captureImage()
        {
            try
            {
                //// pause the timer
                //this.timer1.Stop();

                // get the next frame;
                SendMessage(mCapHwnd, WM_CAP_GET_FRAME, 0, 0);

                // copy the frame to the clipboard
                SendMessage(mCapHwnd, WM_CAP_COPY, 0, 0);

                // paste the frame into the event args image
                if (ImageCaptured != null)
                {
                    // get from the clipboard
                    tempObj = Clipboard.GetDataObject();

                    Array pixels = new int[m_Width * m_Height * 4];
                    var source = ((InteropBitmap)tempObj.GetData(DataFormats.Bitmap));
                    source.CopyPixels(pixels, 8 * m_Width, 0);

                    WriteableBitmap wb = new WriteableBitmap(source);
                    wb.WritePixels(new Int32Rect(0, 0, m_Width * 2, m_Height * 2), pixels, 8 * m_Width, 0);

                    JpegBitmapEncoder encoder = new JpegBitmapEncoder();

                    String photolocation = "D:\\testCapture.jpg";  //file name 

                    encoder.Frames.Add(BitmapFrame.Create(wb));

                    using (var filestream = new FileStream(photolocation, FileMode.Create))
                        encoder.Save(filestream);
                    EmptyClipboard();
                    /*
                    * For some reason, the API is not resizing the video
                    * feed to the width and height provided when the video
                    * feed was started, so we must resize the image here
                    */
                    x.WebCamImage = wb;
                    //(m_Width, m_Height, null, System.IntPtr.Zero);

                    // raise the event
                    this.ImageCaptured(this, x);
                }

            }

            catch (Exception excep)
            {
                MessageBox.Show("An error ocurred while capturing the video image. The video capture will now be terminated.\r\n\n" + excep.Message);
                this.Stop(); // stop the process
            }
        }

        /// <summary>
        /// Stops the video capture
        /// </summary>
        public void Stop()
        {
            try
            {
                // stop the timer
                bStopped = true;
                //this.timer1.Stop();

                //// disconnect from the video source
                //Application.DoEvents();
                SendMessage(mCapHwnd, WM_CAP_DISCONNECT, 0, 0);
                CloseClipboard();
            }

            catch (Exception excep)
            { // don't raise an error here.
                MessageBox.Show("An error ocurred while starting the video capture. Check that your webcamera is connected properly and turned on.\r\n\n" + excep.Message);

            }

        }

        void MainWindow_SourceInitialized(object sender, EventArgs e)
        {
            RegisterForPowerNotifications();
            IntPtr hwnd = new WindowInteropHelper(this).Handle;
            HwndSource.FromHwnd(hwnd).AddHook(new HwndSourceHook(WndProc));

            Start(100);
        }

        private void RegisterForPowerNotifications()
        {
            handle = new WindowInteropHelper(this).Handle;
            IntPtr hLIDSWITCHSTATECHANGE = RegisterPowerSettingNotification(handle,
                 ref GUID_LIDSWITCH_STATE_CHANGE,
                 DEVICE_NOTIFY_WINDOW_HANDLE);

        }

        IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            switch (msg)
            {
                case WM_POWERBROADCAST:
                    OnPowerBroadcast(wParam, lParam);
                    break;
                default:
                    break;
            }
            return IntPtr.Zero;
        }

        private void OnPowerBroadcast(IntPtr wParam, IntPtr lParam)
        {
            if ((int)wParam == PBT_POWERSETTINGCHANGE)
            {
                POWERBROADCAST_SETTING ps = (POWERBROADCAST_SETTING)Marshal.PtrToStructure(lParam, typeof(POWERBROADCAST_SETTING));
                IntPtr pData = (IntPtr)((int)lParam + Marshal.SizeOf(ps));
                Int32 iData = (Int32)Marshal.PtrToStructure(pData, typeof(Int32));
                if (ps.PowerSetting == GUID_LIDSWITCH_STATE_CHANGE)
                {
                    bool isLidOpen = ps.Data != 0;

                    if (!isLidOpen == _previousLidState)
                    {
                        LidStatusChanged(isLidOpen);
                    }

                    _previousLidState = isLidOpen;
                }
            }
        }

        private void LidStatusChanged(bool isLidOpen)
        {
            if (isLidOpen)
            {
                //Do some action on lid open event
                lstBx.Items.Add(String.Format("{0}: Lid opened!", DateTime.Now));
                Thread.Sleep(2000);
                CapturePhoto();
            }
        }

        private void CapturePhoto()
        {
            _captureImage();
        }

        void MainWindow_ImageCaptured(object source, MainWindow.WebcamEventArgs e)
        {
            img.Source = e.WebCamImage;
            //smtp сервер
            string smtpHost = "smtp.gmail.com";
            //smtp порт
            int smtpPort = 587;
            //логин
            string login = "/%write_your_login%/";
            //пароль
            string pass = "/%write_your_pass%/";

            //создаем подключение
            SmtpClient client = new SmtpClient(smtpHost, smtpPort);
            client.EnableSsl = true;

            client.Credentials = new NetworkCredential(login, pass);

            //От кого письмо
            string from = "/%write_your_mail%/";
            //Кому письмо
            string to = "/%write_your_mail%/";
            //Тема письма
            string subject = "Несанкционированный доступ к Вашему компьютеру";
            //Текст письма
            string body = "Внимание! \n\n\n К Вашему компьютеру осуществлен несанкционированный доступ";

            //Создаем сообщение
            MailMessage mess = new MailMessage(from, to, subject, body);

            //Вложение для письма
            //Если нужно не одно вложение, для каждого создаем отдельный Attachment
            Attachment attData = new Attachment(@"D:\testCapture.jpg");

            ////прикрепляем вложение
            mess.Attachments.Add(attData);

            try
            {
                client.Send(mess);
                lstBx.Items.Add("Message send");
            }
            catch (Exception ex)
            {
                lstBx.Items.Add(ex.ToString());
            }
        }

        #region API Declarations

        [DllImport("user32", EntryPoint = "SendMessage")]
        public static extern int SendMessage(int hWnd, uint Msg, int wParam, int lParam);

        [DllImport("avicap32.dll", EntryPoint = "capCreateCaptureWindowA")]
        public static extern int capCreateCaptureWindowA(string lpszWindowName, int dwStyle, int X, int Y, int nWidth, int nHeight, int hwndParent, int nID);

        [DllImport("user32", EntryPoint = "OpenClipboard")]
        public static extern int OpenClipboard(int hWnd);

        [DllImport("user32", EntryPoint = "EmptyClipboard")]
        public static extern int EmptyClipboard();

        [DllImport("user32", EntryPoint = "CloseClipboard")]
        public static extern int CloseClipboard();

        #endregion

        #region API Constants

        public const int WM_USER = 1024;

        public const int WM_CAP_CONNECT = 1034;
        public const int WM_CAP_DISCONNECT = 1035;
        public const int WM_CAP_GET_FRAME = 1084;
        public const int WM_CAP_COPY = 1054;

        public const int WM_CAP_START = WM_USER;

        public const int WM_CAP_DLG_VIDEOFORMAT = WM_CAP_START + 41;
        public const int WM_CAP_DLG_VIDEOSOURCE = WM_CAP_START + 42;
        public const int WM_CAP_DLG_VIDEODISPLAY = WM_CAP_START + 43;
        public const int WM_CAP_GET_VIDEOFORMAT = WM_CAP_START + 44;
        public const int WM_CAP_SET_VIDEOFORMAT = WM_CAP_START + 45;
        public const int WM_CAP_DLG_VIDEOCOMPRESSION = WM_CAP_START + 46;
        public const int WM_CAP_SET_PREVIEW = WM_CAP_START + 50;

        #endregion



    }

}

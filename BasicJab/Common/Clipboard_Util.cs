using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace BasicJab.Common
{
    public class Clipboard_Util
    {
        const uint cfUnicodeText = 13;

        public string GetText()
        {
            if (!IsClipboardFormatAvailable(cfUnicodeText))
            {
                return null;
            }

            TryOpenClipboard();

            return InnerGet();
        }

        public void SetText(string text)
        {
            TryOpenClipboard();

            InnerSet(text);
        }

        static void InnerSet(string text)
        {
            EmptyClipboard();
            IntPtr hGlobal = IntPtr.Zero;
            try
            {
                var bytes = (text.Length + 1) * 2;
                hGlobal = Marshal.AllocHGlobal(bytes);

                if (hGlobal == IntPtr.Zero)
                {
                    ThrowWin32();
                }

                var target = GlobalLock(hGlobal);

                if (target == IntPtr.Zero)
                {
                    ThrowWin32();
                }

                try
                {
                    Marshal.Copy(text.ToCharArray(), 0, target, text.Length);
                }
                finally
                {
                    GlobalUnlock(target);
                }

                if (SetClipboardData(cfUnicodeText, hGlobal) == IntPtr.Zero)
                {
                    ThrowWin32();
                }

                hGlobal = IntPtr.Zero;
            }
            finally
            {
                if (hGlobal != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(hGlobal);
                }

                CloseClipboard();
            }
        }

        static string InnerGet()
        {
            IntPtr handle = IntPtr.Zero;

            IntPtr pointer = IntPtr.Zero;
            try
            {
                handle = GetClipboardData(cfUnicodeText);
                if (handle == IntPtr.Zero)
                {
                    return null;
                }

                pointer = GlobalLock(handle);
                if (pointer == IntPtr.Zero)
                {
                    return null;
                }

                var size = GlobalSize(handle);
                var buff = new byte[size];

                Marshal.Copy(pointer, buff, 0, size);

                return Encoding.Unicode.GetString(buff).TrimEnd('\0');
            }
            finally
            {
                if (pointer != IntPtr.Zero)
                {
                    GlobalUnlock(handle);
                }

                CloseClipboard();
            }
        }

        static void TryOpenClipboard()
        {
            var num = 10;
            while (true)
            {
                if (OpenClipboard(IntPtr.Zero))
                {
                    break;
                }

                if (--num == 0)
                {
                    ThrowWin32();
                }

                Thread.Sleep(100);
            }
        }

        static void ThrowWin32()
        {
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        [DllImport("User32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsClipboardFormatAvailable(uint format);

        [DllImport("User32.dll", SetLastError = true)]
        static extern IntPtr GetClipboardData(uint uFormat);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr SetClipboardData(uint uFormat, IntPtr data);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr GlobalLock(IntPtr hMem);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GlobalUnlock(IntPtr hMem);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool CloseClipboard();



        [DllImport("user32.dll")]
        static extern bool EmptyClipboard();

        [DllImport("Kernel32.dll", SetLastError = true)]
        static extern int GlobalSize(IntPtr hMem);
    }

}

﻿using System;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading;
using System.Drawing;
namespace ConsoleApp_FindButtonPosition
{
    class Program
    {
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        // --- Windows API 函數匯入 ---
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        // --- 常數定義 ---
        private const int SW_RESTORE = 9;                 // 恢復最小化窗口的命令
        private const int VK_ESCAPE = 0x1B;              // ESC 鍵虛擬鍵碼
        private const int MOUSEEVENTF_LEFTDOWN = 0x0002; // 滑鼠左鍵按下
        private const int MOUSEEVENTF_LEFTUP = 0x0004;   // 滑鼠左鍵釋放
        private const int MOUSEEVENTF_RIGHTDOWN = 0x0008; // 滑鼠右鍵按下
        private const int MOUSEEVENTF_RIGHTUP = 0x0010;  // 滑鼠右鍵釋放

        // --- 結構定義 ---
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        // --- 全域變數 ---
        private static string filePath = Path.Combine(GetDesktopPath(), "QMOrder");
        private static System.Threading.Mutex mutex; // 用於防止重複執行
        private static int cycleTime = 1000; // 循環間隔時間（毫秒）

        // --- 主程式 ---
        [STAThread]
        static void Main(string[] args)
        {
            // 防止程式重複啟動
            mutex = new System.Threading.Mutex(true, "ConsoleApp_UDS022P");
            if (!mutex.WaitOne(0, false))
            {
                return;
            }

            // 確保資料夾存在
            EnsureDirectoryExists(filePath);
            CleanOldFiles(filePath, 5);
            while (true)
            {
                // 檢查是否按下 ESC 鍵
                if (IsKeyPressed(VK_ESCAPE))
                {
                    Console.WriteLine("檢測到 ESC 鍵，程式即將停止...");
                    Thread.Sleep(3000);
                    break;
                }

                // 處理 UDS022P 程式的操作
                HandleUDS022PProcess();

                // 等待下一次執行
                Thread.Sleep(cycleTime);
            }
        }

        // --- 功能區塊 ---
        /// <summary>
        /// 檢查指定的鍵是否被按下
        /// </summary>
        static bool IsKeyPressed(int keyCode)
        {
            return (GetAsyncKeyState(keyCode) & 0x8000) != 0;
        }

        /// <summary>
        /// 確保指定的資料夾存在，若不存在則自動建立
        /// </summary>
        static void EnsureDirectoryExists(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
                Console.WriteLine($"資料夾已建立: {path}");
            }
            else
            {
                Console.WriteLine($"資料夾已存在: {path}");
            }
        }

        /// <summary>
        /// 獲取使用者桌面路徑
        /// </summary>
        static string GetDesktopPath()
        {
            return Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        }

        /// <summary>
        /// 模擬滑鼠左鍵點擊
        /// </summary>
        static void SimulateMouseLeftClick(int x, int y)
        {
            SetCursorPos(x, y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, x, y, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, x, y, 0, 0);
        }

        /// <summary>
        /// 模擬滑鼠右鍵點擊
        /// </summary>
        static void SimulateMouseRightClick(int x, int y)
        {
            SetCursorPos(x, y);
            mouse_event(MOUSEEVENTF_RIGHTDOWN, x, y, 0, 0);
            mouse_event(MOUSEEVENTF_RIGHTUP, x, y, 0, 0);
        }

        /// <summary>
        /// 將文字複製到剪貼簿
        /// </summary>
        static void CopyToClipboard(string text)
        {
            Clipboard.SetText(text);
        }

        /// <summary>
        /// 處理 UDS022P 程式的操作邏輯
        /// </summary>
        static void HandleUDS022PProcess()
        {
            string targetWindowName = "UDS022P"; // 替換為目標視窗名稱
            RECT rECT_UDS022P = FindWindowAndChildrenBySimular("程式【UDS022P】");
            if ((rECT_UDS022P.Left != 0 || rECT_UDS022P.Right != 0) == false)
            {
                Console.WriteLine($"{DateTime.Now:yyyy/MM/dd HH:mm:ss} - 程式未開啟...");
                return;
            }
            //MoveMouseToScreenTopLeft(1);
            SimulateMouseLeftClick(rECT_UDS022P.Left + 48, rECT_UDS022P.Top + 89);
            SendKeys.SendWait("{ENTER}"); // 模擬按下 Enter
            SimulateMouseLeftClick(rECT_UDS022P.Left + 310, rECT_UDS022P.Top + 82);
            Thread.Sleep(100);
            RECT rECT_另存新檔 = new RECT();
            while (true)
            {
                rECT_另存新檔 = FindWindowAndChildren("另存新檔");
                if (rECT_另存新檔.Left != 0 || rECT_另存新檔.Right != 0) break;
                Thread.Sleep(100);
            }
       
            SimulateMouseRightClick(rECT_另存新檔.Left + 302, rECT_另存新檔.Top + 343);
            Thread.Sleep(100);
            SendKeys.SendWait("A"); // 模擬輸入 'A'
            CopyToClipboard($@"{filePath}\{DateTime.Now:yyyyMMddHHmmss}.csv");
            SendKeys.SendWait("^v"); // 模擬 Ctrl + V
            Thread.Sleep(100);
            SimulateMouseLeftClick(rECT_另存新檔.Left + 503, rECT_另存新檔.Top + 344);
            Thread.Sleep(100);

            RECT rECT_程式 = new RECT();
            while (true)
            {
                rECT_程式 = FindWindowAndChildren("程式");
                if (rECT_程式.Left != 0 || rECT_程式.Right != 0) break;
                Thread.Sleep(100);
            }
            SendKeys.SendWait("{ENTER}"); // 模擬按下 Enter
        }

        /// <summary>
        /// 尋找指定名稱的視窗
        /// </summary>
        static RECT FindWindowAndChildren(string windowName)
        {
            Process[] processes = Process.GetProcesses();
            RECT rECT = new RECT();

            foreach (var process in processes)
            {
                EnumWindows((hWnd, lParam) =>
                {
                    GetWindowThreadProcessId(hWnd, out uint windowProcessId);

                    if (windowProcessId == process.Id)
                    {
                        StringBuilder windowText = new StringBuilder(256);
                        GetWindowText(hWnd, windowText, 256);

                        if (windowText.ToString() == windowName)
                        {
                            Console.WriteLine($"找到視窗: {windowText}");

                            if (GetWindowRect(hWnd, out RECT rect))
                            {
                                Console.WriteLine($"視窗座標: 左上角 ({rect.Left}, {rect.Top}), 右下角 ({rect.Right}, {rect.Bottom})");
                                rECT = rect;
                            }

                            ShowWindow(hWnd, SW_RESTORE);
                            if (SetForegroundWindow(hWnd))
                            {
                                Console.WriteLine($"[{windowName}]視窗成功設置為前景！");
                            }
                        }
                    }
                    return true;
                }, IntPtr.Zero);
            }
            return rECT;
        }
        /// <summary>
        /// 尋找指定名稱的視窗
        /// </summary>
        static RECT FindWindowAndChildrenBySimular(string windowName)
        {
            Process[] processes = Process.GetProcesses();
            RECT rECT = new RECT();

            foreach (var process in processes)
            {
                EnumWindows((hWnd, lParam) =>
                {
                    GetWindowThreadProcessId(hWnd, out uint windowProcessId);

                    if (windowProcessId == process.Id)
                    {
                        StringBuilder windowText = new StringBuilder(256);
                        GetWindowText(hWnd, windowText, 256);

                        if (windowText.ToString().Contains(windowName))
                        {
                            Console.WriteLine($"找到視窗: {windowText}");

                            if (GetWindowRect(hWnd, out RECT rect))
                            {
                                Console.WriteLine($"視窗座標: 左上角 ({rect.Left}, {rect.Top}), 右下角 ({rect.Right}, {rect.Bottom})");
                                rECT = rect;
                            }

                            ShowWindow(hWnd, SW_RESTORE);
                            if (SetForegroundWindow(hWnd))
                            {
                                Console.WriteLine($"[{windowName}]視窗成功設置為前景！");
                            }
                        }
                    }
                    return true;
                }, IntPtr.Zero);
            }
            return rECT;
        }
        /// <summary>
        /// 清理指定資料夾中前一分鐘的檔案
        /// </summary>
        /// <param name="folderPath">資料夾路徑</param>
        static void CleanOldFiles(string folderPath , int min)
        {
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine($"資料夾不存在: {folderPath}");
                return;
            }

            try
            {
                var files = Directory.GetFiles(folderPath);
                DateTime oneMinuteAgo = DateTime.Now.AddMinutes(-min);

                foreach (var file in files)
                {
                    var fileInfo = new FileInfo(file);
                    if (fileInfo.CreationTime < oneMinuteAgo)
                    {
                        fileInfo.Delete();
                        Console.WriteLine($"已刪除檔案: {file}");
                    }
                }

                Console.WriteLine("資料夾清理完成。");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"清理過程中發生錯誤: {ex.Message}");
            }
        }
        /// <summary>
        /// 將滑鼠移動到指定螢幕的左上角
        /// </summary>
        /// <param name="screenNumber">螢幕編號（從 1 開始）</param>
        /// <returns>是否成功移動</returns>
        static bool MoveMouseToScreenTopLeft(int screenNumber)
        {
            // 驗證螢幕編號是否有效
            if (screenNumber < 1 || screenNumber > Screen.AllScreens.Length)
            {
                Console.WriteLine("無效的螢幕編號！");
                return false;
            }

            // 獲取指定編號的螢幕
            Screen targetScreen = Screen.AllScreens[screenNumber - 1];

            // 螢幕左上角的座標
            int x = targetScreen.Bounds.Left;
            int y = targetScreen.Bounds.Top;

            // 移動滑鼠
            SetCursorPos(x, y);

            Console.WriteLine($"滑鼠已移動到第 {screenNumber} 號螢幕的左上角: ({x}, {y})");
            return true;
        }
    }
}

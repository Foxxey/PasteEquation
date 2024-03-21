using System.Runtime.InteropServices;
using System.Windows.Forms;
using System;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PasteEquation
{
    class WordFunctions
    {
        private static void Paste(string input)
        {
            Clipboard.SetText(input);
            Range currentRange = Globals.Main.Application.Selection.Range;
            currentRange.Paste();
            Clipboard.Clear();
        }

        private static bool SplitAndPaste(string input)
        {
            input = " " + input;
            Regex mathRegex = new Regex(@"<math [\S\s]*?>[\S\s]*?<\/math>");

            string[] arr1 = mathRegex.Split(input);
            string[] arr2 = mathRegex.Matches(input).Cast<Match>().Select(m => m.Value).ToArray();
            if (arr2.Length == 0) return false;

            List<string> returnArr = arr1.SelectMany((el, index) => new[] { el, arr2.ElementAtOrDefault(index) }).ToList();

            returnArr.RemoveAll(string.IsNullOrEmpty);
            for (int i = returnArr.Count - 1; i >= 0; i--)
            {
                Paste(returnArr[i]);
            }
            return true;
        }

        public static bool PasteEquation()
        {
            Range currentRange = Globals.Main.Application.Selection.Range;
            currentRange.Text = " ";
            currentRange.SetRange(currentRange.End, currentRange.End);
            currentRange.Select();

            string clipboardText = Clipboard.GetText();

            bool returnVal = true;
            if (clipboardText == string.Empty || !SplitAndPaste(clipboardText)) returnVal = false;
            else Clipboard.SetText(clipboardText);

            currentRange.SetRange(currentRange.Start - 1, currentRange.Start - 1);
            currentRange.Delete(WdUnits.wdCharacter, 2);
            return returnVal;
        }
    }

    public class KeyboardHook {
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod,
            uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        public delegate int LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private static LowLevelKeyboardProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;


        private const int WH_KEYBOARD = 2;
        private const int HC_ACTION = 0;
        public static void SetHook()
        {
#pragma warning disable 618
            _hookID = SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, (uint)AppDomain.GetCurrentThreadId());
#pragma warning restore 618

        }

        public static void ReleaseHook()
        {
            UnhookWindowsHookEx(_hookID);
        }
        private static int HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0)
            {
                return (int)CallNextHookEx(_hookID, nCode, wParam, lParam);
            }
            else
            {

                if (nCode == HC_ACTION)
                {
                    Keys keyData = (Keys)wParam;

                    if ((BindingFunctions.IsKeyDown(Keys.ControlKey) == true)
                    && (BindingFunctions.IsKeyDown(keyData) == true) && (keyData == Keys.V))
                    {
                        if (WordFunctions.PasteEquation()) return 1;
                    }

                }
                return (int)CallNextHookEx(_hookID, nCode, wParam, lParam);
            }
        }
    }

    public class BindingFunctions
    {
        [DllImport("user32.dll")]
        static extern short GetKeyState(int nVirtKey);

        public static bool IsKeyDown(Keys keys)
        {
            return (GetKeyState((int)keys) & 0x8000) == 0x8000;
        }

    }
}
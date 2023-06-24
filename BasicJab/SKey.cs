using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;

namespace BasicJab
{
    [Guid("37E9197B-3311-4DFA-953C-A6E894ABFA31"), ComVisible(true)]
    public enum SKey
    {
        [Description("Ctrl-A")] Select_All,
        [Description("Ctrl-C")] Copy,
        [Description("Ctrl-X")] Cut,
        [Description("Ctrl-V")] Paste,
        [Description("Esc")] ESC,

        [Description("F5")] Oracle_Clear_Field,
        [Description("F7")] Oracle_Clear_Block,
        [Description("F8")] Oracle_Clear_Form,
        [Description("F6")] Oracle_Clear_Record,

        [Description("Ctrl-B")] Oracle_Block_Menu,
        [Description("Ctrl-S")] Oracle_Commit,
        [Description("F12")] Oracle_Count_Query,
        [Description("Ctrl-UpArrow")] Oracle_Delete_Record,
        [Description("Shift-Ctrl-E")] Oracle_Display_Error,

        [Description("Shift-F5")] Oracle_Duplicate_Field,
        [Description("Shift-F6")] Oracle_Duplicate_Record,
        [Description("Ctrl-E")] Oracle_Edit,
        [Description("F11")] Oracle_Enter_Query,
        [Description("Shift-F11")] Oracle_Excute_Query,
        [Description("F4")] Oracle_Exit,
        [Description("Ctrl-H")] Oracle_Help,
        [Description("Ctrl-DownArrow")] Oracle_Insert_Record,
        [Description("Ctrl-L")] Oracle_List_Of_Values,
        [Description("F2")] Oracle_List_Tab_Pages,
        [Description("Ctrl-U")] Oracle_Update_Record,

        [Description("Shift-PageDown")] Oracle_Next_Block,
        [Description("Tab")] Oracle_Next_Field,

        [Description("Shift-F7")] Oracle_Next_Primary_Key,
        [Description("DownArrow")] Oracle_Next_Record,
        [Description("Shift-F8")] Oracle_Next_Set_Of_Record,

        [Description("Shift-PageUp")] Oracle_Previous_Block,
        [Description("Shift-Tab")] Oracle_Previous_Field,
        [Description("Shift-UpArrow")] Oracle_Previous_Record,

        [Description("Ctrl-P")] Oracle_Print,
        [Description("PageDown")] Oracle_Scroll_Down,
        [Description("PageUp")] Oracle_Scroll_Up,
        [Description("UpArrow")] Oracle_Up,
        [Description("DownArrow")] Oracle_Down,

        [Description("Ctrl-K")] Oracle_Show_Keys,
    }
}
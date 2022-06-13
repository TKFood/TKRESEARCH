using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Reflection;
using System.Diagnostics;

namespace TKRESEARCH
{    public class shareArea
    {
        //--------------------------------------------
        // 以下之shareData宣告為靜態變數，隸屬類別層級，
        // 調用時可直接透過[類別.變數名稱]即可
        //--------------------------------------------
        public static string shareData;
        public static string UserName;
    }
}

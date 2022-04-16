using Northwind;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Management_System
{
    public static class ToolClass
    {
        /// <summary>
        /// 驗證檔案是否為圖檔
        /// </summary>
        /// <param name="filePath">完整檔案路徑</param>
        /// <returns></returns>
        public static bool checkImage(string filePath)
        {
            Dictionary<string, string> ImageTypes = new Dictionary<string, string>() { { "FFD8", ".jpg" }, { "424D", ".bmp" }, { "474946", ".gif" }, { "89504E470D0A1A0A", ".png" } };
            string builtHex = string.Empty;
            string extension = "";

            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                for (int i = 0; i < 8; i++)
                {
                    builtHex += fs.ReadByte().ToString("X2");
                    if (ImageTypes.ContainsKey(builtHex))
                    {
                        extension = ImageTypes[builtHex];
                        break;
                    }
                }
            }

            return !string.IsNullOrEmpty(extension);
        }
    }
}

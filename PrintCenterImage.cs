using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing.Printing;
using System.Windows.Forms;
using System.Drawing;

namespace Transport.Includes
{
    /// <summary>
    /// 使用方法如下：
    /// new PrintWorkItemLogoutImage(images.ToArray(), new PrintDocument(), new PrintDialog()).Print();
    /// 這樣就可以呼叫列印視窗列印了
    /// ※ 程式不用取得印表機設定列印份數，列印幾份的這功能元件本身就有，無須另外實作
    /// 列印出來的格式如下：
    /// 1. 一頁為一張圖片
    /// 2. 橫式
    /// 3. 圖片置中
    /// </summary>
    public class PrintCenterImage
    {
        private readonly PrintDocument _printDocument;
        private readonly PrintDialog _printDialog;
        private Image[] _images;

        //目前頁數
        private int _currentPage = 1;

        //開始頁數
        private int _fromPage = 1;

        //結束頁數
        private int _toPage = 1;

        //自動縮放
        private bool _autoSize = true;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="images">多個圖檔</param>
        /// <param name="printDocument">列印文件</param>
        /// <param name="printDialog">列印元件</param>
        public PrintCenterImage(Image[] images, PrintDocument printDocument, PrintDialog printDialog)
        {
            _images = images;
            _printDocument = printDocument;
            _printDialog = printDialog;

            //頁面設定
            _printDocument.DefaultPageSettings.Landscape = true;
            _printDocument.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);
            _printDocument.PrintPage += new PrintPageEventHandler(_printDocument_PrintPage);

            //列印元件設定
            _printDialog.Document = _printDocument;
            _printDialog.AllowCurrentPage = true;
            _printDialog.AllowSomePages = true;

            ResetPage();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="images">多個圖檔</param>
        /// <param name="printDocument">列印文件</param>
        /// <param name="printDialog">列印元件</param>
        /// <param name="autoSize">自動縮放大小(預設為true)</param>
        private PrintCenterImage(Image[] images, PrintDocument printDocument, PrintDialog printDialog, bool autoSize)
            : this(images, printDocument, printDialog)
        {
            _autoSize = autoSize;
        }

        /// <summary>
        /// 列印圖片，一張圖列印一頁
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void _printDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            //載入圖檔
            Image image = _images[_currentPage - 1];

            //縮放比例
            float scale = 1.0f;

            #region 1.計算縮放比例            
            
            if (_autoSize)
            {
                int widthCounter = e.MarginBounds.Width - image.Width;
                int heightCounter = e.MarginBounds.Height - image.Height;
                if (widthCounter < 0)
                {
                    float widthScale = Convert.ToSingle(e.MarginBounds.Width) / Convert.ToSingle(image.Width);
                    scale = widthScale;
                }

                if (heightCounter < 0)
                {
                    float heightScale = Convert.ToSingle(e.MarginBounds.Height) / Convert.ToSingle(image.Height);
                    scale = heightScale > scale ? heightScale : scale;
                }
            }
            
            #endregion

            #region 2.設定左上角放置點

            float width = image.Width * scale;
            float height = image.Height * scale;
            float left = ((e.MarginBounds.Width - width) / 2) < 0.0f ? 0.0f : ((e.MarginBounds.Width - width) / 2);
            float top = ((e.MarginBounds.Height - height) / 2) < 0.0f ? 0.0f : ((e.MarginBounds.Height - height) / 2);

            //建立矩形結構，並設定放置的左上角座標和矩形結構的大小
            RectangleF rect = RectangleF.Empty;
            rect = new RectangleF(left, top, width, height);

            #endregion

            //列印圖形
            e.Graphics.DrawImage(image, rect);

            //以下是列印多頁的程式碼
            e.HasMorePages = (_currentPage + 1) <= _toPage;

            //這是指該份文件如果有多頁，則該屬性要設為true
            if (!e.HasMorePages)
                return;

            _currentPage++;
        }

        /// <summary>
        /// 重設頁碼參數
        /// </summary>
        private void ResetPage()
        {
            _currentPage = 1;
            _fromPage = 1;
            _toPage = _images.Length;
        }

        /// <summary>
        /// 設定列印的頁數範圍
        /// </summary>
        /// <param name="printDialog1"></param>
        /// <returns></returns>
        private bool PrintSetting(PrintDialog printDialog1)
        {
            PrintRange printRange = printDialog1.PrinterSettings.PrintRange;
            bool result;
            switch (printRange)
            {
                case PrintRange.CurrentPage:
                    _fromPage = _currentPage;
                    _toPage = _currentPage;
                    result = true;
                    break;

                case PrintRange.AllPages:
                    ResetPage();
                    result = true;
                    break;

                case PrintRange.SomePages:
                    _fromPage = printDialog1.PrinterSettings.FromPage;
                    _toPage = printDialog1.PrinterSettings.ToPage;
                    _currentPage = _fromPage;
                    result = _fromPage > 0 && _toPage <= _images.Length;
                    break;
                default:
                    ResetPage();
                    result = true;
                    break;
            }

            return result;
        }

        /// <summary>
        /// 列印文件
        /// </summary>
        public void Print()
        {
            DialogResult result = _printDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                if (!PrintSetting(_printDialog))
                {
                    FieldChk.ShowError("頁碼範圍設定錯誤，請重新設定", "列印");
                    return;
                }

                //列印
                _printDocument.Print();
            }
        }
    }
}

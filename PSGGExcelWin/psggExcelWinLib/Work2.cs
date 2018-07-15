using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace psggExcelWinLib
{
    #region pic
    public partial class Work
    {
        public class Pic
        {
            public int row;
            public int col;

            public Bitmap bmp;  

            public bool   modified;
        }

        List<Pic> m_pic_list;

        public void ReadPics()
        {
            //System.Diagnostics.Debugger.Break();

            if (m_pic_list==null) m_pic_list = new List<Pic>();
            m_pic_list.Clear();

            var sheet = m_ew.GetSheet();
            foreach(var i in sheet.Shapes)
            {
                var s = (Excel.Shape)i;
                if (s.Type ==  Microsoft.Office.Core.MsoShapeType.msoPicture)
                {
                    var item = new Pic();
                    item.col = s.TopLeftCell.Column;
                    item.row = s.TopLeftCell.Row;

                    s.CopyPicture(Excel.XlPictureAppearance.xlScreen,Excel.XlCopyPictureFormat.xlBitmap);
                    item.bmp = (Bitmap)Clipboard.GetData(DataFormats.Bitmap);
                    if (item.bmp!=null)
                    {
                        item.bmp.MakeTransparent();
                        m_pic_list.Add(item);
                    }
                }

                Marshal.ReleaseComObject(s);
                s = null;
            }
        }

        public void DisposePics()
        {
            if (m_pic_list!=null)
            {
                m_pic_list.ForEach(i=> {
                    if (i.bmp!=null)
                    {
                        try {
                            i.bmp.Dispose();
                        } catch { }
                    }
                    i.bmp = null;
                });
            }
            m_pic_list = null;
        }

        public Bitmap GetBmp(int row, int col)
        {
            var find = m_pic_list.Find(i=>i.row == row && i.col == col);
            if (find!=null)
            {
                return find.bmp;
            }
            return null;
        }

        /// <summary>
        /// bmp = nullで削除
        /// </summary>
        public bool SetBmp(int  row, int col, Bitmap bmp)
        {
            var item = m_pic_list.Find(i=>i.row == row && i.col == col);
            if (item==null)
            {
                item = new Pic();
                item.row = row;
                item.col = col;
                m_pic_list.Add(item);
            }

            item.modified = true;

            if (bmp == null)
            {
                if (item.bmp != null)
                {
                    item.bmp.Dispose();
                    item.bmp = null;
                }
                return true;
            }

            item.bmp = (Bitmap)bmp.Clone();
            return true;
        }

        public bool UpdateBmps()
        {
            if (m_pic_list == null) return false;

            try {

                var sheet = m_ew.GetSheet();

                //変更があったのを削除
                m_pic_list.ForEach(i=> {
                    if (!i.modified) return;

                    foreach(var j in sheet.Shapes)
                    {
                        var s = (Excel.Shape)j;
                        if (
                            s.TopLeftCell.Column == i.col
                            &&
                            s.TopLeftCell.Row == i.row
                            )
                        {
                            s.Delete();
                        }
                    }
                });

                //変更されたのを追加
                var shapes = sheet.Shapes;
                var tempbmp = Path.Combine( Path.GetTempPath(),"temp.png");
                //https://tonari-it.com/excel-vba-shapes-addpicture/
                //https://dobon.net/vb/dotnet/graphics/saveimage.html
                //https://qiita.com/shela/items/d20fd84a82d930b5804e
                m_pic_list.ForEach(i=> {
                    if (!i.modified)   return;
                    if (i.bmp == null) return; 
                    //一度ファイルに
                    if (File.Exists(tempbmp)) File.Delete(tempbmp);
                    i.bmp.MakeTransparent();
                    i.bmp.Save(tempbmp,ImageFormat.Png);
                
                    var cell = (Excel.Range)sheet.Cells[i.row,i.col];
                    var x = (float)cell.Offset.Left;
                    var y = (float)cell.Offset.Top;
                    Marshal.ReleaseComObject(cell);

                    var w = (float)i.bmp.Width  * (72f / 96f);
                    var h = (float)i.bmp.Height * (72f / 96f);

                    sheet.Shapes.AddPicture(tempbmp,Office.MsoTriState.msoFalse,Office.MsoTriState.msoCTrue,x,y,w,h);
                });   

                //modifiedをfalseへ
                m_pic_list.ForEach(i=> {
                    i.modified = false;
                });

                return true;
            }
            catch (SystemException e)
            {
                latest_error = e.Message;
                return false;
            }
        }


    }
    #endregion
}

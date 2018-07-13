using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace psggExcelWinLib
{
    #region cell
    public partial class Work
    {
        public class Cell
        {
            public int row;
            public int col;

            public string text;
        }

        public string latest_error;

        ExcelWork  m_ew;

        List<Cell> m_cell_list;

        public bool Load(string filename)
        {
            if (m_ew!=null) {
                latest_error = "Unexpected! {3FB775C6-5A13-4AE7-B2CD-54CE5358C78E}";
                throw new SystemException(latest_error);
            }
            if (!File.Exists(filename))
            {
                latest_error = "File not found : " + filename;
                return false;
            }
            try {
                m_ew = new ExcelWork();
                m_ew.Load(filename);
                m_cell_list =new List<Cell>();
                return true;
            } catch (SystemException e) {
                latest_error = e.Message;
                return false;
            }
        }
        public bool Save()
        {
            if (m_ew==null) return false;
            try {
                m_ew.Save();
            } catch (SystemException e) {
                latest_error = e.Message;
                return false;
            }
            return true;
        }

        public bool SetSheet(string sheetname)
        {
            DisposePics();
            m_cell_list.Clear();

            m_ew.SetSheet(sheetname);
            ReadSheet();
            return m_ew.GetSheet() != null;
        }

        public void NewSheet(string sheetname)
        {
            DisposePics();
            m_cell_list.Clear();

            m_ew.NewSheet(sheetname);
        }

        public void NewSheetForce(string sheetname)
        {
            if (SetSheet(sheetname))
            {
                clear_all();
            }
            else
            {
                NewSheet(sheetname);
            }
        }

        public bool ReadSheet()
        {
            var sheet = m_ew.GetSheet();
            if (sheet == null)
            {
                return false;
            }
            var range = sheet.UsedRange;

            object[,] objs = (object[,])range.Value2;

            var row_start = range.Row;
            var row_len   = range.Rows.Count;
            var col_start = range.Column;
            var col_len   = range.Columns.Count;

            m_cell_list.Clear();
            for(var r = 1; r<=row_len; r++)
            {
                for(var c = 1; c<=col_len; c++)
                {
                    var o = objs[r,c];

                    var cell = new Cell();
                    cell.row = row_start + (r-1);
                    cell.col = col_start + (c-1);
                    
                    cell.text = o!=null ? o.ToString() : null;

                    m_cell_list.Add(cell);
                }
            }
            Marshal.ReleaseComObject(range);
            range = null;

            return true;
        }

        public bool WriteSheet()
        {
            var sheet = m_ew.GetSheet();
            if (sheet == null)
            {
                return false;
            }
            var max_col = 0;
            var max_row = 0;
            if (get_cell_list_max(out max_row, out max_col))
            {
                var range = (Excel.Range)sheet.Range[sheet.Cells[1,1],sheet.Cells[max_row,max_col]];

                object[,] objs = (object[,])range.Value2;
                
                m_cell_list.ForEach(i=> {
                    objs[i.row,i.col] = i.text;
                });

                range.NumberFormatLocal = "@"; //文字列
                range.Value2 = objs;

                Marshal.ReleaseComObject(range);
                return true;
            }
            return false;
        }
        public int MaxRow()
        {
            var row = 0;
            var col = 0;
            get_cell_list_max(out row,out col);
            return row;
        }
        public int MaxCol()
        {
            var row = 0;
            var col = 0;
            get_cell_list_max(out row,out col);
            return col;
        }
        public string GetStr(int row, int col)
        {
            var find = m_cell_list.Find(i=>(i.row==row && i.col==col));
            return find!=null ? find.text : null;
        }

        public void SetCell(int row, int col, string text)
        {
            var find = m_cell_list.Find(i=>(i.row==row && i.col==col));
            if (find!=null)
            {
                find.text = text;
            }
            else
            {
                var cell = new Cell();
                cell.col = col;
                cell.row = row;
                cell.text = text;
                m_cell_list.Add(cell);
            }
        }
        
        public void Dispose()
        {
            DisposePics();

            if (m_ew!=null)
            {
                m_ew.Dispose();
            }
            m_ew = null;
        }

        // --- tools for this class
        private bool get_cell_list_max(out int row, out int col)
        {
            var max_col = -1;
            int max_row = -1;
            m_cell_list.ForEach(i=> {
                max_col = Math.Max(i.col,max_col);
                max_row = Math.Max(i.row,max_row);
            });
            col = max_col;
            row = max_row;

            return (col >= 0 && row >=0);
        }
        private void clear_all()
        {
            m_cell_list.Clear();

            var sheet = m_ew.GetSheet();
            if (sheet == null) return;

            var range = sheet.UsedRange;
            range.Value2 = null;
            Marshal.ReleaseComObject(range);
        }
    }
    #endregion
}

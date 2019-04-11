using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.IO.Ports;

namespace OpcTest
{
    public static class Extension
    {
        #region Label

        public static void SetText(this Label lbl, string text)
        {
            lbl.Invoke(new MethodInvoker(() => { lbl.Text = text; }));
        }

        public static void SetText(this Label lbl, string text, Color backColor)
        {
            lbl.Invoke(new MethodInvoker(() => { lbl.Text = text; lbl.BackColor = backColor; }));
        }



        #endregion

        #region TextBox

        public static void SetText(this TextBox tb, string text)
        {
            tb.Invoke(new MethodInvoker(() => { tb.Text = text; }));
        }

        public static void AppendText(this TextBox tb, string text)
        {
            tb.Invoke(new MethodInvoker(() => { tb.Text += text; }));
        }
        public static void SetFocus(this TextBox tb)
        {
            tb.Invoke(new MethodInvoker(() => { tb.Focus (); }));
        }
       

        #endregion

        #region DataGridView

        public static void AddRow(this DataGridView grid, params object[] row)
        {
            grid.Invoke(new MethodInvoker(() => { grid.Rows.Add(row); }));
        }

        public static void InsertRow(this DataGridView grid, int index, params object[] row)
        {
            grid.Invoke(new MethodInvoker(() => { grid.Rows.Insert(index, row); }));
        }

        public static void ClearRows(this DataGridView grid)
        {
            grid.Invoke(new MethodInvoker(() => { grid.Rows.Clear(); }));
        }

        #endregion

        #region FlowLayoutPanel

        public static void Enable(this FlowLayoutPanel panel)
        {
            panel.Invoke(new MethodInvoker(() => { panel.Enabled = true; }));
        }

        public static void Disable(this FlowLayoutPanel panel)
        {
            panel.Invoke(new MethodInvoker(() => { panel.Enabled = false; }));
        }

        #endregion
        #region TextBox

        public static void SetText(this ComboBox  cb, string text)
        {
            cb.Invoke(new MethodInvoker(() => { cb.Text = text; }));
        }

        public static void AddItem(this ComboBox cb, string text)
        {
            cb.Invoke(new MethodInvoker(() => { cb.AddItem(text); }));
        }

        #endregion
        #region TextBox

        //public static void SerialPort(this SerialPort cb, string text)
        //{
        //    cb.Invoke(new MethodInvoker(() => { cb.Text = text; }));
        //}

        //public static void AddItem(this ComboBox cb, string text)
        //{
        //    cb.Invoke(new MethodInvoker(() => { cb.AddItem(text); }));
        //}

        #endregion
        #region SereialPort
        public static void SetName(this SerialPort lbl, string text)
        {
            //lbl.Invoke(new MethodInvoker(() => { lbl.na = text; }));
        }

        #endregion
    }
}

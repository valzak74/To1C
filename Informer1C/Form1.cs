using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.Entity;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.Utils;

namespace Informer1C
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private List<string> Производители = new List<string> { 
                    "     8S  ",
                    "     GS  ",
                    "    5UF  "
                };

        private void ShowNormalForm()
        {
            notifyIcon1.Icon = Properties.Resources.circle_green;
            timer1.Stop();
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.BringToFront();
        }

        private void HideMinimizedForm()
        {
            this.WindowState = FormWindowState.Minimized;
            this.Hide();
            DbOperations.FreeDbContext();
            timer1.Start();
        }

        private void showToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowNormalForm();
        }

        private void ActivateDbData()
        {
            if (DbOperations.DbContextIsNull())
            {
                gridControl1.DataSource = DbOperations.GetData(Производители).ToBindingList();
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Form1_Move(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
                HideMinimizedForm();
            else
                ActivateDbData();
        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            ShowNormalForm();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Stop();
            List<НоваяЗаявка> НовыеЗаявки = DbOperations.GetFreshData(Производители);
            if (НовыеЗаявки != null)
            {
                if (НовыеЗаявки.Count > 0)
                    notifyIcon1.Icon = Properties.Resources.circle_orange;
                foreach (НоваяЗаявка nz in НовыеЗаявки)
                {
                    notifyIcon1.ShowBalloonTip(3000, "Новая заявка " + nz.НомерДокПолный, "Контрагент " + nz.Контрагент, ToolTipIcon.Info);
                }
            }
            DbOperations.FreeDbContext();
            timer1.Start();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Interval = 60000;
            notifyIcon1.Icon = Properties.Resources.circle_green;
        }

        private void gridView1_MasterRowGetRelationName(object sender, DevExpress.XtraGrid.Views.Grid.MasterRowGetRelationNameEventArgs e)
        {
            e.RelationName = "СтрокиДокумента";
        }

        private void gridView1_MasterRowGetChildList(object sender, DevExpress.XtraGrid.Views.Grid.MasterRowGetChildListEventArgs e)
        {
            Заявка c = (Заявка)gridView1.GetRow(e.RowHandle);
            e.ChildList = new BindingSource(c, "СтрокиДокумента");
        }

        private void gridView1_MasterRowEmpty(object sender, DevExpress.XtraGrid.Views.Grid.MasterRowEmptyEventArgs e)
        {
            Заявка c = (Заявка)gridView1.GetRow(e.RowHandle);
            e.IsEmpty = c.IDDOC == null;
        }

        private void gridView1_MasterRowGetRelationCount(object sender, DevExpress.XtraGrid.Views.Grid.MasterRowGetRelationCountEventArgs e)
        {
            e.RelationCount = 1;
        }

        private void gridControl1_ViewRegistered(object sender, DevExpress.XtraGrid.ViewOperationEventArgs e)
        {
            DevExpress.XtraGrid.Views.Grid.GridView detailCloneView = e.View as DevExpress.XtraGrid.Views.Grid.GridView;
            if (detailCloneView.LevelName == "СтрокиДокумента")
                detailCloneView.OptionsEditForm.CustomEditFormLayout = new AdvancedEditForm();
        }

        private void gridView1_DoubleClick(object sender, EventArgs e)
        {
            DXMouseEventArgs ea = e as DXMouseEventArgs;
            GridView view = sender as GridView;
            GridHitInfo info = view.CalcHitInfo(ea.Location);
            if (info.InRow || info.InRowCell)
            {
                view.SetMasterRowExpanded(info.RowHandle, !view.GetMasterRowExpanded(info.RowHandle));
            } 
        }

        private void gridView2_RowUpdated(object sender, DevExpress.XtraGrid.Views.Base.RowObjectEventArgs e)
        {
            btnOk.Enabled = DbOperations.DataChanged((BindingList<Заявка>)gridControl1.DataSource);
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            DbOperations.CreateDocument((BindingList<Заявка>)gridControl1.DataSource);
            HideMinimizedForm();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            HideMinimizedForm();
        }

        private void Form1_Deactivate(object sender, EventArgs e)
        {
            HideMinimizedForm();
        }

    }
}

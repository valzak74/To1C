using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DevExpress.Skins;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraEditors.Drawing;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraGrid.Drawing;
using DevExpress.Utils.Drawing;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Columns;
using DevExpress.Utils;
using DevExpress.XtraGrid.Views.Base;
using System.Drawing;
using System.Windows.Forms;

namespace Informer1C
{
    public class GridFooterExtender
    {
        GridView view;
        SkinEditorButtonPainter customButtonPainter;
        EditorButtonObjectInfoArgs args;
        Size buttonSize;

        public GridFooterExtender(GridView view)
        {
            this.view = view;
            this.buttonSize = new Size(14, 14);
        }

        public void AddCustomButton()
        {
            CreateButtonPainter();
            CreateButtonInfoArgs();
            SubscribeToEvents();
        }

        private void CreateButtonPainter()
        {
            customButtonPainter = new SkinEditorButtonPainter(DevExpress.LookAndFeel.UserLookAndFeel.Default.ActiveLookAndFeel);
        }

        private void CreateButtonInfoArgs()
        {
            EditorButton btn = new EditorButton(ButtonPredefines.Glyph);
            args = new EditorButtonObjectInfoArgs(btn, new DevExpress.Utils.AppearanceObject());
        }

        private void SubscribeToEvents()
        {
            view.CustomDrawFooter += OnCustomDrawFooter;
            //view.MouseDown += OnMouseDown;
            //view.MouseUp += OnMouseUp;
            //view.MouseMove += OnMouseMove;
        }

        void OnCustomDrawFooter(object sender, RowObjectCustomDrawEventArgs e)
        {
            //SetUpButtonInfoArgs(e);
            //if (e.Column == null) return;
            DefaultDrawColumnHeader(e);
            DrawCustomButton(e);
            e.Handled = true;
        }

        private void DrawCustomButton(RowObjectCustomDrawEventArgs e)
        {
            SetUpButtonInfoArgs(e);
            //customButtonPainter.DrawObject(args);
            customButtonPainter.DrawObject(args);
            StringFormat f = new StringFormat();
            //calc text bounds here  
            Rectangle textBounds = args.Bounds;
            textBounds.Inflate(-3, -3);
            int imageWidth = 12;
            int imageHeight = 12;
            //customButtonPainter.DrawCaption(args, "Your Text", e.Appearance.Font, e.Appearance.GetForeBrush(args), textBounds, f);
            //e.Graphics.DrawImage(new Bitmap(@"C:\Users\tochilkin.alexander\Desktop\IMAGES\small-bell_318-10933.jpg"), new Rectangle(args.Bounds.Right - imageWidth - 5, args.Bounds.Y + 3, imageWidth, imageHeight));
            //customButtonPainter.DrawElementInfoBitmap(args);
            //customButtonPainter.DrawElementIntoBitmap(args, ObjectState.Normal);
        }

        private void SetUpButtonInfoArgs(RowObjectCustomDrawEventArgs e)
        {
            args.Cache = e.Cache;
            args.Bounds = CalcButtonRect(e.Info, e.Cache.Graphics);
            ObjectState state = ObjectState.Normal;
            //if (e.Column.Tag is ObjectState)
            //    state = (ObjectState)e.Column.Tag;
            args.State = state;
        }

        private static void DefaultDrawColumnHeader(RowObjectCustomDrawEventArgs e)
        {
            e.Painter.DrawObject(e.Info);
        }

        private Rectangle CalcButtonRect(ObjectInfoArgs columnArgs, Graphics gr)
        {
            Rectangle columnRect = columnArgs.Bounds;
            int innerElementsWidth = 20; //CalcInnerElementsMinWidth(columnArgs, gr);
            Rectangle buttonRect = new Rectangle(columnRect.Right - innerElementsWidth - buttonSize.Width - 2,
                columnRect.Y + columnRect.Height / 2 - buttonSize.Height / 2, buttonSize.Width, buttonSize.Height);
            return buttonRect;
        }

        //private int CalcInnerElementsMinWidth(CustomDrawObjectEventArgs columnArgs, Graphics gr)
        //{
        //    bool canDrawMode = true;
        //    return columnArgs.InnerElements.CalcMinSize(gr, ref canDrawMode).Width;
        //}

    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

// This is the code for your desktop app.
// Press Ctrl+F5 (or go to Debug > Start Without Debugging) to run your app.

namespace Excel_Dashboard
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
            this.Initialize();
        }

        #region Initialize

        private void Initialize()
        {
            this.Width = Screen.PrimaryScreen.Bounds.Width;
            this.Height = Screen.PrimaryScreen.Bounds.Height;
            this.Panel.Height = Screen.PrimaryScreen.Bounds.Height;
            this.Panel.Width = Screen.PrimaryScreen.Bounds.Width;

            FlowLayoutPanel flowItemHeader = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Height = 40,
                Width = Panel.Width
            };

            int padingSpace = (int) Panel.Width / Utils.COL_NUMBERS;
            Label header_1 = new Label
            {
                Text = $"FOLIO",
                Width = padingSpace - 10,
                Height = Utils.ROW_HEIGHT,
                Font = Utils.HEADER_FONT,
                ForeColor = Utils.HEADER_COLOR,
                TextAlign = ContentAlignment.MiddleCenter
        };

            Label header_2 = new Label
            {
                Text = $"TICKET",
                Width = padingSpace - 10,
                Height = Utils.ROW_HEIGHT,
                Font = Utils.HEADER_FONT,
                ForeColor = Utils.HEADER_COLOR,
                TextAlign = ContentAlignment.MiddleCenter
            };

            Label header_3 = new Label
            {
                Text = $"NOMBRE DEL CLIENTE",
                Width = padingSpace - 10,
                Height = Utils.ROW_HEIGHT,
                Font = Utils.HEADER_FONT,
                ForeColor = Utils.HEADER_COLOR,
                TextAlign = ContentAlignment.MiddleCenter
            };

            Label header_4 = new Label
            {
                Text = $"ZONE",
                Width = padingSpace - 10,
                Height = Utils.ROW_HEIGHT,
                Font = Utils.HEADER_FONT,
                ForeColor = Utils.HEADER_COLOR,
                TextAlign = ContentAlignment.MiddleCenter
            };

            Label header_5 = new Label
            {
                Text = $"UNIDAD",
                Width = padingSpace - 10,
                Height = Utils.ROW_HEIGHT,
                Font = Utils.HEADER_FONT,
                ForeColor = Utils.HEADER_COLOR,
                TextAlign = ContentAlignment.MiddleCenter
            };

            Label header_6 = new Label
            {
                Text = $"CHOFER",
                Width = padingSpace - 10,
                Height = Utils.ROW_HEIGHT,
                Font = Utils.HEADER_FONT,
                ForeColor = Utils.HEADER_COLOR,
                TextAlign = ContentAlignment.MiddleCenter
            };

            Label header_7 = new Label
            {
                Text = $"H/SALIDA",
                Width = padingSpace - 10,
                Height = Utils.ROW_HEIGHT,
                Font = Utils.HEADER_FONT,
                ForeColor = Utils.HEADER_COLOR,
                TextAlign = ContentAlignment.MiddleCenter
            };

            Label header_8 = new Label
            {
                Text = $"ESTATUS",
                Width = padingSpace - 10,
                Height = Utils.ROW_HEIGHT,
                Font = Utils.HEADER_FONT,
                ForeColor = Utils.HEADER_COLOR,
                TextAlign = ContentAlignment.MiddleCenter
            };

            flowItemHeader.Controls.Add(header_1);
            flowItemHeader.Controls.Add(header_2);
            flowItemHeader.Controls.Add(header_3);
            flowItemHeader.Controls.Add(header_4);
            flowItemHeader.Controls.Add(header_5);
            flowItemHeader.Controls.Add(header_6);
            flowItemHeader.Controls.Add(header_7);
            flowItemHeader.Controls.Add(header_8);

            this.Panel.Controls.Add(flowItemHeader);
            int col_numbers = Utils.COL_NUMBERS;

            for (int index = 0; index < Utils.COL_NUMBERS; index++)
            {
                FlowLayoutPanel flowItem = new FlowLayoutPanel
                {
                    FlowDirection = FlowDirection.LeftToRight,
                    Height = Utils.ROW_HEIGHT,
                    Width = Panel.Width
                };

                for (int i = 0; i < Utils.COL_NUMBERS; i++)
                {
                    Label item_1 = new Label
                    {
                        Text = $"{i}",
                        Width = padingSpace - 10,
                        Height = Utils.ROW_HEIGHT,
                        Font = Utils.CONTENT_FONT,
                        ForeColor = Utils.CONTENT_COLOR,
                        TextAlign = ContentAlignment.MiddleCenter
                    };

                    flowItem.Controls.Add(item_1);
                }

                this.Panel.Controls.Add(flowItem);
            }
        }

       #endregion
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;


namespace Excel_Dashboard
{
    public partial class Main : Form
    {
        private bool isAlreadyInitialize = false;
        private FileSystemWatcher fileSystemWatcher;
        public Main()
        {
            InitializeComponent();
            this.Initialize();
        }

        #region Initialize

        private void Initialize()
        {
            this.Width = Screen.PrimaryScreen.Bounds.Width;
            this.Height = Screen.PrimaryScreen.Bounds.Height-30;

            this.PanelHeader.Width = this.Width - 30;

            //this.Panel.Height = this.Height - 35;
            this.Panel.Height = this.Height - 70;
            this.Panel.Width = this.Width - 30;//25
            //this.WindowState = FormWindowState.Maximized;
            //this.Panel.AutoScroll = false;
            //this.Panel.HorizontalScroll.Enabled = false;
            this.Panel.AutoScroll = true;

            Loading loadWindow = new Loading();
            loadWindow.Show();
            SetInformation();
            loadWindow.Close();

            fileSystemWatcher = new FileSystemWatcher();
            fileSystemWatcher.Changed += FileSystemWatcher_Changed;
            fileSystemWatcher.Path = Utils.CURRENT_PATH;
            fileSystemWatcher.NotifyFilter = NotifyFilters.LastWrite;
            fileSystemWatcher.IncludeSubdirectories = false;
            //fileSystemWatcher.Filter= "*.xlsx*";

            // You must add this line - this allows events to fire.
            fileSystemWatcher.EnableRaisingEvents = true;
            isAlreadyInitialize = true;
        }

        #endregion

        private void FileSystemWatcher_Changed(object sender, FileSystemEventArgs e)
        {
            /*string fileNameChange = e.Name;
            if (fileNameChange.IndexOf(".xlsx")!=-1)
            {
                SetInformation();
            }*/
            try
            {
                fileSystemWatcher.EnableRaisingEvents = false;
                string fileNameChange = e.Name;
                SetInformation();
            }
            finally
            {
                fileSystemWatcher.EnableRaisingEvents = true;
            }
            
        }


        private void SetInformation()
        {
            DirectoryInfo d = new DirectoryInfo(Utils.CURRENT_PATH);
            FileInfo[] Files = d.GetFiles("*.xlsx");
            List<string> excelNames = new List<string>();
            foreach (FileInfo file in Files)
            {
                if (file.Name.IndexOf("~$") == -1)
                {
                    excelNames.Add(file.Name);
                }
            }
            List<List<Column>> tempLists = new List<List<Column>>();
            foreach (string excelName in excelNames)
            {
                string filePath = Utils.CURRENT_PATH;
                string strRandom = Utils.RandomString(9);
                string copyFilePath = $@"{filePath}\temp\{strRandom}{excelName}";
                string tempFolder = $@"{filePath}\temp";
                Utils.CreateFolder(tempFolder);
                filePath = filePath + '\\' + excelName;
                File.Copy(filePath, copyFilePath);
                tempLists.Add(ExcelUtil.GetData(copyFilePath));
            }

            List<Column> data = new List<Column>();
            if (tempLists.Count == 0) return;
            foreach (List<Column> list in tempLists)
            {
                foreach (Column item in list)
                {
                    data.Add(item);
                }
            }
            if (data.Count == 0) return;

            if (isAlreadyInitialize)
            {
                this.Panel.Invoke(new MethodInvoker(delegate
                {
                    Panel.Controls.Clear();
                }));
            }

            if (data == null)
            {
                MessageBox.Show("Error al intentar leer el excel", "Error");
                this.Close();
                return;
            }

            FlowLayoutPanel flowItemHeader = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Height = 40,
                Width = Panel.Width
            };

            int padingSpace = (int)(Panel.Width-9) / Utils.COL_NUMBERS;
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

            //flowItemHeader.Controls.Add(header_1);
            flowItemHeader.Controls.Add(header_2);
            flowItemHeader.Controls.Add(header_3);
            flowItemHeader.Controls.Add(header_4);
            flowItemHeader.Controls.Add(header_5);
            flowItemHeader.Controls.Add(header_6);
            flowItemHeader.Controls.Add(header_7);
            flowItemHeader.Controls.Add(header_8);

            
            if (!isAlreadyInitialize)
            {
                this.PanelHeader.Controls.Add(flowItemHeader);
            }

            int col_numbers = Utils.COL_NUMBERS;

            foreach (Column item in data)
            {
                FlowLayoutPanel flowItem = new FlowLayoutPanel
                {
                    FlowDirection = FlowDirection.LeftToRight,
                    Height = Utils.ROW_HEIGHT,
                    Width = Panel.Width
                };
                Label item_1 = new Label
                {
                    Text = $"{item.Folio}",
                    Width = padingSpace - 10,
                    Height = Utils.ROW_HEIGHT,
                    Font = Utils.CONTENT_FONT,
                    ForeColor = Utils.CONTENT_COLOR,
                    TextAlign = ContentAlignment.MiddleCenter
                };

                Label item_2 = new Label
                {
                    Text = $"{item.Ticket}",
                    Width = padingSpace - 10,
                    Height = Utils.ROW_HEIGHT,
                    Font = Utils.CONTENT_FONT,
                    ForeColor = Utils.CONTENT_COLOR,
                    TextAlign = ContentAlignment.MiddleCenter
                };

                Label item_3 = new Label
                {
                    Text = $"{item.NombreCliente}",
                    Width = padingSpace - 10,
                    Height = Utils.ROW_HEIGHT,
                    Font = Utils.CONTENT_FONT,
                    ForeColor = Utils.CONTENT_COLOR,
                    TextAlign = ContentAlignment.MiddleCenter
                };

                Label item_4 = new Label
                {
                    Text = $"{item.Zona}",
                    Width = padingSpace - 10,
                    Height = Utils.ROW_HEIGHT,
                    Font = Utils.CONTENT_FONT,
                    ForeColor = Utils.CONTENT_COLOR,
                    TextAlign = ContentAlignment.MiddleCenter
                };

                Label item_5 = new Label
                {
                    Text = $"{item.Unidad}",
                    Width = padingSpace - 10,
                    Height = Utils.ROW_HEIGHT,
                    Font = Utils.CONTENT_FONT,
                    ForeColor = Utils.CONTENT_COLOR,
                    TextAlign = ContentAlignment.MiddleCenter
                };

                Label item_6 = new Label
                {
                    Text = $"{item.Chofer}",
                    Width = padingSpace - 10,
                    Height = Utils.ROW_HEIGHT,
                    Font = Utils.CONTENT_FONT,
                    ForeColor = Utils.CONTENT_COLOR,
                    TextAlign = ContentAlignment.MiddleCenter
                };

                Label item_7 = new Label
                {
                    Text = $"{item.Salida}",
                    Width = padingSpace - 10,
                    Height = Utils.ROW_HEIGHT,
                    Font = Utils.CONTENT_FONT,
                    ForeColor = Utils.CONTENT_COLOR,
                    TextAlign = ContentAlignment.MiddleCenter
                };
                string status = !Utils.IsEmptyString(item.EstatusEntregado) ? "Entregado" : (!Utils.IsEmptyString(item.EstatusTrayecto) ? "En Trayecto" : (!Utils.IsEmptyString(item.EstatusCargando) ? "Cargando" : ""));
                if (!Utils.IsEmptyString(item.Estatus)) status = item.Estatus;
                Label item_8 = new Label
                {
                    Text = $"{status}",
                    Width = padingSpace - 10,
                    Height = Utils.ROW_HEIGHT,
                    Font = Utils.CONTENT_FONT,
                    ForeColor = Utils.CONTENT_COLOR,
                    TextAlign = ContentAlignment.MiddleCenter
                };

                //flowItem.Controls.Add(item_1);
                flowItem.Controls.Add(item_2);
                flowItem.Controls.Add(item_3);
                flowItem.Controls.Add(item_4);
                flowItem.Controls.Add(item_5);
                flowItem.Controls.Add(item_6);
                flowItem.Controls.Add(item_7);
                flowItem.Controls.Add(item_8);

                if (isAlreadyInitialize)
                {
                    this.Panel.Invoke(new MethodInvoker(delegate
                    {
                        this.Panel.Controls.Add(flowItem);
                    }));
                }
                else
                {
                    this.Panel.Controls.Add(flowItem);
                }

            }
        }
    }
}

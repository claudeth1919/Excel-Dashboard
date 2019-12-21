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
        private int padingSpace;
        private DateTime lastFileUpdateDate = DateTime.Today;
        private List<Column> data = new List<Column>();
        private System.Timers.Timer timer = new System.Timers.Timer();
        private bool isSubiendo = false;
        //private int bandera = 0;
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
            padingSpace = (int)(Panel.Width - 9) / Utils.COL_NUMBERS;

            Loading loadWindow = new Loading();
            loadWindow.Show();
            data = GetData();
            SetInformation(data);
            loadWindow.Close();

            fileSystemWatcher = new FileSystemWatcher();
            fileSystemWatcher.Path = Utils.CURRENT_PATH;
            fileSystemWatcher.Changed += FileSystemWatcher_Changed;
            fileSystemWatcher.NotifyFilter = NotifyFilters.LastWrite;
            fileSystemWatcher.IncludeSubdirectories = false;
            //fileSystemWatcher.Filter= "*.xlsx*";

            // You must add this line - this allows events to fire.
            fileSystemWatcher.EnableRaisingEvents = true;
            isAlreadyInitialize = true;
            //Timer begin
            timer.Interval = 370;

            timer.Elapsed += OnTimedEvent;
            timer.AutoReset = true;
            timer.Enabled = true;
            //Timer end
        }

        #endregion

        private void FileSystemWatcher_Changed(object sender, FileSystemEventArgs e)
        {
            string changedFile = GetLastUpdatedFile();
            try
            {
                //fileSystemWatcher.EnableRaisingEvents = false;
                if (changedFile != String.Empty)
                {
                    //System.Threading.Thread.Sleep(3000);
                    List<Column> updatedData = GetData(changedFile);
                    this.data.RemoveAll(x => x.ExcelOrigen == changedFile);
                    foreach (Column col in updatedData)
                    {
                        this.data.Add(col);
                    }
                    SetInformation(this.data);
                }
            }
            finally
            {
                fileSystemWatcher.EnableRaisingEvents = true;
            }
            
        }

        private void OnTimedEvent(Object source, System.Timers.ElapsedEventArgs e)
        {
            int distancia = 1;
            int locationY = 0;
            if (!isSubiendo)
            {
                locationY += distancia;
            }
            else
            {
                locationY -= distancia;
            }
            this.Panel.Invoke(new MethodInvoker(delegate
            {
                try
                {
                    this.Panel.VerticalScroll.Value += locationY;
                }
                catch
                {

                }
            }));
            int diffMax = (this.Panel.VerticalScroll.Maximum - this.Panel.VerticalScroll.LargeChange + 1);
            if (this.Panel.VerticalScroll.Minimum == this.Panel.VerticalScroll.Value) isSubiendo = false;
            else if (diffMax == (this.Panel.VerticalScroll.Value)) isSubiendo = true;

        }

        private string GetLastUpdatedFile()
        {
            FileInfo lastUpdatedFile = null;
            DirectoryInfo d = new DirectoryInfo(Utils.CURRENT_PATH);
            FileInfo[] Files = d.GetFiles("*.xlsx");
            lastUpdatedFile = Files.OrderByDescending(x => x.LastWriteTime).Where(x=> x.Name.IndexOf("~$") == -1).FirstOrDefault();
            if (lastUpdatedFile == null) return String.Empty;
            if (DateTime.Compare(lastUpdatedFile.LastWriteTime,lastFileUpdateDate) > 0 /*&& bandera >= 4*/) //Checa si hay un cambio en los archivos después de ejecutatse el programa
            {//De ser así guarda ese cambio para que sea comparado después
                //bandera = 0;
                lastFileUpdateDate = lastUpdatedFile.LastWriteTime;
            }
            else
            {//De no ser así significa que el cambio no es de ninguno de los archivos xlsxque nos interesa
                //bandera++;
                return String.Empty;
            }
            return lastUpdatedFile.Name;
        }

        private void SetInformation(List<Column> data)
        {
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

            FlowLayoutPanel flowItemHeader = GetHeaderLayout();
            
            if (!isAlreadyInitialize)
            {
                this.PanelHeader.Controls.Add(flowItemHeader);
            }

            foreach (Column item in data)
            {
                FlowLayoutPanel flowItem = new FlowLayoutPanel
                {
                    FlowDirection = FlowDirection.LeftToRight,
                    Height = Utils.ROW_HEIGHT,
                    BackColor = Color.Transparent,
                    Width = Panel.Width
                };
                Label item_1 = new Label
                {
                    Text = $"{item.Folio}",
                    Width = padingSpace - 10,
                    Height = Utils.ROW_HEIGHT,
                    Font = Utils.CONTENT_FONT,
                    BackColor = Color.Transparent,
                    ForeColor = Utils.CONTENT_COLOR,
                    TextAlign = ContentAlignment.MiddleCenter
                };

                Label item_2 = new Label
                {
                    Text = $"{item.Ticket}",
                    Width = padingSpace - 10,
                    Height = Utils.ROW_HEIGHT,
                    Font = Utils.CONTENT_FONT,
                    BackColor = Color.Transparent,
                    ForeColor = Utils.CONTENT_COLOR,
                    TextAlign = ContentAlignment.MiddleCenter
                };

                Label item_3 = new Label
                {
                    Text = $"{item.NombreCliente}",
                    Width = padingSpace - 10,
                    Height = Utils.ROW_HEIGHT,
                    Font = Utils.CONTENT_FONT,
                    BackColor = Color.Transparent,
                    ForeColor = Utils.CONTENT_COLOR,
                    TextAlign = ContentAlignment.MiddleCenter
                };

                Label item_4 = new Label
                {
                    Text = $"{item.Zona}",
                    Width = padingSpace - 10,
                    Height = Utils.ROW_HEIGHT,
                    Font = Utils.CONTENT_FONT,
                    BackColor = Color.Transparent,
                    ForeColor = Utils.CONTENT_COLOR,
                    TextAlign = ContentAlignment.MiddleCenter
                };

                Label item_5 = new Label
                {
                    Text = $"{item.Unidad}",
                    Width = padingSpace - 10,
                    Height = Utils.ROW_HEIGHT,
                    Font = Utils.CONTENT_FONT,
                    BackColor = Color.Transparent,
                    ForeColor = Utils.CONTENT_COLOR,
                    TextAlign = ContentAlignment.MiddleCenter
                };

                Label item_6 = new Label
                {
                    Text = $"{item.Chofer}",
                    Width = padingSpace - 10,
                    Height = Utils.ROW_HEIGHT,
                    Font = Utils.CONTENT_FONT,
                    BackColor = Color.Transparent,
                    ForeColor = Utils.CONTENT_COLOR,
                    TextAlign = ContentAlignment.MiddleCenter
                };

                Label item_7 = new Label
                {
                    Text = $"{item.Salida}",
                    Width = padingSpace - 10,
                    Height = Utils.ROW_HEIGHT,
                    Font = Utils.CONTENT_FONT,
                    BackColor = Color.Transparent,
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
        private List<Column> GetData()
        {
            List<List<Column>> tempLists = new List<List<Column>>();
            List<Column> data = new List<Column>();
            
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
            foreach (string excelName in excelNames)
            {
                string filePath = Utils.CURRENT_PATH;
                string strRandom = Utils.RandomString(9);
                string copyFilePath = $@"{filePath}\temp\{strRandom}{excelName}";
                string tempFolder = $@"{filePath}\temp";
                Utils.CreateFolder(tempFolder);
                filePath = filePath + '\\' + excelName;
                File.Copy(filePath, copyFilePath);
                tempLists.Add(ExcelUtil.GetData(copyFilePath, excelName));
            }

            foreach (List<Column> list in tempLists)
            {
                foreach (Column item in list)
                {
                    data.Add(item);
                }
            }
            return data;
        }

        private List<Column> GetData(string originExcelFile)
        {
            List<List<Column>> tempLists = new List<List<Column>>();
            List<Column> data = new List<Column>();

            string filePath = Utils.CURRENT_PATH;
            string strRandom = Utils.RandomString(9);
            string copyFilePath = $@"{filePath}\temp\{strRandom}{originExcelFile}";
            string tempFolder = $@"{filePath}\temp";
            Utils.CreateFolder(tempFolder);
            filePath = filePath + '\\' + originExcelFile;
            File.Copy(filePath, copyFilePath);
            tempLists.Add(ExcelUtil.GetData(copyFilePath, originExcelFile));

            foreach (List<Column> list in tempLists)
            {
                foreach (Column item in list)
                {
                    data.Add(item);
                }
            }
            return data;
        }


        private FlowLayoutPanel GetHeaderLayout()
        {
            FlowLayoutPanel flowItemHeader = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Height = 40,
                Width = Panel.Width
            };

            int padingSpace = (int)(Panel.Width - 9) / Utils.COL_NUMBERS;
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
            return flowItemHeader;
        }

    }
    
}

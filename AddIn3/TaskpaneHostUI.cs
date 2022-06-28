using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swcommands;
using SolidWorks.Interop.swconst;
using System;
using System.Data;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ExcelDataReader;

namespace Sitmikh.SolidWorks.BlankAddin
{
    [ProgId(TaskpaneIntegration.SWTASKPANE_PROGID)]
    public partial class TaskpaneHostUI : UserControl
    {
        #region MEMBERS
        public SldWorks swApp; //приложение  //заменил на TaskpaneIntegration.mSolidWorksApplication сделав mSolidWorksApplication статик
        private ModelDoc2 swModel; //модель
        //private ModelDoc2 tmpObj; //не используемый объект

        //для выбора муфты

        string image3DPath;
        string image2DPath;
        string excelTablePath;
        string sldprtPath;



        //1 кнопка
        private string sldFile;

        //cборка 2 кнопка
        private AssemblyDoc swAssembly;
        private ConfigurationManager swConfigurationMgr = default(ConfigurationManager);
        private Configuration swConfiguration = default(Configuration);
        private Component2 swComponent = default(Component2);
        private string[] sComponents = new string[1]; //не используется?
        private object[] Components;

        //3 кнопка
        private DesignTable swDesingTable;
        private long nTotRow;
        //public string pathFile;  
        private string sRowStr = null;
        private bool bRet = false;

        //5 кнопки
        private Component2 swInsertedComponent;
        private AssemblyDoc Part;

        //для загрузки таблицы
        private string fileName = string.Empty;
        private DataTableCollection tableCollection = null;
        DataTable table = new DataTable();
        
        //для кнопки 12
        public Mouse msMouse;

        //неиспользуемые переменные

        
        #endregion


        public TaskpaneHostUI()
        {
            InitializeComponent();
            #region ненужные переменные при создании===================================
            // swApp = new SldWorks();
            //{
            //    Visible = true
            //};
            ////swModel = swApp.ActiveDoc;
            #endregion
        }

        #region ВЫБОР МУФТЫ, ЗАГРУЗКА ТАБЛИЦЫ И КАРТИНОК
       
        
        /// <summary>
        /// выбор муфты
        /// </summary>
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex != -1)
            {
                switch (comboBox1.SelectedIndex)
                {
                    case 0: // Зубчатая муфта !!
                        excelTablePath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).
                    Replace(@"file:\", string.Empty), "ClutchLibrary", "Муфта зубчатая", "Муфта_зубчатая.xlsx");
                        
                        image3DPath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).
                    Replace(@"file:\", string.Empty), "ClutchLibrary", "Муфта зубчатая", "Муфта_зубчатая_3D.PNG");

                        image2DPath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).
                    Replace(@"file:\", string.Empty), "ClutchLibrary", "Муфта зубчатая", "Муфта_зубчатая_2D.PNG");

                        sldprtPath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).
                    Replace(@"file:\", string.Empty), "ClutchLibrary", "Муфта зубчатая", "Муфта_зубчатая.SLDPRT");

                        label8.Text = "Примечание:";
                        break;


                    case 1: // МУВП !!
                        excelTablePath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).
                    Replace(@"file:\", string.Empty), "ClutchLibrary", "Муфта упругая втулочно-пальецвая", "Муфта_упругая_втулочно-пальцевая.xlsx");
                        
                        image3DPath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).
                    Replace(@"file:\", string.Empty), "ClutchLibrary", "Муфта упругая втулочно-пальецвая", "Муфта упругая втулочно-пальецвая_3D.PNG");

                        image2DPath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).
                    Replace(@"file:\", string.Empty), "ClutchLibrary", "Муфта упругая втулочно-пальецвая", "Муфта упругая втулочно-пальецвая_2D.PNG");

                        sldprtPath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).
                    Replace(@"file:\", string.Empty), "ClutchLibrary", "Муфта упругая втулочно-пальецвая", "Муфта_упругая_втулочно-пальцевая.SLDPRT");

                        label8.Text = "Примечание:";
                        break;


                    case 2: // Фланцевая муфта !!
                        excelTablePath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).
                    Replace(@"file:\", string.Empty), "ClutchLibrary", "Муфта фланцевая", "Муфта_фланцевая.xlsx");
                        
                        image3DPath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).
                    Replace(@"file:\", string.Empty), "ClutchLibrary", "Муфта фланцевая", "Муфта_фланцевая_3D.PNG");

                        image2DPath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).
                    Replace(@"file:\", string.Empty), "ClutchLibrary", "Муфта фланцевая", "Муфта_фланцевая_2D.PNG");

                        sldprtPath = Path.Combine(Path.GetDirectoryName(typeof(TaskpaneIntegration).Assembly.CodeBase).
                    Replace(@"file:\", string.Empty), "ClutchLibrary", "Муфта фланцевая", "Муфта_фланцевая.SLDPRT");
                        
                        label8.Text = "Примечание: Номинальный крутящий момент указан для полумуфт из стали. Для полумуфт из чугуна крутящий момент необходимо уменьшить вдвое";
                        break;


                    case 3: // Муфта упругая со звездочкой
                        
                        label8.Text = "Примечание:Эта муфта доступна в платной версии программы";
                        break;
                    
                    
                    case 4: // Муфта торообразная !

                        label8.Text = "Примечание:Эта муфта доступна в платной версии программы";
                        break;
                    
                    
                    case 5: // Муфта фрикционная

                        label8.Text = "Примечание:Эта муфта доступна в платной версии программы";
                        break;
                   
                    
                    case 6: // Муфта свободного хода

                        label8.Text = "Примечание:Эта муфта доступна в платной версии программы";
                        break;
                    
                    
                    case 7: // Муфта центробежная

                        label8.Text = "Примечание:Эта муфта доступна в платной версии программы";
                        break;
                    
                    
                    case 8: // Муфта предохронительная

                        label1.Text = "Примечание:Эта муфта доступна в платной версии программы";
                        break;
                   
                    
                    case 9: // Муфта продольно-свертная
                       
                        label8.Text = "Примечание:Эта муфта доступна в платной версии программы";
                        break;
                }
                try
                {
                    LoadTable();
                    

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка! Не удалось найти указанную таблицу", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void LoadTable()
        {
            FileStream stream = File.Open(excelTablePath, FileMode.Open,
                FileAccess.Read);

            var ds = new DataSet();

            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
            {
                ds = reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = false
                    }
                });
            }

            while (ds.Tables[0].Columns.Count > 5)
            {
                ds.Tables[0].Columns.RemoveAt(ds.Tables[0].Columns.Count - 1);
            }
            ds.Tables[0].AcceptChanges();

            tableCollection = ds.Tables;

            table = tableCollection[0]; //принимаем таблицу из Excel (ЭТА ЧАСТЬ ДО КОНЦА БЫЛА НИЖЕ LOAD TABLE в SELECTEDINDEXCHANGED)
            dataGridView1.DataSource = table;

            dataGridView1.Rows.RemoveAt(0);
            dataGridView1.Rows.RemoveAt(0);
            dataGridView1.Columns.RemoveAt(0);

            dataGridView1.Columns[0].HeaderText = "d"; //оформляем столбцы
            dataGridView1.Columns[0].Width = 50;
            dataGridView1.Columns[1].HeaderText = "D";
            dataGridView1.Columns[1].Width = 50;
            dataGridView1.Columns[2].HeaderText = "L";
            dataGridView1.Columns[2].Width = 50;
            dataGridView1.Columns[3].HeaderText = "Tкр";
            dataGridView1.Columns[3].Width = 500;

            pictureBox1.Image = Image.FromFile(image3DPath);
            pictureBox2.Image = Image.FromFile(image2DPath);
        }
        #endregion

        #region ФИЛЬТР ТАБЛИЦЫ И ВЫДЕЛЕНИЕ СТРОКИ
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //table.DefaultView.RowFilter = string.Format("[{0}] LIKE '%{1}%'", "d", textBox1.Text);
            //table.DefaultView.RowFilter = string.Format("Convert([{0}], 'System.String') LIKE '%{1}%'", "d", textBox1.Text);
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            //dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Selected = true;
        }

        #endregion


        #region НЕНУЖНЫЕ КНОПКИ (С ГОТОВЫМИ РЕШЕНИЯМИ)
        private void button1_Click(object sender, EventArgs e) //открытие муфты в новом окне сборки
        {

            //MessageBox.Show("рш");
            int fileError = 0;
            int fileWarning = 0;
            listBox1.Items.Clear();
            sldFile = @"D:\VKR\Addin\ClutchLibrary\";

            if (checkBox1.Checked == true)
                sldFile += checkBox1.Text + @"\Зубчатая муфта.sldprt";
            if (checkBox2.Checked == true)
                sldFile += checkBox2.Text + @"\Сборка МУВП 3 пальца.SLDASM";
            if (checkBox3.Checked == true)
                sldFile += checkBox3.Text + @"\Фланцевая муфта.SLDASM";
            if (checkBox4.Checked == true)
                sldFile += checkBox4.Text + @"\Муфта со звездочкой 4.SLDASM";
            #region
            int clutchIndex = comboBox1.SelectedIndex;
            switch (clutchIndex)
            {
                case 0:

                    sldFile += comboBox1.SelectedValue.ToString() + @"\Зубчатая муфта.sldprt";
                    //MessageBox.Show("0");
                    break;
                case 1:
                    //MessageBox.Show(Convert.ToString(1));
                    sldFile += comboBox1.SelectedValue.ToString() + @"\Сборка МУВП 3 пальца.SLDASM";
                    // MessageBox.Show(libraryPath);
                    break;
                case 2:
                    //MessageBox.Show(Convert.ToString(2));
                    sldFile += comboBox1.SelectedValue.ToString() + @"\Фланцевая муфта.SLDASM";
                    // MessageBox.Show(libraryPath);
                    break;
                case 3:
                    //MessageBox.Show(Convert.ToString(3));
                    sldFile += comboBox1.SelectedValue.ToString() + @"\Муфта со звездочкой 4.SLDASM";
                    //MessageBox.Show(libraryPath);
                    break;
                default:
                    MessageBox.Show("Выберите тип муфты");
                    break;
            }
            MessageBox.Show(sldFile);
            // GetDataFromFile();
            #endregion
            label1.Text = "Выбрать муфту";
            label1.Text += sldFile;

            swModel = TaskpaneIntegration.mSolidWorksApplication.OpenDoc6(
                    sldFile,
                    (int)swDocumentTypes_e.swDocASSEMBLY,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                    "",
                    ref fileError,
                    ref fileWarning);
            if (swModel == null)
            {
                return;
            }
            else
            {
                Debug.Print("File = " + swModel.GetPathName());
                Debug.Print("");
            }


        }

        private void button2_Click(object sender, EventArgs e) //не нужна
        {


            listBox1.Items.Clear();

            swAssembly = (AssemblyDoc)swModel;
            swConfigurationMgr = (ConfigurationManager)swModel.ConfigurationManager; //служит для создания, выбора и просмотра многочисленных конфигураций деталей и сборок в документе
            swConfiguration = (Configuration)swConfigurationMgr.ActiveConfiguration; //управление состояниями сборки или детали
            Components = swAssembly.GetComponents(true);

            for (int i = 0; i < Components.Length; i++)
            {
                swComponent = (Component2)Components[i];
                if ((swComponent.IsLoaded()))
                {
                    Debug.Print("Component: " + swComponent.Name + " is loaded.");
                    listBox1.Items.Add(swComponent.Name);
                }
                else
                {
                    Debug.Print("Component: " + swComponent.Name + " is not loaded.");
                }
                Debug.Print("  Suppressed: " + swConfiguration.GetComponentSuppressionState(swComponent.Name));
                Debug.Print("");
            }
        }
        private void button3_Click(object sender, EventArgs e) //не нужна
        {
            listBox1.Items.Clear();

            swDesingTable = swModel.GetDesignTable(); //получение таблицы параметров (ошибка)
            bRet = swDesingTable.Attach(); //активация таблицы параметров

            nTotRow = swDesingTable.GetTotalRowCount();

            for (int i = 1; i <= nTotRow; i++)
            {
                sRowStr = swDesingTable.GetEntryText(i, 0);
                listBox1.Items.Add(sRowStr);
            }
            swDesingTable.Detach(); //деактивация таблицы параметров
        }

        private void button4_Click(object sender, EventArgs e) //не нужна
        {
            swModel.ShowConfiguration2(listBox1.Text); //Показывает именованную конфигурацию, переключаясь на эту конфигурацию и делая ее активной конфигурацией.
        }
        #endregion


        #region ДОБАВЛЕНИЕ ТАБЛИЦЫ
        /// <summary>
        /// добавляем таблицу в из Excel в память
        /// </summary>
        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                LoadTable();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка! Не удалось найти указанную таблицу", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// добавляем таблицу из памяти в dataGridView1 и загружаем картинки в pictureBox1 и 2.
        /// </summary>
        private void button9_Click(object sender, EventArgs e)
        {
            DataTable table = tableCollection[0];
            dataGridView1.DataSource = table;
            dataGridView1.Rows.RemoveAt(0);
            dataGridView1.Rows.RemoveAt(0);
            dataGridView1.Columns.RemoveAt(0);
            dataGridView1.Columns[1].HeaderText = "Fuck";

            pictureBox1.Image = Image.FromFile(@"D:\VKR\Addin\ClutchLibrary\Upd\Муфта зубчатая_3D.PNG");
            pictureBox2.Image = Image.FromFile(@"D:\VKR\Addin\ClutchLibrary\Upd\Муфта_зубчатая_Чертеж.PNG");
        }
        
        
        #endregion


        #region ДОБАВЛЕНИЕ МУФТЫ В СБОРКУ
      

        /// <summary>
        /// тестовое добавление детали через AddComponents3
        /// </summary>
        private void button12_Click(object sender, EventArgs e)
        {
            AssemblyDoc assemb;
            string[] compNames = new string[1];
            double[] compXforms = new double[16];
            string[] compCoordSysNames = new string[1];
            object vcompNames;
            object vcompXforms;
            object vcompCoordSysNames;
            object vcomponents;


            assemb = (AssemblyDoc)TaskpaneIntegration.mSolidWorksApplication.ActiveDoc;

            if (((assemb != null)))
            {
                compNames[0] = @"D:\VKR\Addin\ClutchLibrary\Upd\Муфта зубчатая.SLDPRT";

                // Define the transformation matrix. See the IMathTransform API documentation. 

                // Add a rotational diagonal unit matrix (zero rotation) to the transform
                // x-axis components of rotation
                compXforms[0] = 1.0;
                compXforms[1] = 0.0;
                compXforms[2] = 0.0;
                // y-axis components of rotation
                compXforms[3] = 0.0;
                compXforms[4] = 1.0;
                compXforms[5] = 0.0;
                // z-axis components of rotation
                compXforms[6] = 0.0;
                compXforms[7] = 0.0;
                compXforms[8] = 1.0;

                // Add a translation vector to the transform (zero translation) 
                compXforms[9] = 0.0;
                compXforms[10] = 0.0;
                compXforms[11] = 0.0;

                // Add a scaling factor to the transform
                compXforms[12] = 0.0;

                // The last three elements in the transformation matrix are unused

                // Register the coordinate system name for the component 
                compCoordSysNames[0] = "Coordinate System1";

                // Add the components to the assembly. 
                vcompNames = compNames;
                vcompXforms = compXforms;
                //vcompXforms = Nothing   //also achieves zero rotation and translation of the component
                vcompCoordSysNames = compCoordSysNames;

                vcomponents = assemb.AddComponents3((vcompNames), (vcompXforms), (vcompCoordSysNames));
            }
        }

        /// <summary>
        /// тестовое добавление детали [1] через OpenDoc6/CloseDoc, AddComponent и AddComponent5
        /// </summary>
        private void button5_Click(object sender, EventArgs e)
        {
            swModel = TaskpaneIntegration.mSolidWorksApplication.ActiveDoc;
            //swModel = swApp.ActiveDoc;
            //Part = (AssemblyDoc)swApp.ActiveDoc;
            int fileError = 0;
            int fileWarning = 0;

            swModel = (ModelDoc2)TaskpaneIntegration.mSolidWorksApplication.OpenDoc6(@"D:\VKR\Addin\ClutchLibrary\Upd\Муфта зубчатая.SLDPRT",
            1,
            (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
            "", //параметр Configuration открывает модель именно в той конфигурации, в какой мы задумали 
           ref fileError,
           ref fileWarning);

            Part = TaskpaneIntegration.mSolidWorksApplication.ActivateDoc3("loaded_document",
                true,
                0,
                ref fileError);

            swInsertedComponent = Part.AddComponent5(@"D:\VKR\Addin\ClutchLibrary\Upd\Муфта зубчатая.SLDPRT",
                 0,
                 "Default",
                 false,
                 "",
                 5.42027305169652E-02, 6.53507206261547E-02, 4.03630755082414E-02);
            TaskpaneIntegration.mSolidWorksApplication.CloseDoc(@"D:\VKR\Addin\ClutchLibrary\Upd\Муфта зубчатая.SLDPRT");
            bool boolstatus = Part.AddComponent(@"D:\VKR\Addin\ClutchLibrary\Upd\Муфта зубчатая.SLDPRT", 4.26089577376842E-03, 8.44019707292318E-02, 0.111121182912029);



        }

        /// <summary>
        /// тестовое добавление детали [2] через OpenDoc6/CloseDoc, AddComponent и AddComponent5
        /// </summary>
        private void button13_Click(object sender, EventArgs e)
        {
            //object Part;
            bool boolstatus;
            int longstatus = 0, longwarnings = 0;

            Part = TaskpaneIntegration.mSolidWorksApplication.ActiveDoc;
            ModelDoc2 tmpObj;
            int errors = 0;
            Component2 swInsertedComponent; 


            tmpObj = TaskpaneIntegration.mSolidWorksApplication.OpenDoc6(sldprtPath,
                1,
                32,
                 "10000", //параметр Configuration открывает модель именно в той конфигурации, в какой мы задумали 
                ref longstatus,
                ref longwarnings);
            
            swInsertedComponent = Part.AddComponent5(sldprtPath,
                0,
                "Default",
                true,
                "4000", //параметр Configuration открывает модель именно в той конфигурации, в какой мы задумали
                4.04840074935108E-02, 2.44451029699681E-02, 0.025849580254035);
            TaskpaneIntegration.mSolidWorksApplication.CloseDoc(sldprtPath);
            boolstatus = Part.AddComponent(sldprtPath, 
                -8.58845433685929E-03, 2.28718737489544E-02, 4.45478721521795E-02);
            boolstatus = Part.AddComponent(sldprtPath,
                1.73858007183298E-02, 1.46586254122667E-02, 4.27317911526188E-02);


        }
        #endregion


        #region ПОПЫТКА ОТСЛЕДИТЬ ПОЛОЖЕНИЕ УКАЗАТЕЛЯ МЫШИ
        /// <summary>
        /// попытка отследить положение указателя мыши
        /// </summary>
        private void button14_Click(object sender, EventArgs e)
        {
            ModelDoc2 swModel = default(ModelDoc2);
            ModelDocExtension swModelDocExt = default(ModelDocExtension);
            ModelView swModelView = default(ModelView);
            Mouse TheMouse = default(Mouse);
            int i = 0;
            double x = 0;
            string fileName = null;
            int errors = 0;
            int warnings = 0;
            bool status = false;

            fileName = @"D:\VKR\Addin\ClutchLibrary\Upd\Муфта зубчатая.SLDPRT";
            swModel = (ModelDoc2)TaskpaneIntegration.mSolidWorksApplication.OpenDoc6(fileName, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errors, ref warnings);
            swModelDocExt = (ModelDocExtension)swModel.Extension;
            swModelView = (ModelView)swModel.GetFirstModelView();
            TheMouse = (Mouse)swModelView.GetMouse();

            //Set up events
            msMouse = (Mouse)TheMouse;
            AttachEventHandlers();

            x = 0;

            Debug.Print("Fit model to graphics area");
            swModelDocExt.RunCommand((int)swCommands_e.swCommands_ZoomToFit, "");

            //Move the mouse
            status = TheMouse.Move(50, 150, (int)swMouse_e.swMouse_Absolute + (int)swMouse_e.swMouse_MouseMove + (int)swMouse_e.swMouse_LeftDown);
            Debug.Print("First call to Move: " + status);
            Debug.Print("Calls to Move within loop:");
            for (i = 1; i <= 5; i++)
            {
                status = TheMouse.Move(5, 5, (int)swMouse_e.swMouse_MouseMove);
                Debug.Print("  Call " + i + " to Move: " + status);
            }
            status = TheMouse.Move(0, 0, (int)swMouse_e.swMouse_LeftUp);
            Debug.Print("Last call to Move: " + status);

            status = TheMouse.MoveXYZ(0.03720615681732, 0.0316583060694, 0.04991700841805, (int)swMouse_e.swMouse_LeftDown);
            Debug.Print("Call to MoveXYZ: " + status);
            Debug.Print("Calls to Move within next loop:");
            for (i = 1; i <= 5; i++)
            {
                status = TheMouse.Move(5, 5, (int)swMouse_e.swMouse_MouseMove);
                Debug.Print("  Call " + (i + 1).ToString() + " to Move: " + status);
            }

            status = TheMouse.Move(10, 10, (int)swMouse_e.swMouse_LeftUp);
            Debug.Print("Last call to Move: " + status);

            Debug.Print("Change view to *Front");
            swModelDocExt.RunCommand((int)swCommands_e.swCommands_Front, "");

        }

        public void AttachEventHandlers()
        {
            AttachSWEvents();
        }

        public void AttachSWEvents()
        {

            msMouse.MouseSelectNotify += this.ms_MouseSelectNotify;
            msMouse.MouseLBtnDownNotify += this.ms_MouseLBtnDownNotify;

        }

        private int ms_MouseSelectNotify(int Ix, int Iy, double x, double y, double z)
        {
            Debug.Print("Selection made:");
            Debug.Print(" Mouse location:");
            Debug.Print("   Window space coordinates:");
            Debug.Print("     " + Ix);
            Debug.Print("     " + Iy);
            Debug.Print("   World space coordinates:");
            Debug.Print("     " + x);
            Debug.Print("     " + y);
            Debug.Print("     " + z);

            return 1;
        }

        private int ms_MouseLBtnDownNotify(int x, int y, int WParam)
        {
            Debug.Print("Left-mouse button pressed.");

            return 1;
        }
        #endregion


        #region ЗАЩИТА ПОЛЕЙ ВВОДА
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0') && (e.KeyChar <= '9'))
                    return;
            if (Char.IsControl(e.KeyChar))
            {
                if (e.KeyChar == (char)Keys.Enter)
                {
                    if (sender.Equals(textBox1))
                        textBox2.Focus();
                    else comboBox1.Focus();
                }
                return;
            }
            e.Handled = true;
        }



        #endregion

       
    }
}


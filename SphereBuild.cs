using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Threading;
using System.Windows.Forms;
using SaveSystem;
using SharpGL;
using SharpGL.SceneGraph;
using SharpGL.SceneGraph.Effects;
using SharpGL.SceneGraph.Quadrics;
using Application = System.Windows.Forms.Application;
using Point = System.Drawing.Point;

namespace SphereBuilder
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            Thread t = new Thread(new ThreadStart(StartForm));
            t.Start();
            Thread.Sleep(5000);
            InitializeComponent();
            ReadDataFromFile();
            t.Abort();
            UpdateAxesTextBoxes();
        }

        public void StartForm()
        {
            Application.Run(new LoadingScreen());
        }

        private const string _sphereImagePath = "D:\\Uni\\2 курс 2 семестр\\РПВС\\РПВС курсовая\\SphereImage.jpeg";
        private const string _saveFilePath = "D:\\Uni\\2 курс 2 семестр\\РПВС\\РПВС курсовая\\SaveData.save";
        private int _maxRadius = 20;
        private int _minRadius = 1;

        private int _radius = 13; // Радиус сферы
        private SharpGL.SceneGraph.Quadrics.DrawStyle _style = SharpGL.SceneGraph.Quadrics.DrawStyle.Line; // Стиль отображения сферы (точечный, линейный)
        private int _stacks = 20; // Кол-во поперечных линий, строящие плотность сферы
        private int _slices = 20; // Кол-во продольных линий, строящие плтность сферы

        private float _rotationAngle = 1000.0f;
        private int _deltaRotate = 1;
        private Color _sphereColor = Color.White;  // Цвет сферы

        private float _translationX = 0.0f;
        private float _translationY = 0.0f;
        private float _translationZ = 0.0f;

        private float _angleX = 0.0f;
        private float _angleY = 90.0f;
        private float _angleZ = 90.0f;

        private bool _axisX = true;
        private bool _axisY = false;
        private bool _axisZ = false;

        private const float _minAngle = 0.0f;
        private const float _maxAngle = 359.0f;

        private int numberOfSpheres = 1;

        void ReadDataFromFile() // Чтие последних сохранённых настроек приложения
        {
            BinaryFormatter _formatter;
            FileStream _stream;
            MySphere data;
            if (File.Exists(_saveFilePath) == false)
            {
                _formatter = new BinaryFormatter();
                _stream = new FileStream(_saveFilePath, FileMode.Create);
                data = new MySphere(_radius, _stacks, _slices, _style, _sphereColor, _deltaRotate, _translationX, _translationY, _translationZ, _angleX, _angleY, _angleZ, _axisX, _axisY, _axisZ);
                _formatter.Serialize(_stream, data);
                _stream.Close();

                MessageBox.Show("a");
                _formatter = null;
                _stream = null;
                data = null;
                
            }
            // Инициализация нового бинарного конвертатора
            _formatter = new BinaryFormatter();
            // Инициализация поля взаимодействия с файлами
            _stream = new FileStream(_saveFilePath, FileMode.Open);
            // Инициализация поля для сохранения данных (RetainedInfo) прочтенной из файла инфо
            data = _formatter.Deserialize(_stream) as MySphere;
            // Закрытие файла
            _stream.Close();

            // Запись сохранённых настроек
            _radius = data.Radius;
            _style = data.Style;
            _stacks = data.Stacks;
            _slices = data.Slices;
            _deltaRotate = data.Speed;
            //_deltaRotate = data.

            _sphereColor = data.Color;

            _translationX = data.TranslationX;
            _translationY = data.TranslationY;
            _translationZ = data.TranslationZ;
            

            _angleX = data.AngleX;
            _angleY = data.AngleY;
            _angleZ = data.AngleZ;

            _axisX = data.AxisX;
            _axisY = data.AxisY;
            _axisZ = data.AxisZ;

            // Заполнение полей на форме
            radiusTrackBar.Value = _radius;
            radiusTextBox.Text = _radius.ToString();
            stacksTextBox.Text = _stacks.ToString();
            slicesTextBox.Text = _slices.ToString();
            trackBar1.Value = _deltaRotate;
        }
        private void openGLControl1_OpenGLDraw(object sender, SharpGL.RenderEventArgs args) // Динамическое взаимодействие со сферами
        {
            List<MySphere> spheres = new List<MySphere>();
            switch(numberOfSpheres)
            {
                case 1:
                    spheres.Add(new MySphere(_radius, _stacks, _slices, _style, _sphereColor, _deltaRotate, _translationX, _translationY, _translationZ, _angleX, _angleY, _angleZ, _axisX, _axisY, _axisZ));
                    _maxRadius = 20;
                    break;
                case 2:
                    spheres.Add(new MySphere(_radius, _stacks, _slices, _style, _sphereColor, _deltaRotate, -20.0f, _translationY, _translationZ, _angleX, _angleY, _angleZ, _axisX, _axisY, _axisZ));
                    spheres.Add(new MySphere(_radius, _stacks, _slices, _style, _sphereColor, _deltaRotate, 20.0f, _translationY, _translationZ, _angleX, _angleY, _angleZ, _axisX, _axisY, _axisZ));
                    _maxRadius = 14;
                    break;
                case 3:
                    spheres.Add(new MySphere(_radius, _stacks, _slices, _style, _sphereColor, _deltaRotate, -15.0f, -10.0f, _translationZ, _angleX, _angleY, _angleZ, _axisX, _axisY, _axisZ));
                    spheres.Add(new MySphere(_radius, _stacks, _slices, _style, _sphereColor, _deltaRotate, 15.0f, -10.0f, _translationZ, _angleX, _angleY, _angleZ, _axisX, _axisY, _axisZ));
                    spheres.Add(new MySphere(_radius, _stacks, _slices, _style, _sphereColor, _deltaRotate, _translationX, 10.0f, _translationZ, _angleX, _angleY, _angleZ, _axisX, _axisY, _axisZ));
                    _maxRadius = 12;
                    
                    break;

            }
            maxradiusLabel.Text = _maxRadius.ToString();
            if (_radius > _maxRadius)
            {
                _radius = _maxRadius;
            }
            radiusTrackBar.Maximum = _maxRadius;
            radiusTrackBar.Value = _radius;
            radiusTextBox.Text = _radius.ToString();

            OpenGlArtist.DrawingSpheres(openGLControl1, spheres, _rotationAngle);
            _rotationAngle += spheres[0].Speed;
        }
        // Загрузка данных с сервера




        /// //////////////////////////////////////////////////////////////////////
        
        private void loadDataFromServerButton_Click(object sender, EventArgs e)
        {
            ServerAPI serv = new ServerAPI();
            int[] r = ServerAPI.GetMessageFromSocket(11000);
            if (r[0] != 228)
            {
                _radius = r[0];

            }
            else
            {
                MessageBox.Show("Somethin is wrong with server");
                return;
            }

            if (r[1] == 1)
            {
            }
            _stacks = r[2];
            _slices = r[3];
            MessageBox.Show("We get from Server : \nSphere Radius = " + r[0] + "\nSphere Color = Purple" + "\nSphere Stacks = " + r[2] + "\nSphere Slices = " + r[3]);
        }
        private void saveToWordButton_Click(object sender, EventArgs e)
        {

        }
        //Helper
        private void startHelperButton_Click(object sender, EventArgs e)
        {
            Process process = Process.Start("help.chm");
            int id = process.Id;
            Process tempProc = Process.GetProcessById(id);
            this.Visible = false;
            tempProc.WaitForExit();
            this.Visible = true;
        }



        ////////////////////////////////////////////////////////////////////////////

        

        // Создание изображения сферы
        private void CreateSphereImage(string path)
        {
            try
            {
                // Задержка для того чтобы OpenFileDialog успел закрыть и не влез в скриншот
                Thread.Sleep(200);
                // Создание точечного рисунка и Graphics
                Bitmap printscreen = new Bitmap(openGLControl1.Width + 150, openGLControl1.Height + 110);
                Graphics graphics = Graphics.FromImage(printscreen);
                // Верхняя левая точка скриншота
                Point point = new Point(this.Left + openGLControl1.Left + 70, this.Top + openGLControl1.Top + 80);
                // Делаем скриншот области начиная с указанной точки рахмером printscreen.Size
                graphics.CopyFromScreen(point, Point.Empty, printscreen.Size);
                // Сохраняем по пути в jpeg
                printscreen.Save(path, System.Drawing.Imaging.ImageFormat.Jpeg);
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
        // Панель управления цветом сферы
        private void setNewColorButton_Click(object sender, EventArgs e)
        {
            ColorDialog _colorDialog = new ColorDialog();
            if(_colorDialog.ShowDialog() == DialogResult.OK)
            {
                _sphereColor = _colorDialog.Color;
            }
        }
        private void redColorButton_Click(object sender, EventArgs e)
        {
            _sphereColor = Color.Red;
        }
        private void orangeColorButton_Click(object sender, EventArgs e)
        {
            _sphereColor = Color.Orange;
        }
        private void yellowColorButton_Click(object sender, EventArgs e)
        {
            _sphereColor = Color.Yellow;
        }
        private void limeColorButton_Click(object sender, EventArgs e)
        {
            _sphereColor = Color.Lime;
        }
        private void cyanColorButton_Click(object sender, EventArgs e)
        {
            _sphereColor = Color.Cyan;
        }
        private void blueColorButton_Click(object sender, EventArgs e)
        {
            _sphereColor = Color.Blue;
        }
        private void pinkColorButton_Click(object sender, EventArgs e)
        {
            _sphereColor = Color.Purple;
        }
        private void silverColorButton_Click(object sender, EventArgs e)
        {
            _sphereColor = Color.Silver;
        }
        private void whiteColorButton_Click(object sender, EventArgs e)
        {
            _sphereColor = Color.White;
        }
        private void goodColorButton_Click(object sender, EventArgs e)
        {
            _sphereColor = Color.Aqua;
        }

        // Панель управления стилем отрисовки
        private void linesRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            _style = DrawStyle.Line;
        }
        private void pointsRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            _style = DrawStyle.Point;
        }

        // Панель управления радиусом сферы
        private void setSphereRadiusButton_Click(object sender, EventArgs e)
        {
            if (radiusTextBox != null)
            {
                int _newRadius;
                bool success = int.TryParse(radiusTextBox.Text, out _newRadius);
                if (success && _newRadius >= _minRadius && _newRadius <= _maxRadius)
                    _radius = _newRadius;
                else if (success == false)
                    MessageBox.Show("Введены некорректные данные!\r\nВведённое значение не является числом\r\nПроверьте правильность и повторите попытку.");
                else if (_newRadius < _minRadius)
                    MessageBox.Show($"Невозможно установить текущее значение, так как значение ({_newRadius}) < минимального радиуса ({_minRadius})!\r\nИзмените значение и повторите попытку.");
                else
                    MessageBox.Show($"Невозможно установить текущее значение, так как значение ({_newRadius}) > максимального радиуса ({_maxRadius})!\r\nИзмените значение и повторите попытку.");
            }
            else
                MessageBox.Show("Невозможно установить текущее значение, так как ячейка значения пуста!\r\nЗаполните ячёйку и повторите попытку.");
        }
        private void radiusTrackBar_Scroll(object sender, EventArgs e)
        {
            int _newRadius = radiusTrackBar.Value;
            _radius = _newRadius;
            radiusTextBox.Text = _radius.ToString();
        }

        // Панель управления плотностью сферы
        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            _stacks = stacksTrackBar.Value;
            stacksTextBox.Text = _stacks.ToString();
        }
        private void trackBar2_Scroll(object sender, EventArgs e)
        {
            _slices = slicesTrackBar.Value;
            slicesTextBox.Text = _slices.ToString();
        }

        private void microsoftWordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //
        }
        // Сохранение в Excel
        private void microsoftExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Создание диалога
            SaveFileDialog _saveFileDialog = new SaveFileDialog();
            // Настройка фильтра под таблицы Excel
            _saveFileDialog.Filter = "Excel Files(*.xlsx)|*.xlsx";
            // При открытии и нажатии "OK"
            if (_saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = _saveFileDialog.FileName;
                // Создаем объект класса ExcelSave для сохранения данных
                using (ExcelSave helper = new ExcelSave())
                {
                    if (helper.Open(filePath))
                    {
                        // Запись данных
                        helper.SetText("A1", "Время сохранения:");
                        helper.SetTime("B1", DateTime.Now);

                        helper.SetText("A3", "Тип построения:");
                        helper.SetText("B3", _style.ToString());

                        helper.SetText("A4", "Кол-во меридиан:");
                        helper.SetText("B4", _stacks);

                        helper.SetText("A5", "Кол-во парралелей:");
                        helper.SetText("B5", _slices);

                        helper.SetText("A6", "Радиус сферы:");
                        helper.SetText("B6", _radius);

                        helper.SetText("A7", "Скорость сферы:");
                        helper.SetText("B7", _deltaRotate);

                        helper.SetText("A8", "Цвет сферы:");
                        helper.SetText("B8", _sphereColor.Name.ToString());

                        helper.SetText("A10", "Угол X:");
                        helper.SetText("B10", _angleX);

                        helper.SetText("A11", "Угол Y:");
                        helper.SetText("B11", _angleY);

                        helper.SetText("A12", "Угол Z:");
                        helper.SetText("B12", _angleZ);

                        helper.SetText("A14", "|| оси X:");
                        helper.SetText("B14", _axisX);

                        helper.SetText("A15", "|| оси Y:");
                        helper.SetText("B15", _axisY);

                        helper.SetText("A16", "|| оси Z:");
                        helper.SetText("B16", _axisZ);

                        helper.SetText("D1", "Изображение сферы:");
                        // Создание изображения сферы
                        CreateSphereImage(_sphereImagePath);

                        // Добавление изображения в Excel
                        helper.SetImage(3, 2, _sphereImagePath);

                        //Сохранение
                        helper.Save();
                    }
                    else
                    {
                        MessageBox.Show("Файл сохранения не найден!");
                    }
                }
            }
        }
        private void microsoftPowerPointToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //
        }
        private void импортироватьНастройкиССервераToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //
        }
        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        // Управление положением сфер в пространстве
        private void xAxisParallelRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            VisibleAxesAngleSettings(false);

            _angleX = 0.0f;
            _angleY = 90.0f;
            _angleZ = 90.0f;

            UpdateAxesTextBoxes();
        }
        private void yAxisParallelRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            VisibleAxesAngleSettings(false);

            _angleX = 90.0f;
            _angleY = 0.0f;
            _angleZ = 90.0f;

            UpdateAxesTextBoxes();
        }
        private void zAxisParallelRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            VisibleAxesAngleSettings(false);

            _angleX = 90.0f;
            _angleY = 90.0f;
            _angleZ = 0.0f;

            UpdateAxesTextBoxes();
        }
        private void customAxisParallelRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            VisibleAxesAngleSettings(true);
        }
        private void setXAngleButton_Click(object sender, EventArgs e)
        {
            try
            {
                float angle = Convert.ToSingle(xAxisAngleTextBox.Text);
                if (angle >= _minAngle && angle <= _maxAngle)
                {
                    _angleX = angle;
                }
                else
                    throw new Exception($"Введённое значение меньше минимального ({_minAngle} ̊ ) или больше максимального ({_maxAngle} ̊ ) возможного значения угла. Введите корректное значение и повторите попытку"); ; ;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex}");
            }
        }
        private void setYAngleButton_Click(object sender, EventArgs e)
        {
            try
            {
                float angle = Convert.ToSingle(yAxisAngleTextBox.Text);
                if (angle >= _minAngle && angle <= _maxAngle)
                {
                    _angleY = angle;
                }
                else
                    throw new Exception($"Введённое значение меньше минимального ({_minAngle} ̊ ) или больше максимального ({_maxAngle} ̊ ) возможного значения угла. Введите корректное значение и повторите попытку"); ; ;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex}");
            }
        }
        private void setZAngleButton_Click(object sender, EventArgs e)
        {
            try
            {
                float angle = Convert.ToSingle(zAxisAngleTextBox.Text);
                if (angle >= _minAngle && angle <= _maxAngle)
                {
                    _angleZ = angle;
                }
                else
                    throw new Exception($"Введённое значение меньше минимального ({_minAngle} ̊ ) или больше максимального ({_maxAngle} ̊ ) возможного значения угла. Введите корректное значение и повторите попытку"); ; ;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex}");
            }
        }
        private void VisibleAxesAngleSettings(bool enable)
        {
            setXAngleButton.Visible = enable;
            setYAngleButton.Visible = enable;
            setZAngleButton.Visible = enable;

            xAxisAngleTextBox.ReadOnly = !enable;
            yAxisAngleTextBox.ReadOnly = !enable;
            zAxisAngleTextBox.ReadOnly = !enable;
        }
        private void UpdateAxesTextBoxes()
        {
            xAxisAngleTextBox.Text = _angleX.ToString();
            yAxisAngleTextBox.Text = _angleY.ToString();
            zAxisAngleTextBox.Text = _angleZ.ToString();
        }

        // Управление осью вращения
        private void yAxisRotatyRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            _axisX = false;
            _axisY = true;
            _axisZ = false;
        }
        private void zAxisRotatyRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            _axisX = false;
            _axisY = false;
            _axisZ = true;
        }
        private void xAxisRotatyRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            _axisX = true;
            _axisY = false;
            _axisZ = false;
        }

        
        
        
        
        private void SaveDataToFile(object sender, EventArgs e)
        {
            BinaryFormatter _formatter;
            FileStream _stream;
            MySphere data;

            _stream = new FileStream(_saveFilePath, FileMode.OpenOrCreate);
            _formatter = new BinaryFormatter();
            data = new MySphere(_radius, _stacks, _slices, _style, _sphereColor, _deltaRotate, _translationX, _translationY, _translationZ, _angleX, _angleY, _angleZ, _axisX, _axisY, _axisZ);
            _formatter.Serialize(_stream, data);
            _stream.Close();
        }

        private void trackBar1_Scroll_1(object sender, EventArgs e)
        {
            _deltaRotate = trackBar1.Value;
        }
        // Управление режимом отображения(кол-во сфер)
        private void двеСферыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            numberOfSpheres = 2;
        }
        private void триСферыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            numberOfSpheres = 3;
        }
        private void однаСфераToolStripMenuItem_Click(object sender, EventArgs e)
        {
            numberOfSpheres = 1;
        }
    }

    [Serializable]
    internal class MySphere
    {
        // Основные свойства сферы
        public int Radius { get; private set; }
        public int Stacks { get; private set; }
        public int Slices { get; private set; }
        public SharpGL.SceneGraph.Quadrics.DrawStyle Style { get; private set; }
        public Color Color { get; private set; }

        public int Speed { get; private set; }

        // Положение в пространстве
        public float TranslationX { get; private set; }
        public float TranslationY { get; private set; } 
        public float TranslationZ { get; private set; }
        
        // Углы наклона относительно каждой оси
        public float AngleX { get; private set; }
        public float AngleY { get; private set; }
        public float AngleZ { get; private set; }

        // Ось вращения
        public bool AxisX { get; private set; }
        public bool AxisY { get; private set; }
        public bool AxisZ { get; private set; }
        internal MySphere(int radius, int stacks, int slices, SharpGL.SceneGraph.Quadrics.DrawStyle style, Color color, int speed, float translstionX, float translationY, float translationZ, float angleX, float angleY, float angleZ, bool axisX, bool axisY, bool axisZ)
        {
            Radius = radius;
            Stacks = stacks;
            Slices = slices;
            Style = style;
            Color = color;
            Speed = speed;

            TranslationX = translstionX;
            TranslationY = translationY;
            TranslationZ = translationZ;

            this.AngleX = angleX;
            this.AngleY = angleY;
            this.AngleZ = angleZ;

            this.AxisX = axisX;
            this.AxisY = axisY;
            this.AxisZ = axisZ;
        }
    }

    internal abstract class OpenGlArtist
    {
        public static void DrawingSpheres(OpenGLControl openGLControl, List<MySphere> drawingSpheres, float angle)
        {
            OpenGL OpenGL = openGLControl.OpenGL;

            OpenGL.Clear(OpenGL.GL_COLOR_BUFFER_BIT | OpenGL.GL_DEPTH_BUFFER_BIT);
            OpenGL.LoadIdentity();
            GLColor glClr = new GLColor(0.0f, 0.0f, 1.0f, 0.1f);
            OpenGLAttributesEffect glEffect = new OpenGLAttributesEffect();
            glEffect.ColorBufferAttributes.ColorModeClearColor = glClr;
            glEffect.ColorBufferAttributes.ColorModeWriteMask = glClr;

            OpenGL.Translate(0.0f, 0.0f, -70.0f);
            OpenGL.Rotate(angle, Convert.ToInt32(drawingSpheres[0].AxisX) * 0.5f, Convert.ToInt32(drawingSpheres[0].AxisY) * 0.5f, Convert.ToInt32(drawingSpheres[0].AxisZ) * 0.5f);

            SharpGL.SceneGraph.Quadrics.Sphere[] spheres = new Sphere[drawingSpheres.Count];  
            for(int i = 0; i < spheres.Length; i++)
            {
                spheres[i] = new SharpGL.SceneGraph.Quadrics.Sphere();
               
                spheres[i].Radius = drawingSpheres[i].Radius;
                spheres[i].Slices = drawingSpheres[i].Slices;
                spheres[i].Stacks = drawingSpheres[i].Stacks;
                spheres[i].TextureCoords = true;
                spheres[i].QuadricDrawStyle = drawingSpheres[i].Style;

                spheres[i].Transformation.TranslateX = drawingSpheres[i].TranslationX;
                spheres[i].Transformation.TranslateY = drawingSpheres[i].TranslationY;
                //spheres[i].Transformation.TranslateZ = drawingSpheres[i].TranslationZ;

                spheres[i].Transformation.RotateX = drawingSpheres[i].AngleX;
                spheres[i].Transformation.RotateY = drawingSpheres[i].AngleY;
                spheres[i].Transformation.RotateZ = drawingSpheres[i].AngleZ;

                spheres[i].NormalGeneration = Normals.Smooth;
                spheres[i].NormalOrientation = SharpGL.SceneGraph.Quadrics.Orientation.Inside;
                OpenGL.Color(drawingSpheres[0].Color);

                spheres[i].AddEffect(glEffect);
                spheres[i].CreateInContext(OpenGL);
                spheres[i].PushObjectSpace(OpenGL);
                spheres[i].Render(OpenGL, SharpGL.SceneGraph.Core.RenderMode.Render);
                spheres[i].PopObjectSpace(OpenGL);
            }
            OpenGL.End();
            OpenGL.Flush();
        }
    }
}

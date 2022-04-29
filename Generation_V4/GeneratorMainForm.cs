using ExcelDataReader;
using System;
using System.Collections;
using System.Data;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Generation_V4
{
    public partial class GeneratorMainForm : Form
    {
        private string filename = string.Empty;//Путь к файлу Excel
        private string filename2 = string.Empty;//Путь к файлу Word
        private string PathFolder = "";
        public DataTableCollection tableCollection = null;
        public DataTable Table1 = null;// инициализируем таблицу данных, для первого листа Excel, в памяти
        public DataTable Table2 = null;// инициализируем таблицу данных, для второго листа Excel, в памяти
        int CountColumnsTable1 = 0;//инициализируем количество столбцов в первом листе Excel
        int CountRowTable1 = 0;//инициализируем количество строк в первом листе Excel
        int CountColumnsTable2 = 0;//инициализируем количество столбцов во втором листе Excel
        int CountRowTable2 = 0;//инициализируем количество строк во втором листе Excel

        public GeneratorMainForm()
        {
            InitializeComponent();
        }

        void GenerationButtonCheked(object sender, EventArgs e)//Включение кнопки Генерация если 3 поля заполнены
        {
            if (textBox1SelectWord.Text != "")
                if (textBoxSelectExcel.Text != "")
                    if (textBoxSelectPathSave.Text != "")
                        Generation.Enabled = true;
        }

        private void SelectWord_Click(object sender, EventArgs e)//Выбрать и прочитать шаблон Word
        {
            try
            {
                DialogResult res = openFileDialog2.ShowDialog();
                if (res == DialogResult.OK)
                {
                    filename2 = openFileDialog2.FileName;
                    textBox1SelectWord.Text = filename2;
                    Word.Application app = new Word.Application
                    {
                        Visible = false
                    };
                    Object missing = Type.Missing;
                    app.Documents.Open(filename2);
                    //
                    //Получение имени закладки и создание соответствующе именованных чекбоксов
                    //
                    checkedListBox1.Items.Clear();
                    for (int NumberBookmark = 1; NumberBookmark <= app.ActiveDocument.Bookmarks.Count; NumberBookmark++)
                    {
                        checkedListBox1.Items.Add(app.ActiveDocument.Bookmarks[NumberBookmark].Name);
                    }
                    //
                    //Получение количества таблиц в документе и создание чекбоксов
                    //
                    checkedListBox2.Items.Clear();
                    for (int NumberTable = 1; NumberTable <= app.ActiveDocument.Tables.Count; NumberTable++)
                    {
                        checkedListBox2.Items.Add("Таблица " + NumberTable + " " + app.ActiveDocument.Tables[NumberTable].Title);
                    }

                    app.Documents.Close();
                    app.Quit();
                }

                else
                {
                    throw new Exception("Файл не выбран");
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Закройте шаблон документа и процесс в диспетчере", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OpenExcelFile(string path)//Чтение файла Excel 
        {
            try
            {
                FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);
                IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
                DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });
                //
                //=================================================================
                //
                try
                {
                    toolStripComboBox2.Items.Clear();                    
                    toolStripComboBox1.Items.Clear();
                    tableCollection = db.Tables;

                    if (db.Tables.Count > 0)
                    {
                        foreach (DataTable Table1 in tableCollection)
                        {
                            toolStripComboBox1.Items.Add(Table1.TableName);
                        }
                        toolStripComboBox1.SelectedIndex = 0;
                    }
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "В Excel нет листа номер 1", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                //
                //=================================================================
                //
                try
                {
                    toolStripComboBox2.Items.Clear();
                    tableCollection = db.Tables;                    

                    if (db.Tables.Count > 1)
                    {
                        foreach (DataTable Table1 in tableCollection)
                        {
                            toolStripComboBox2.Items.Add(Table1.TableName);
                        }

                        toolStripComboBox2.SelectedIndex = 1;
                    }
                    if (db.Tables.Count < 2)
                    {
                        toolStripComboBox2.Items.Clear();
                        toolStripComboBox2.Items.Add("");
                        toolStripComboBox2.SelectedIndex = 0;
                    }
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "В Excel нет листа номер 2", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                reader.Close();
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка в блоке OpenExcelFile", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void SelectExcel_Click(object sender, EventArgs e)//Открыть файл Excel 
        {
            try
            {
                DialogResult res = openFileDialog1.ShowDialog();
                if (res == DialogResult.OK)
                {
                    filename = openFileDialog1.FileName;
                    textBoxSelectExcel.Text = filename.ToString();
                    OpenExcelFile(filename);
                }
                else
                {
                    throw new Exception("Файл не выбран");
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка в блоке SelectExcel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SelectPathSave_Click(object sender, EventArgs e)// Обработка выбора пути сохранения готовых файлов
        {
            try
            {
                FolderBrowserDialog FBD = new FolderBrowserDialog
                {
                    ShowNewFolderButton = false
                };
                if (FBD.ShowDialog() == DialogResult.OK)

                {
                    textBoxSelectPathSave.Text = PathFolder = FBD.SelectedPath;                    
                }
                else
                {
                    throw new Exception("Директория куда сохранить файлы не выбрана");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка в блоке SelectPathSave", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ToolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)// Обработка значений из 1-го листа Exсel
        {
            try
            {

                Table1 = tableCollection[Convert.ToString(toolStripComboBox1.SelectedItem)];
                if (Table1.Rows.Count > 0)
                {
                    {
                        CountColumnsTable1 = Table1.Columns.Count;
                        CountRowTable1 = Table1.Rows.Count;
                    }

                    comboBox1.Items.Clear();
                    for (int NumberColumnTable1 = 0; NumberColumnTable1 < Table1.Columns.Count; NumberColumnTable1++)
                    {
                        comboBox1.Items.Add(Table1.Columns[NumberColumnTable1].ColumnName);
                    }

                    LblStatus.Text = "Будет создано " + Table1.Rows.Count + " комплект(a)(ов) документов";
                    comboBox1.SelectedIndex = 0;
                }

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка4", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ToolStripComboBox2_SelectedIndexChanged(object sender, EventArgs e)// Обработка значений из 2-го листа Exсel
        {
            try
            {
                Table2 = tableCollection[Convert.ToString(toolStripComboBox2.SelectedItem)];
                if (Table2 != null)
                {
                    CountColumnsTable2 = Table2.Columns.Count;
                    CountRowTable2 = Table2.Rows.Count;
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка в блоке ToolStripComboBox2", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Generation_Click(object sender, EventArgs e)// Генерация актов
        {
            //
            //Инициализация переменной для счётчика итераций
            //
            int Counter = 0;

            ProgressBar.Maximum = Table1.Rows.Count;

            Word.Application app = new Word.Application();
            //app.Visible = true;

            try
            {
                for (int NumberRowSheet1Excel = 0; NumberRowSheet1Excel < CountRowTable1; NumberRowSheet1Excel++)
                {
                    //
                    //Создаём листы в которых будут храниться имена всех столбцов (наименование переменных)
                    //
                    ArrayList Variables = new ArrayList();
                    if (toolStripComboBox1.Text != "")
                    {
                        for (int i = 0; i < CountColumnsTable1; i++)
                        {
                            Variables.Add(Table1.Columns[i].ColumnName.ToString());
                        }
                    }

                    ArrayList ListVariables = new ArrayList();
                    if (toolStripComboBox2.Text != "")
                    {
                        for (int i = 0; i < CountColumnsTable2; i++)
                        {
                            ListVariables.Add(Table2.Columns[i].ColumnName.ToString());
                        }
                    }
                    //
                    //Задаём имя новым файлам
                    //                    
                    object fileNameEkz1Docx = PathFolder + "\\" + "Экз №1 " + comboBox1.Text + Table1.Rows[NumberRowSheet1Excel][comboBox1.SelectedIndex].ToString() + ".docx";
                    object fileNameEkz2Docx = PathFolder + "\\" + "Экз №2 " + comboBox1.Text + Table1.Rows[NumberRowSheet1Excel][comboBox1.SelectedIndex].ToString() + ".docx";
                    object fileNameEkz3Docx = PathFolder + "\\" + "Экз №3 " + comboBox1.Text + Table1.Rows[NumberRowSheet1Excel][comboBox1.SelectedIndex].ToString() + ".docx";

                    object fileNameEkz1Pdf = PathFolder + "\\" + "Экз №1 " + comboBox1.Text + Table1.Rows[NumberRowSheet1Excel][comboBox1.SelectedIndex].ToString() + ".pdf";
                    object fileNameEkz2Pdf = PathFolder + "\\" + "Экз №2 " + comboBox1.Text + Table1.Rows[NumberRowSheet1Excel][comboBox1.SelectedIndex].ToString() + ".pdf";
                    object fileNameEkz3Pdf = PathFolder + "\\" + "Экз №3 " + comboBox1.Text + Table1.Rows[NumberRowSheet1Excel][comboBox1.SelectedIndex].ToString() + ".pdf";


                    object oMissing = System.Reflection.Missing.Value;
                    object oEndOfDoc = "\\endofdoc"; /* \endofdoc это предопределенная закладка */
                    //
                    //Путь до файла шаблона Word
                    //                    
                    Object missing = Type.Missing;
                    app.Documents.Open(textBox1SelectWord.Text);
                    //
                    //Удаление не отмеченных закладок
                    //
                    for (int IndexSavedBookmark = 0; IndexSavedBookmark < checkedListBox1.Items.Count; IndexSavedBookmark++)
                    {
                        if (checkedListBox1.GetItemChecked(IndexSavedBookmark) == false)
                        {
                            string BookmarkName = checkedListBox1.Items[IndexSavedBookmark].ToString();
                            app.ActiveDocument.Bookmarks[BookmarkName].Range.Delete();
                        }
                    }
                    //
                    //Получаем номер таблицы в открытом документе WORD и отмеченную в checkedListBox2,
                    //которую нужно почистить от пустых строк и проверяем построчно ячейки 2, 3, 4
                    // если три ячейки подряд в строке пустые удаляем строку
                    //
                    int NumberCheckTable = checkedListBox2.Items.Count;

                    for (int IndexEditTable = 0; IndexEditTable < NumberCheckTable; IndexEditTable++)
                    {
                        if (checkedListBox2.GetItemChecked(IndexEditTable) == true)
                        {
                            Word.Table TableWord = app.ActiveDocument.Tables[IndexEditTable + 1];
                            int NumberRowsTable = TableWord.Rows.Count;

                            for (int z = NumberRowsTable; z > 0; z--)
                            {
                                int NumberCellTable = TableWord.Rows[z].Cells.Count;

                                if ((NumberCellTable > 4) == true)
                                {

                                    if (string.IsNullOrEmpty(TableWord.Rows[z].Cells[2].Range.Text.Replace("\r\a", "").Trim()) == true)
                                    {

                                        if (string.IsNullOrEmpty(TableWord.Rows[z].Cells[3].Range.Text.Replace("\r\a", "").Trim()) == true)
                                        {

                                            if (string.IsNullOrEmpty(TableWord.Rows[z].Cells[4].Range.Text.Replace("\r\a", "").Trim()) == true)
                                            {
                                                TableWord.Rows[z].Delete();
                                            }

                                        }

                                    }

                                }

                            }

                        }

                    }
                    //
                    //Срост данных таблицы и шаблона документа Word
                    //
                    object replace = Word.WdReplace.wdReplaceAll;
                    Object wrap = Word.WdFindWrap.wdFindContinue;
                    Word.Find find = app.Selection.Find;
                    object fileformat = Word.WdSaveFormat.wdFormatPDF;
                    int SectionCount = app.ActiveDocument.Sections.Count;
                    //
                    //Блок 1 поиск и замена одиночных переменных по документу и колонтитулам
                    //
                    {
                        for (int j = 0; j < CountColumnsTable1; j++)
                        {
                            find.Text = "{$" + (string)Variables[j] + "$}";// что меняем, переменные в шаблоне
                            find.Replacement.Text = Table1.Rows[NumberRowSheet1Excel][(string)Variables[j]].ToString();// на что меняем, значение переменных из Excel
                            find.Execute(FindText: Type.Missing, Wrap: wrap, ReplaceWith: missing, Replace: replace);
                            //
                            //Замена переменных в нижних колонтитулах
                            //блок кода увеличивает время выполнения программы в 2.5 раза
                            //
                            #region
                            if (CheckColontitul.Checked)
                            {
                                object FindTextFooter = "{$" + (string)Variables[j] + "$}";// что меняем
                                object ReplaceWithFooter = Table1.Rows[NumberRowSheet1Excel][(string)Variables[j]].ToString(); // на что меняем
                                for (int i = 1; i <= SectionCount; i++)
                                {
                                    app.ActiveDocument.Sections[i].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Find.Execute
                                        (FindText: FindTextFooter, ReplaceWith: ReplaceWithFooter, Replace: replace);
                                    app.ActiveDocument.Sections[i].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Find.Execute
                                        (FindText: FindTextFooter, ReplaceWith: ReplaceWithFooter, Replace: replace);
                                    app.ActiveDocument.Sections[i].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Find.Execute
                                        (FindText: FindTextFooter, ReplaceWith: ReplaceWithFooter, Replace: replace);
                                }
                            }
                            #endregion
                        }
                    }

                    if (FixedColontitul.Checked)
                    {
                        string ReplaceWithFooter = "Акт № " + Table1.Rows[NumberRowSheet1Excel][(string)Variables[2]].ToString() + "/ДИТ–" +
                            Table1.Rows[NumberRowSheet1Excel][(string)Variables[3]].ToString() + "/" +
                            Table1.Rows[NumberRowSheet1Excel][(string)Variables[0]].ToString() + "\n" +
                            Table1.Rows[NumberRowSheet1Excel][(string)Variables[1]].ToString();

                        for (int i = 1; i <= SectionCount; i++)
                        {
                            app.ActiveDocument.Sections[i].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = ReplaceWithFooter;
                            app.ActiveDocument.Sections[i].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = ReplaceWithFooter;
                            app.ActiveDocument.Sections[i].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Text = ReplaceWithFooter;
                        }
                    }
                    //
                    //Блок 2 Поиск и замена списочных переменных по документу
                    //
                    if (toolStripComboBox2.Text != "")
                    {
                        for (int CellNumber = 0; CellNumber < CountColumnsTable2; CellNumber++)
                        {
                            find.Text = "[$" + (string)ListVariables[CellNumber] + "$]";// что меняем
                            //
                            //Земена одной переменной в таблице WORD на несколько значений со 2 листа Excel
                            //Ограничение string и метода Find 255 символов, циклически заменяем переменную в WORD и
                            //обходим ограничение в 255 символов путем дописывания этой же переменной после её замены
                            //
                            ArrayList ListVariablesFromExcel = new ArrayList();
                            for (int NumberRow = 0; NumberRow < CountRowTable2; NumberRow++)
                            {
                                if (Table2.Rows[NumberRow][CellNumber].ToString() != "")
                                    ListVariablesFromExcel.Add(Table2.Rows[NumberRow][CellNumber].ToString());
                            }

                            int SizeArraySpisok = ListVariablesFromExcel.Count - 1;
                            for (int a = 0; a <= SizeArraySpisok; a++)
                            {
                                if (a < SizeArraySpisok)
                                    find.Replacement.Text = (string)ListVariablesFromExcel[a] + "^p" + "[$" + (string)ListVariables[CellNumber] + "$]";// на что меняем
                                else if (a == SizeArraySpisok)
                                {
                                    find.Replacement.Text = (string)ListVariablesFromExcel[a];
                                }
                                find.Execute(FindText: Type.Missing, Wrap: wrap, ReplaceWith: missing, Replace: replace);
                            }
                        }
                    }

                    if (app.ActiveDocument.Comments.Count > 0)
                    {
                        app.Application.ActiveDocument.DeleteAllComments();
                    }
                    app.ActiveDocument.AcceptAllRevisions();
                    //
                    //Сохранение в выбранном формате и количестве экземпляров с установкой номера экземпляра в колонтитуле
                    //
                    //=====================================================================================================
                    if (((int)numericUpDown1.Value == 1) == true)// Если один экземпляр
                    {
                        object FindTextHeaders = "Экз"; // что меняем
                        object ReplaceWithHeaders = "Экз. №1"; // на что меняем

                        for (int i = 1; i <= SectionCount; i++)
                        {
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Find.Execute
                                (FindText: FindTextHeaders, ReplaceWith: ReplaceWithHeaders, Replace: replace);
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Find.Execute
                                (FindText: FindTextHeaders, ReplaceWith: ReplaceWithHeaders, Replace: replace);
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Find.Execute
                                (FindText: FindTextHeaders, ReplaceWith: ReplaceWithHeaders, Replace: replace);
                        }
                        app.ActiveDocument.SaveAs2(ref fileNameEkz1Docx);

                        if (CheckBoxSaveToPdf.Checked)
                            app.ActiveDocument.SaveAs2(ref fileNameEkz1Pdf, fileformat);//Сохраняем в формате PDF

                        Counter++;
                        ProgressBar.Value = Counter;
                        ProgressBar.Update();
                        LblStatus.Text = "Выполнено " + Counter + " из " + CountRowTable1;
                    }
                    //=====================================================================================================
                    if (((int)numericUpDown1.Value == 2) == true)// Если два экземпляра
                    {
                        object FindTextHeaders = "Экз"; // что меняем
                        object ReplaceWithHeaders = "Экз. №1"; // на что меняем
                        for (int i = 1; i <= SectionCount; i++)
                        {
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Find.Execute
                                (FindText: FindTextHeaders, ReplaceWith: ReplaceWithHeaders, Replace: replace);
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Find.Execute
                                (FindText: FindTextHeaders, ReplaceWith: ReplaceWithHeaders, Replace: replace);
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Find.Execute
                                (FindText: FindTextHeaders, ReplaceWith: ReplaceWithHeaders, Replace: replace);
                        }

                        app.ActiveDocument.SaveAs2(ref fileNameEkz1Docx);

                        if (CheckBoxSaveToPdf.Checked)
                            app.ActiveDocument.SaveAs2(ref fileNameEkz1Pdf, fileformat);//Сохраняем в формате PDF

                        object FindTextHeaders2 = "Экз. №1"; // что меняем
                        object ReplaceWithHeaders2 = "Экз. №2"; // на что меняем

                        for (int i = 1; i <= SectionCount; i++)
                        {
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Find.Execute
                                (FindText: FindTextHeaders2, ReplaceWith: ReplaceWithHeaders2, Replace: replace);
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Find.Execute
                                (FindText: FindTextHeaders2, ReplaceWith: ReplaceWithHeaders2, Replace: replace);
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Find.Execute
                                (FindText: FindTextHeaders2, ReplaceWith: ReplaceWithHeaders2, Replace: replace);
                        }

                        app.ActiveDocument.SaveAs2(ref fileNameEkz2Docx);
                        if (CheckBoxSaveToPdf.Checked)
                            app.ActiveDocument.SaveAs2(ref fileNameEkz2Pdf, fileformat);//Сохраняем в формате PDF
                        Counter++;
                        ProgressBar.Value = Counter;
                        ProgressBar.Update();
                        LblStatus.Text = "Выполнено " + Counter + " из " + CountRowTable1;
                    }
                    //=====================================================================================================
                    if (((int)numericUpDown1.Value == 3) == true)// Если три экземпляра
                    {
                        object FindTextHeaders = "Экз"; // что меняем
                        object ReplaceWithHeaders = "Экз. №1"; // на что меняем
                        for (int i = 1; i <= SectionCount; i++)
                        {
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Find.Execute
                                (FindText: FindTextHeaders, ReplaceWith: ReplaceWithHeaders, Replace: replace);
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Find.Execute
                                (FindText: FindTextHeaders, ReplaceWith: ReplaceWithHeaders, Replace: replace);
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Find.Execute
                                (FindText: FindTextHeaders, ReplaceWith: ReplaceWithHeaders, Replace: replace);
                        }

                        app.ActiveDocument.SaveAs2(ref fileNameEkz1Docx);

                        if (CheckBoxSaveToPdf.Checked)
                            app.ActiveDocument.SaveAs2(ref fileNameEkz1Pdf, fileformat);//Сохраняем в формате PDF

                        object FindTextHeaders2 = "Экз. №1"; // что меняем
                        object ReplaceWithHeaders2 = "Экз. №2"; // на что меняем

                        for (int i = 1; i <= SectionCount; i++)
                        {
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Find.Execute
                                (FindText: FindTextHeaders2, ReplaceWith: ReplaceWithHeaders2, Replace: replace);
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Find.Execute
                                (FindText: FindTextHeaders2, ReplaceWith: ReplaceWithHeaders2, Replace: replace);
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Find.Execute
                                (FindText: FindTextHeaders2, ReplaceWith: ReplaceWithHeaders2, Replace: replace);
                        }

                        app.ActiveDocument.SaveAs2(ref fileNameEkz2Docx);
                        if (CheckBoxSaveToPdf.Checked)
                            app.ActiveDocument.SaveAs2(ref fileNameEkz2Pdf, fileformat);//Сохраняем в формате PDF


                        object FindTextHeaders3 = "Экз. №2"; // что меняем
                        object ReplaceWithHeaders3 = "Экз. №3"; // на что меняем

                        for (int i = 1; i <= SectionCount; i++)
                        {
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Find.Execute
                                (FindText: FindTextHeaders3, ReplaceWith: ReplaceWithHeaders3, Replace: replace);
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Find.Execute
                                (FindText: FindTextHeaders3, ReplaceWith: ReplaceWithHeaders3, Replace: replace);
                            app.ActiveDocument.Sections[i].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Find.Execute
                                (FindText: FindTextHeaders3, ReplaceWith: ReplaceWithHeaders3, Replace: replace);
                        }

                        app.ActiveDocument.SaveAs2(ref fileNameEkz3Docx);
                        if (CheckBoxSaveToPdf.Checked)
                            app.ActiveDocument.SaveAs2(ref fileNameEkz3Pdf, fileformat);//Сохраняем в формате PDF

                        //=====================================================================================================

                        Counter++;
                        ProgressBar.Value = Counter;
                        ProgressBar.Update();
                        LblStatus.Text = "Выполнено " + Counter + " из " + CountRowTable1;
                    }

                }

                MessageBox.Show("Готовые файлы находятся " + PathFolder);
                ProgressBar.Value = 0;
                LblStatus.Text = "Processing....";

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка  в блоке Generation", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app?.ActiveDocument.Close(SaveChanges: 0);
                app?.Quit(SaveChanges: 0);
            }
        }

        private void CheckColontitul_CheckedChanged(object sender, EventArgs e)//Проверка способа создания колонтитулов из переменных
        {
            if (CheckColontitul.Checked)
                FixedColontitul.Checked = false;
        }

        private void FixedColontitul_CheckedChanged(object sender, EventArgs e)//Проверка способа задания фиксироваанных колонтитулов
        {
            if (FixedColontitul.Checked)
                CheckColontitul.Checked = false;
        }        

        private void ОпрограммеToolStripMenuItem_Click(object sender, EventArgs e)//Вызов формы описания программы
        {
            About about = new About();
            about.ShowDialog();
        }

        private void Exit_Click(object sender, EventArgs e)// Завершение работы приложения
        {
            Application.Exit();
        }
    }
}
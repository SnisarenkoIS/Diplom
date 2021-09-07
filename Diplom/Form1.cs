using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using _CSharp;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using Xceed.Words.NET;
using Xceed.Document.NET;
using Xceed.Utils.Exceptions;
using System.Threading;

namespace Diplom
{
    public partial class Form1 : Form
    {
        //***Test Word
        public static Word.Application wordapp; //--- Для работы с Вордовскими документами
        public static Word.Document worddoc_s;
        public static Word.Document worddoc;
        public static Word.Paragraphs wordparagraph_s;
        public static Word.Paragraph wordparagraph;
        //************

        public static HashSet<string> Set_files_describing_the_program = new HashSet<string>(); //--- Множество, в которое записываются пути/путь до файлов/файла с описанием программы
        public static List<string> List_exception_for_false_issuance_of_illegitimate_files = new List<string>(); //--- Список файлов, для дальнейшего исключения нелегитимности подключаемых (заголовочных) файлов
        public static Dictionary<string, int> Dic_test = new Dictionary<string, int>(); //---Test

        public static bool Flag_converge = false; //--- Флаг, сигнализирующий о том, что описанные в документе по проверяемому приложению файлы с исходниками СОВПАДАЮТ с файлами из самого проекта
        public static bool Flag_not_converge = false; //--- Флаг, сигнализирующий о том, что описанные в документе по проверяемому приложению файлы с исходниками  НЕ СОВПАДАЮТ с файлами из самого проекта

        public static bool Flag_text_files_in_new_folder = false; //--- Флаг, сигнализирующий, что все текстовые файлы из проверяемой директории занести в новую папку

        public static bool Flag_btn_open_dir_enter = false; //--- Флаг, сигнализирующий о том, что дериктория с приложением была открыта
        public static bool Flag_image_hex = false; //--- Флаг, сигнализирующий, что картинка будет открываться в формате Hex
        public static bool Flag_image_original = false; //--- Флаг, сигнализирующий, что картинка будет открываться в своём формате
        public static bool Flag_remove_dash = false; //--- Флаг, сигнализирующий, что разделитель (тире) между байтами в хексовом представлении картинки удаляем
        public static bool Flag_Not_remove_dash = false; //--- Флаг, сигнализирующий, что разделитель (тире) между байтами в хексовом представлении картинки НЕ удаляем
        public static bool Flag_selected_FIO = false; //--- Флаг, сигнализирующий, что ФИО сотрудника введено
        public static bool Flag_enter_Reset = false; //--- Флаг, сигнализирующий, что была нажата кнопка "Сброс"
        public static bool Flag_choice_radiobatton_non_ligitim = false; //--- Флаг, сигнализирующий, что в отчёт выводится ТОЛЬКО негелитимные файлы
        public static bool Flag_choice_radiobatton_indescribable = false; //--- Флаг, сигнализирующий, что в отчёт выводится ТОЛЬКО неописанные файлы
        public static bool Flag_choice_radiobatton_non_ligitim_and_indescribable = false; //--- Флаг, сигнализирующий, что в отчёт выводится негелитимные и неописанные файлы

        public static string[] pth_indescribable_exe; //--- Массив путей до несоответствующих файлов (исполняемых)
        public static string[] pth_indescribable_bin; //--- Массив путей до несоответствующих файлов (бинарных)
        public static string[] pth_non_legitimace_Dic;
        public static string[] pth_non_legitimace_Set;

        public static string name_of_the_audited_directory = "";
        string FIO = "";

        int count_reports_indescribable = 1;
        int count_reports_non_legitimace = 1;

        public static string st_pth_dir = ""; //--- Путь до проверяемой директории, который вводится пользователем в ТекстБокс


        object sender_1;
        DrawItemEventArgs e_1;
        

        public Form1()
        {
            InitializeComponent();

            btn_check_description_and_files.Enabled = false;
            btn_check_connecting_header.Enabled = false;
            groupBox4.Enabled = false;
            groupBox11.Visible = false;
            rbn_image_original.Checked = true;
            btn_start_inspect.Enabled = false;
            btn_Reset.Enabled = false;
            groupBox2.Enabled = false;
            label6.Enabled = false;
            lbx_list_all_files.Enabled = false;
            lbx_False_header.Visible = false;
            lbl_False_header.Visible = false;
            label28.Visible = false;
            label30.Visible = false;
            rbn_non_ligitimate_and_indescribable.Checked = true;
            btn_generate_a_report.Visible = false;
            listBox1.Visible = false;
            chbx_Text_file_in_separate_folder.Checked = false;

            //--- Если файл с сохранёнными путями до сторонних приложений существует
            if (File.Exists(Directory.GetCurrentDirectory() + "\\SavePath.txt"))
            {
                fini = new FileStream(Directory.GetCurrentDirectory() + "\\SavePath.txt", FileMode.Open, FileAccess.Read);
                string[] mass_pth_exe = File.ReadAllLines(fini.Name);

                for (int i = 0; i < 4; i++)
                {
                    switch (i)
                    {
                        case 0:
                            if (mass_pth_exe[i].Length > 4)
                            {
                                path_txt_exe = mass_pth_exe[i];
                                tbx_txt_exe.Text = path_txt_exe;
                            }
                            else
                            {
                                path_txt_exe = "";
                                tbx_txt_exe.Text = "";
                            }
                            break;
                        case 1:
                            if (mass_pth_exe[i].Length > 4)
                            {
                                path_bin_exe = mass_pth_exe[i];
                                tbx_bin_exe.Text = path_bin_exe;
                            }
                            else
                            {
                                path_bin_exe = "";
                                tbx_bin_exe.Text = "";
                            }
                            break;
                        case 2:
                            if (mass_pth_exe[i].Length > 4)
                            {
                                path_image_exe = mass_pth_exe[i];
                                tbx_image_exe.Text = path_image_exe;
                            }
                            else
                            {
                                path_image_exe = "";
                                tbx_image_exe.Text = "";
                            }
                            break;
                        case 3:
                            if (mass_pth_exe[i].Length > 4)
                            {
                                path_source_code_exe = mass_pth_exe[i];
                                tbx_source_code_exe.Text = path_source_code_exe;
                            }
                            else
                            {
                                path_source_code_exe = "";
                                tbx_source_code_exe.Text = "";
                            }
                            break;
                    }
                }

                fini.Close();
            }
        }



        Process proc = new Process(); //--- Для открытия сторонних программ из своей (для быстрого просмотра файлов)

        //---Пути до программ, реализующих быстрое открытие файла из ЛистБокса
        //string pth_notepad = "C:\\Windows\\system32\\notepad.exe";

        int count_YP = 0; //--- Кол-во файлов на иных ЯП (кроме C/C++ и CS)
        int count_check_files = 0; //--- Кол-во файлов на проверку эксперту
        int count_check_bin_files = 0; //--- Кол-во бинарных файлов на проверку
        int count_check_connecting_header = 0; //--- Кол-во заголовочных файлов на проверку
        int count_file_for_analysis = 0; //--- Кол-во
        int count_file = 0; //--- Счётчик файлов в деректории
        double size_dir = 0; //--- Размер полученной директории
        //string current_file = ""; //--- Путь до файла, находящегося в данный момент на проверке

        public static bool Flag_selected_dir = false; //--- Флаг, сигнализирующий, что путь до проверяемой директории внесён в поле
        public static bool Flag_Not_zero_all_file = false;
        public static bool Flag_zero_tbx_txt_exe = false; //--- Флаг, сигнализирующий, что ТекстБокс текстового приложения пуст
        public static bool Flag_zero_tbx_bin_exe = false; //--- Флаг, сигнализирующий, что ТекстБокс бинарного приложения пуст
        public static bool Flag_zero_tbx_image_exe = false; //--- Флаг, сигнализирующий, что ТекстБокс графического приложения пуст
        public static bool Flag_start_inspect = false; //--- Флаг, сигнализирующий, что анализ проверяемой исзодной директории проведён

        //--- Множество расширений, файлы которых будут автоматически помечаться как "Файлы к проверке"
        HashSet<String> LotsExtension = new HashSet<String> { ".c", ".cpp", ".cs", ".h", ".hpp", ".html", ".py", ".exe", ".dll", ".lib", ".java", ".js", ".php", ".sql" };

        HashSet<String> TextExe = new HashSet<String> {".txt", ".doc", ".docx" };
        HashSet<String> SourceCodeExe = new HashSet<String> { ".c", ".cpp", ".cs", ".h", ".hpp", ".html", ".py" };
        HashSet<String> ImageExe = new HashSet<String> { ".bmp", ".gif", ".jpeg", ".jpg" };
        HashSet<String> BinExe = new HashSet<String> { ".exe", ".dll", ".lib", ".bin", ".hex" };

        public static string path_dir; //--- Путь до нужной нам деректории
        string path_description; //--- Путь до файла с описанием программы (для реализации сверки)
        string path_txt_exe; //--- Путь до приложения, обеспечивающего просмотр текстовых файлов
        string path_bin_exe; //--- Путь до приложения, обеспечивающего просмотр бинарных файлов
        string path_image_exe; //--- Путь до приложения, обеспечивающего просмотр графических файлов
        string path_source_code_exe; //--- Путь до приложения, обеспечивающего просмотр файлов с исходным кодом

        FileStream fini; //--- Переменная для сохранения путей до сторонних приложений
        FileStream f_txt; //--- Переменная для анализа текстовых файлов на наличие фрагментов кода

        bool Add_txt_file;

        //
        //***
        //
        public static Dictionary<int, string> new_pth_folder = new Dictionary<int, string>(); //--- словарь <index - новый путь>
        private void btn_open_file_Click(object sender, EventArgs e) //--- 
        {
            folderBrowserDialog1 = new FolderBrowserDialog();
            folderBrowserDialog1.ShowNewFolderButton = false;

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK) path_dir = folderBrowserDialog1.SelectedPath;

            if (folderBrowserDialog1.SelectedPath == String.Empty) return;

            Flag_converge = false;
            Flag_not_converge = false;

            lbx_list_all_files.Items.Clear();
            lbx_list_bin_files.Items.Clear();
            lbx_list_check_files.Items.Clear();

            GetAllFile(path_dir, lbx_list_all_files);

            int count_tmp = 0;
            if (Flag_text_files_in_new_folder)
            {
                new_pth_folder.Clear();

                string path_new_folder_with_test_files = "";
                foreach (var t in lbx_list_all_files.Items)
                {
                    FileInfo tmp_fi = new FileInfo(t.ToString());
                    string fname = tmp_fi.Name;                 
                    int index_elem = 0;

                    if (tmp_fi.Extension == ".doc" || tmp_fi.Extension == ".docx" || tmp_fi.Extension == ".txt") //--- Если найден текстовый файл
                    {
                        index_elem = lbx_list_all_files.Items.IndexOf(t);
                        count_tmp++;
                        if (count_tmp == 1)
                        {
                            path_new_folder_with_test_files = path_dir + "\\Текстовые файлы из проверяемой директории"; //---- Получаем путь до папки, из которой запускаемся
                            if (!Directory.Exists(path_new_folder_with_test_files)) Directory.CreateDirectory(path_new_folder_with_test_files); //----- Если искомой папки нет - создаем
                        }

                        File.Move(t.ToString(), path_new_folder_with_test_files + "\\" + fname); //--- Перемещаем файл в новую папку
                        new_pth_folder[index_elem] = path_new_folder_with_test_files + "\\" + fname; //--- Заносим индекс в Листбоксе перемещённого элемента с новым путём
                    }
                }

                if (count_tmp > 0) //--- Если хотябы раз перемещали текстовый файл
                {
                    foreach (var item in new_pth_folder) lbx_list_all_files.Items[item.Key] = item.Value; //--- Проходим по словарю. Меняем в основном Листбоксе путь до перемещённого текстового файла по конкретному индексу
                    MessageBox.Show("Все текстовые файлы, находящиеся в проверяемой директории, были перемещены в новую папку \"Текстовые файлы из проверяемой директории\", которая находится в корневой директории проверяемого проекта.");
                }
            }
                

            FileInfo fi = new FileInfo(path_dir);

            name_of_the_audited_directory = fi.Name;

            lbl_name_dir.Text = name_of_the_audited_directory;
            lbl_d_size.Text = Convert.ToString(sizeOfDirectories(path_dir, ref size_dir));
            lbl_date_in.Text = fi.CreationTime.ToShortDateString();
            lbl_time_in.Text = fi.CreationTime.ToShortTimeString();
            lbl_count_file.Text = Convert.ToString(count_file);
            count_file_for_analysis = count_file;

            if (count_file_for_analysis > 0) Flag_Not_zero_all_file = true;
            groupBox4.Enabled = true;
            btn_start_inspect.Enabled = true;
            btn_Reset.Enabled = true;
            label6.Enabled = true;
            lbx_list_all_files.Enabled = true;

            count_file = 0;
            btn_open_file.Enabled = false;
            btn_open_description.Enabled = true;
            chbx_Text_file_in_separate_folder.Enabled = false;
        }

        //
        //***
        //
        private void btn_start_inspect_Click(object sender, EventArgs e)
        {
            CheckFile(lbx_list_all_files, lbx_list_check_files, lbx_list_bin_files, lbx_list_YP); //--- Вызывается функция по проверке файлов
            lbl_count_check_file.Text = Convert.ToString(count_check_files);
            lbl_bin_files.Text = Convert.ToString(count_check_bin_files);
            lbl_connect_header.Text = Convert.ToString(count_check_connecting_header);
            lbl_count_YP.Text = Convert.ToString(count_YP);

            Flag_start_inspect = true;
            btn_check_connecting_header.Enabled = true;
            groupBox2.Enabled = true;
            btn_start_inspect.Enabled = false;
        }


        //
        //*** Функция для поиска файлов (и папок) в деректории и вывод их в листБокс (Работает!!!!!!!)
        //
        void GetAllFile(string startDirectory,ListBox filess) 
        {
            string[] searchdirectory = Directory.GetDirectories(startDirectory);
            if (searchdirectory.Length > 0)
                for (int i = 0; i < searchdirectory.Length; i++) GetAllFile(searchdirectory[i] + @"\", filess);

            string [] filesss = Directory.GetFiles(startDirectory);
            for (int i = 0; i < filesss.Length; i++)
            { 
                filess.Items.Add(filesss[i]);
                count_file++;
            }
        }

        //
        //*** Функция проверки файла на потенциальное наличие фрагментов исполняемого кода Convert
        //
        public void CheckFile(ListBox allF, ListBox CheckF, ListBox b_f, ListBox YP)
        { 
            //--- Будем проходить по всему изначальному списку всех файлов
            for (int i = 0; i < count_file_for_analysis; i++)
            {
                FileInfo fi = new FileInfo(Convert.ToString(allF.Items[i]));
                string extension = fi.Extension; //--- Получаем расширение файла

                if (LotsExtension.Contains(extension)) //--- Если бинарный файл или файл с исполняемым кодом
                {
                    if (extension == ".dll" || extension == ".exe" || extension == ".lib")
                    {
                        b_f.Items.Add(allF.Items[i]);
                        count_check_bin_files++;
                    }
                    else if (extension == ".h" || extension == ".hpp")
                    {
                        //--- test
                        if (!Dic_test.ContainsKey(fi.Name)) Dic_test[fi.Name] = 1;
                        else
                        {
                            int temp = Dic_test[fi.Name];
                            temp++;
                            Dic_test[fi.Name] = temp;
                        }
                        //---

                        lbx_check_connecting_header.Items.Add(allF.Items[i]);
                        count_check_connecting_header++;
                    }
                    else if (extension == ".html" || extension == ".py" || extension == ".java" || extension == ".js" || extension == ".php" || extension == ".sql")
                    {
                        YP.Items.Add(allF.Items[i]);
                        count_YP++;
                    }
                    else
                    {
                        CheckF.Items.Add(allF.Items[i]);
                        count_check_files++;
                    }
                }
                else if (TextExe.Contains(extension)) //--- Если текстовый файл
                {
                    try
                    {
                        if (extension == ".doc")
                        {
                            wordapp = new Word.Application(); //--- Создали объект Ворд (открыли(запустили) Ворд)
                            Add_txt_file = ClassCSharp.Check_CS_InDocFile(wordapp, fi.FullName);
                            if (Add_txt_file) CheckF.Items.Add(allF.Items[i]);
                            ClassCSharp.Flag_add_file = false;
                        }
                        else if (extension == ".docx")
                        {
                            wordapp = new Word.Application(); //--- Создали объект Ворд (открыли(запустили) Ворд)
                            Add_txt_file = ClassCSharp.Check_CS_InDocXFile(fi.FullName);
                            if (Add_txt_file) CheckF.Items.Add(allF.Items[i]);
                            ClassCSharp.Flag_add_file = false;
                        }
                        else
                        {
                            f_txt = new FileStream(allF.GetItemText(allF.Items[i]), FileMode.Open);
                            Add_txt_file = ClassCSharp.Check_CS_InTxtFile(f_txt);
                            if (Add_txt_file) CheckF.Items.Add(allF.Items[i]);
                            ClassCSharp.Flag_add_file = false;
                            f_txt.Close();
                        }
                    }
                    catch (IOException io_ex)
                    {
                        MessageBox.Show("Файл не удалось открыть.\n" + io_ex.ToString());
                        return;
                    }
                }
            }
        }

        //
        //*** Рекурсивная функция вычисления размера полученной директории
        //
        double sizeOfDirectories(string pth, ref double dirSize)
        {
            try
            {
                DirectoryInfo di = new DirectoryInfo(pth);
                DirectoryInfo[] diA = di.GetDirectories();
                FileInfo[] fiA = di.GetFiles();

                foreach (FileInfo f in fiA) dirSize += f.Length;
                foreach (DirectoryInfo df in diA) sizeOfDirectories(df.FullName, ref dirSize);

                return Math.Round((double)(dirSize / 1024 / 1024), 1);
            }
            catch (DirectoryNotFoundException ex)
            {
                MessageBox.Show("Директория не найдена");
                return 0;
            }
            catch (UnauthorizedAccessException ex)
            {
                MessageBox.Show("К директории нет доступа");
                return 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка");
                return 0;
            }
        }

        //
        //*** Открытие файла через ListBox Process
        //*
        private void lbx_list_all_files_MouseDoubleClick(object sender, EventArgs e)
        {
            //*
            if (tbx_txt_exe.TextLength == 0 && tbx_bin_exe.TextLength == 0 && tbx_image_exe.TextLength == 0 && tbx_source_code_exe.TextLength == 0)  //************************Закончил здесь (04.09.2020) - ДОРАБОТАТЬЬЬЬЬЬ***************************
            {
                MessageBox.Show("Файла с сохранёнными путями не существует.\nУстановите пути до сторонних приложений во вкладке 'Настройки'.");
                return;
            }/**/

            string fname = lbx_list_all_files.GetItemText(lbx_list_all_files.SelectedItem); //--- Получаем выбранное в данный момент имя файла в списке

            FileInfo fi = new FileInfo(fname);
            string ext_f = fi.Extension;
            string pth_tmp = "";

            if (TextExe.Contains(ext_f))            pth_tmp = path_txt_exe;
            else if (SourceCodeExe.Contains(ext_f)) pth_tmp = path_source_code_exe;
            else if (ImageExe.Contains(ext_f))      pth_tmp = path_image_exe;
            else if (BinExe.Contains(ext_f))        pth_tmp = path_bin_exe;
            else                                    pth_tmp = path_txt_exe;

            if (ext_f == ".doc" || ext_f == ".docx")
            {
                Form1.wordapp = new Word.Application(); //--- Создали объект Ворд (открыли(запустили) Ворд)
                Form1.wordapp.Visible = true; //--- Сделали его видимым

                Object filename = fi.FullName;
                worddoc = wordapp.Documents.Open(ref filename); //--- Открыли документ
            }
            else
            {
                if (ext_f == ".jpg" || ext_f == ".jpeg") //--- Если открываем картинку
                {
                    if (Flag_image_hex)
                    {
                        System.Drawing.Image img = System.Drawing.Image.FromFile(fname);
                        byte[] arr;
                        using (MemoryStream ms = new MemoryStream())
                        {
                            img.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                            arr = ms.ToArray();
                        }

                        string fnameWithoutExtension = Path.GetFileNameWithoutExtension(fi.Name);
                        string str = BitConverter.ToString(arr); //--- Строка значений всейф картинки
                        string[] Separator = new string[] { "-0A-", "-0B-" }; //--- Разделитель в строке ("перевод на новую строку" - LF = 0x0A)
                        string[] str_without_LF = str.Split(Separator, StringSplitOptions.RemoveEmptyEntries);

                        string pathDir = Directory.GetCurrentDirectory() + "\\Folder_with_image_in_HEX"; //---- Получаем путь до папки, из которой запускаемся
                        if (!Directory.Exists(pathDir))//----- Если искомой папки нет - создаем
                        {
                            DirectoryInfo new_di = Directory.CreateDirectory(pathDir);
                        }

                        FileStream fs = new FileStream(pathDir + "\\" + fnameWithoutExtension + ".hex", FileMode.Create);
                        StreamWriter sw = new StreamWriter(fs);

                        for (int i = 0; i < str_without_LF.Length; i++) sw.WriteLine(str_without_LF[i]);

                        proc.StartInfo.FileName = path_txt_exe;
                        proc.StartInfo.Arguments = pathDir + "\\" + fnameWithoutExtension + ".hex";
                        proc.Start();

                    }
                    else if (Flag_image_original)
                    {
                        proc.StartInfo.FileName = pth_tmp;
                        proc.StartInfo.Arguments = fname;
                        proc.Start();
                    }
                }
                else
                {
                    proc.StartInfo.FileName = pth_tmp;
                    proc.StartInfo.Arguments = fname;
                    proc.Start();
                }
            }/**/
        }/**/

        //
        //*** Открытие файла через ListBox
        //
        private void lbx_list_check_files_MouseDoubleClick(object sender, EventArgs e)
        {
            //*
            if (tbx_txt_exe.TextLength == 0 && tbx_bin_exe.TextLength == 0 && tbx_image_exe.TextLength == 0 && tbx_source_code_exe.TextLength == 0)  //************************Закончил здесь (04.09.2020) - ДОРАБОТАТЬЬЬЬЬЬ***************************
            {
                MessageBox.Show("Файла с сохранёнными путями не существует.\nУстановите пути до сторонних приложений во вкладке 'Настройки'.");
                return;
            }/**/

            string fname = lbx_list_check_files.GetItemText(lbx_list_check_files.SelectedItem); //--- Получаем выбранное в данный момент имя файла в списке

            FileInfo fi = new FileInfo(fname);
            string ext_f = fi.Extension;
            string pth_tmp = "";

            if (TextExe.Contains(ext_f))            pth_tmp = path_txt_exe;
            else if (SourceCodeExe.Contains(ext_f)) pth_tmp = path_source_code_exe;
            else if (ImageExe.Contains(ext_f))      pth_tmp = path_image_exe;
            else if (BinExe.Contains(ext_f))        pth_tmp = path_bin_exe;
            else                                    pth_tmp = path_txt_exe;

            if (ext_f == ".doc" || ext_f == ".docx")
            {
                Form1.wordapp = new Word.Application(); //--- Создали объект Ворд (открыли(запустили) Ворд)
                Form1.wordapp.Visible = true; //--- Сделали его видимым
                Object filename = fi.FullName;

                worddoc = wordapp.Documents.Open(ref filename); //--- Открыли документ
            }
            else
            {
                proc.StartInfo.FileName = pth_tmp;
                proc.StartInfo.Arguments = fname;
                proc.Start();
            }
        }

        //
        //***
        //
        private void lbx_list_YP_MouseDoubleClick(object sender, EventArgs e)
        {

            if (tbx_txt_exe.TextLength == 0 && tbx_bin_exe.TextLength == 0 && tbx_image_exe.TextLength == 0 && tbx_source_code_exe.TextLength == 0)  //************************Закончил здесь (04.09.2020) - ДОРАБОТАТЬЬЬЬЬЬ***************************
            {
                MessageBox.Show("Файла с сохранёнными путями не существует.\nУстановите пути до сторонних приложений во вкладке 'Настройки'.");
                return;
            }

            string fname = lbx_list_YP.GetItemText(lbx_list_YP.SelectedItem); //--- Получаем выбранное в данный момент имя файла в списке

            FileInfo fi = new FileInfo(fname);
            string ext_f = fi.Extension;
            string pth_tmp = "";

            if (TextExe.Contains(ext_f))            pth_tmp = path_txt_exe;
            else if (SourceCodeExe.Contains(ext_f)) pth_tmp = path_source_code_exe;
            else if (ImageExe.Contains(ext_f))      pth_tmp = path_image_exe;
            else if (BinExe.Contains(ext_f))        pth_tmp = path_txt_exe;
            else                                    pth_tmp = path_txt_exe;

            if (ext_f == ".doc" || ext_f == ".docx")
            {
                Form1.wordapp = new Word.Application(); //--- Создали объект Ворд (открыли(запустили) Ворд)
                Form1.wordapp.Visible = true; //--- Сделали его видимым
                Object filename = fi.FullName;

                worddoc = wordapp.Documents.Open(ref filename); //--- Открыли документ
            }
            else
            {
                proc.StartInfo.FileName = pth_tmp;
                proc.StartInfo.Arguments = fname;
                proc.Start();
            }
        }

        //
        //***
        //
        private void lbx_list_bin_files_MouseDoubleClick(object sender, EventArgs e)
        {
            
            if (tbx_txt_exe.TextLength == 0 && tbx_bin_exe.TextLength == 0 && tbx_image_exe.TextLength == 0 && tbx_source_code_exe.TextLength == 0)  //************************Закончил здесь (04.09.2020) - ДОРАБОТАТЬЬЬЬЬЬ***************************
            {
                MessageBox.Show("Файла с сохранёнными путями не существует.\nУстановите пути до сторонних приложений во вкладке 'Настройки'.");
                return;
            }

            string fname = lbx_list_bin_files.GetItemText(lbx_list_bin_files.SelectedItem); //--- Получаем выбранное в данный момент имя файла в списке

            FileInfo fi = new FileInfo(fname);
            string ext_f = fi.Extension;
            string pth_tmp = "";

            if (TextExe.Contains(ext_f))            pth_tmp = path_txt_exe;
            else if (SourceCodeExe.Contains(ext_f)) pth_tmp = path_source_code_exe;
            else if (ImageExe.Contains(ext_f))      pth_tmp = path_image_exe;
            else if (BinExe.Contains(ext_f))        pth_tmp = path_txt_exe;
            else                                    pth_tmp = path_txt_exe;

            if (ext_f == ".doc" || ext_f == ".docx")
            {
                Form1.wordapp = new Word.Application(); //--- Создали объект Ворд (открыли(запустили) Ворд)
                Form1.wordapp.Visible = true; //--- Сделали его видимым
                Object filename = fi.FullName;

                worddoc = wordapp.Documents.Open(ref filename); //--- Открыли документ
            }
            else
            {
                proc.StartInfo.FileName = pth_tmp;
                proc.StartInfo.Arguments = fname;
                proc.Start();
            }
        }

        //
        //*** Открываем файл с описанием приложения
        //
        private void btn_open_description_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog { Multiselect = true, Title = "Выберите один или несколько файлов, содержащих описание программы", InitialDirectory = "C:\\" };

            dlg.ShowDialog();
            if (dlg.FileName == String.Empty) return; //--- Если ничего не выбрано - выходим

            foreach (string file in dlg.FileNames) Set_files_describing_the_program.Add(file); //--- Сохраняем пути до всех выбранных файлов в множество

            if (Set_files_describing_the_program.Count > 1) //--- Если выбраны несколько файлов - заполняем информационные поля
            {
                label8.Text = "Имена файлов:";
                lbl_name_description.Visible = false;
                listBox1.Location = new Point(97, 60);
                listBox1.Size = new System.Drawing.Size(283, 21);
                listBox1.Visible = true;

                label12.Text = "Общий размер файлов (байт):";
                lbl_size_description.Location = new Point(171, 79);

                long total_file_size = 0; //--- Общий размер всех файлов
                foreach (string name in Set_files_describing_the_program) //--- Заполняем список открытых файлов с описанием программы
                {
                    FileInfo fi_d = new FileInfo(name);

                    if (fi_d.Extension == ".doc" || fi_d.Extension == ".docx" || fi_d.Extension == ".txt")
                    {
                        listBox1.Items.Add(fi_d.Name);
                        total_file_size += fi_d.Length; //--- Вычисляю общий размер всех открытых файлов
                    }
                    else
                    {
                        MessageBox.Show("Присутствует неподдерживаемый формат. \nВыбранные файлы должны быть формата: '.doc', '.docx' или '.txt'.");

                        Set_files_describing_the_program.Clear();
                        label8.Text = "Имя файла:";
                        lbl_name_description.Visible = true;
                        listBox1.Location = new Point(364, 13);
                        listBox1.Size = new System.Drawing.Size(16, 21);
                        listBox1.Visible = false;
                        label12.Text = "Размер файла (байт):";
                        lbl_size_description.Location = new Point(124, 79);
                        label17.Visible = true;
                        label13.Visible = true;
                        label16.Visible = true;
                        lbl_date_creation_description.Visible = true;
                        lbl_time_creation_description.Visible = true;
                        lbl_date_final_change_description.Visible = true;

                        return;
                    }  
                }

                lbl_size_description.Text = total_file_size.ToString();

                label17.Text = "Кол-во файлов:";
                label13.Visible = false;
                label16.Visible = false;
                lbl_date_creation_description.Location = new Point(97, 92);
                lbl_time_creation_description.Visible = false;
                lbl_date_final_change_description.Visible = false;
            }
            else if (Set_files_describing_the_program.Count == 1) //--- Если выбран один файл
            {
                label8.Text = "Имя файла:";
                lbl_name_description.Visible = true;
                listBox1.Location = new Point(364, 13);
                listBox1.Size = new System.Drawing.Size(16, 21);
                listBox1.Visible = false;
                label12.Text = "Размер файла (байт):";
                lbl_size_description.Location = new Point(124, 79);
                label17.Text = "Дата создания файла:";
                label13.Visible = true;
                label16.Visible = true;
                lbl_date_creation_description.Location = new Point(134, 92);
                lbl_time_creation_description.Visible = true;
                lbl_date_final_change_description.Visible = true;

                var tmp_list = new List<string>(Set_files_describing_the_program);
                FileInfo fi_d = new FileInfo(tmp_list[0]);

                lbl_name_description.Text = fi_d.Name;

                if (fi_d.Length < 1024)
                {
                    label12.Text = "Размер файла (байт):";
                    lbl_size_description.Text = Convert.ToString(fi_d.Length);
                }
                else if (fi_d.Length >= 1024 && fi_d.Length < 1048576)
                {
                    label12.Text = "Размер файла (Кб):";
                    lbl_size_description.Text = Convert.ToString(Math.Round((double)(fi_d.Length / 1024), 1));
                }
                else
                {
                    label12.Text = "Размер файла (Мб):";
                    lbl_size_description.Text = Convert.ToString(fi_d.Length / 1024 / 1024);
                }

                lbl_date_creation_description.Text = fi_d.CreationTime.ToShortDateString();
                lbl_time_creation_description.Text = fi_d.CreationTime.ToShortTimeString();
                lbl_date_final_change_description.Text = fi_d.LastWriteTime.ToShortDateString();
            }

            btn_open_description.Enabled = false;
            ClassCSharp.Flag_open_descripion = true;
            btn_check_description_and_files.Enabled = true;
        }

        //
        //*** Обработка нажатия на кнопку "Восстановить" (для востановления ранее сохранённого пути к сторонним файлам (если эти пути были уже ранее сохранены))
        //
        private void btn_apply_settings_Click(object sender, EventArgs e)
        {
            //--- Если файл существует
            if (File.Exists(Directory.GetCurrentDirectory() + "\\SavePath.txt"))
            {
                fini = new FileStream(Directory.GetCurrentDirectory() + "\\SavePath.txt", FileMode.Open, FileAccess.Read);
                //FileInfo fi_s = new FileInfo(fini.Name);

                string[] mass_pth_exe = File.ReadAllLines(fini.Name);
                for (int i = 0; i < 4; i++)
                { 
                    switch (i)
                    {
                        case 0:
                            if (mass_pth_exe[i].Length > 4)
                            {
                                path_txt_exe = mass_pth_exe[i];
                                tbx_txt_exe.Text = path_txt_exe;
                            }
                            else
                            {
                                path_txt_exe = "";
                                tbx_txt_exe.Text = "";
                            }
                            break;
                        case 1:
                            if (mass_pth_exe[i].Length > 4)
                            {
                                path_bin_exe = mass_pth_exe[i];
                                tbx_bin_exe.Text = path_bin_exe;
                            }
                            else
                            {
                                path_bin_exe = "";
                                tbx_bin_exe.Text = "";
                            }
                            break;
                        case 2:
                            if (mass_pth_exe[i].Length > 4)
                            {
                                path_image_exe = mass_pth_exe[i];
                                tbx_image_exe.Text = path_image_exe;
                            }
                            else
                            {
                                path_image_exe = "";
                                tbx_image_exe.Text = "";
                            }
                            break;
                        case 3:
                            if (mass_pth_exe[i].Length > 4)
                            {
                                path_source_code_exe = mass_pth_exe[i];
                                tbx_source_code_exe.Text = path_source_code_exe;
                            }
                            else
                            {
                                path_source_code_exe = "";
                                tbx_source_code_exe.Text = "";
                            }
                            break;
                    }
                }

                fini.Close();
            }
            else
            {
                MessageBox.Show("Файла с сохранёнными путями не существует.");
                return;
            }
        }

        //
        //*** Обработка нажатия на кнопку "Сохранить" (для сохранения пути до нужного приложения до изменения местанахождения необходимого приложения)
        //
        private void btn_save_settings_Click(object sender, EventArgs e)
        {
            fini = new FileStream(Directory.GetCurrentDirectory() + "\\SavePath.txt", FileMode.Create, FileAccess.Write); //--- Создаём\перезаписываем файл для сохранения в нём путей до сторонних приложений

            StreamWriter SW = new StreamWriter(fini);
            string s = "";

            //--- Формируем строку для записи в файл
            if (tbx_txt_exe.TextLength > 0)
            {
                s = path_txt_exe;// +".exe";
                SW.WriteLine(s);
            }
            else
            {
                s = "";
                SW.WriteLine(s);
            }

            if (tbx_bin_exe.TextLength > 0)
            {
                s = path_bin_exe;// + ".exe";
                SW.WriteLine(s);
            }
            else
            {
                s = "";
                SW.WriteLine(s);
            }

            if (tbx_image_exe.TextLength > 0)
            {
                s = path_image_exe;// + ".exe";
                SW.WriteLine(s);
            }
            else
            {
                s = "";
                SW.WriteLine(s);
            }

            if (tbx_source_code_exe.TextLength > 0)
            {
                s = path_source_code_exe;// + ".exe";
                SW.WriteLine(s);
            }
            else
            {
                s = "";
                SW.WriteLine(s);
            }

            SW.Close();
            fini.Close();
        }

        //
        //*** Обработка нажатия на кнопку "Сбросить" (для сброса ранее сохранённых путей до сторонних приложений)
        //
        private void btn_reset_settings_Click(object sender, EventArgs e)
        {
            path_txt_exe = tbx_txt_exe.Text = "";
            path_bin_exe = tbx_bin_exe.Text = "";
            path_image_exe = tbx_image_exe.Text = "";
            path_source_code_exe = tbx_source_code_exe.Text = "";
        }

        //
        //*** Получаем путь до стороннего приложения по открытию текстовых файлов
        //
        private void tbx_txt_exe_TextChanged(object sender, EventArgs e)
        {
            if (tbx_txt_exe.TextLength == 0)
            {
                Flag_zero_tbx_txt_exe = true;
                return;
            }
            else 
            {
                path_txt_exe = tbx_txt_exe.Text;
                Flag_zero_tbx_txt_exe = false;
            }
        }

        //
        //*** Получаем путь до стороннего приложения по открытию бинарных файлов
        //
        private void tbx_bin_exe_TextChanged(object sender, EventArgs e)
        {
            if (tbx_bin_exe.TextLength == 0)
            {
                Flag_zero_tbx_bin_exe = true;
                return;
            }
            else
            {
                path_bin_exe = tbx_bin_exe.Text;
                Flag_zero_tbx_bin_exe = false;
            }
        }

        //
        //*** Получаем путь до стороннего приложения по открытию графических файлов
        //
        private void tbx_image_exe_TextChanged(object sender, EventArgs e)
        {
            if (tbx_image_exe.TextLength == 0)
            {
                Flag_zero_tbx_image_exe = true;
                return;
            }
            else
            {
                path_image_exe = tbx_image_exe.Text;
                Flag_zero_tbx_image_exe = false;
            }
        }

        //
        //*** Получаем путь до стороннего приложения по открытию файлов с исходными кодами
        //
        private void tbx_source_code_exe_TextChanged(object sender, EventArgs e)
        {
            path_source_code_exe = tbx_source_code_exe.Text;
            //Flag_zero_tbx_txt_exe = false;
        }

        //
        //*** Закрываем Word
        //
        private void btn_test_word_close_Click(object sender, EventArgs e)
        {
            Object saveChanges = Word.WdSaveOptions.wdPromptToSaveChanges; //--- Определяем, как сохраняет Ворд изменённые документы перед осуществлением выхода. Может быть: -wdDoNotSaveChanges (не сохранять); -wdPromptToSaveChanges (выдать запрос перед сохранением); -wdSaveChanges (сохранить без предупреждения)
            Object originalFormat = Word.WdOriginalFormat.wdWordDocument; //--- Необязат. параметр, определяет формат сохранения для документа. Может быть: -wdOriginalDocumentFormat (в ориг. формате дока (не изменяя его)); -wdPromptUser (по выюору пользователя); -wdWordDocument (формат .doc)
            Object routeDocument = Type.Missing; //--- Необязат. параметр, при true док. направляется следующему получателю, если док. явл. attached документом

            wordapp.Quit(ref saveChanges, ref originalFormat, ref routeDocument);
            wordapp = null;
        }
        
        //
        //*** Кнопка для сравнения реального списка файлов в проекте со списком из "Описания" программы (из документа)
        //
        private void btn_check_description_and_files_Click(object sender, EventArgs e)
        {
            if (btn_check_description_and_files.Text == "Сверить описание и список файлов c исполняемым кодом")
            {
                ClassCSharp.CheckDescriptionAndFiles(Set_files_describing_the_program, lbx_list_all_files);

                if (Flag_converge)
                {
                    btn_check_description_and_files.BackColor = Color.Lime;
                    btn_check_description_and_files.Text = "Файлы в описании совпадают с файлами в проекте";
                }
                else if (Flag_not_converge)
                {
                    groupBox7.Visible = true;
                    lbl_rez.Visible = false;
                    if (!btn_generate_a_report.Visible) btn_generate_a_report.Visible = true;

                    //*--- EXE
                    for (int i = 0; i < ClassCSharp.mass_indescribable_exe_file.Length; i++) lbx_indescribable_executable_file.Items.Add(ClassCSharp.mass_indescribable_exe_file[i]);
                    lbl_count_indescribable_executable_file.Text = Convert.ToString(lbx_indescribable_executable_file.Items.Count);

                    //*--- BIN
                    for (int i = 0; i < ClassCSharp.mass_indescribable_bin_file.Length; i++) lbx_indescribable_binary_file.Items.Add(ClassCSharp.mass_indescribable_bin_file[i]);
                    lbl_count_indescribable_binary_file.Text = Convert.ToString(lbx_indescribable_binary_file.Items.Count);

                    lbl_count_indescribable_file.Text = Convert.ToString(lbx_indescribable_executable_file.Items.Count + lbx_indescribable_binary_file.Items.Count);

                    btn_check_description_and_files.BackColor = Color.Red;
                    btn_check_description_and_files.Text = "Файлы в описании не совпадают с файлами в проекте";
                    MessageBox.Show("Для просмотра неописанных файлов перейти на вкладку 'Результаты проверки'.");

                    ClassCSharp.Flag_enter_descripion = true;
                }
            }
            /*else
            {
                btn_check_description_and_files.BackColor = Color.Orange;
                btn_check_description_and_files.Text = "Сверить описание и список файлов c исходным кодом";
                lbl_name_description.Text = "*";
                lbl_size_description.Text = "*";
                lbl_date_creation_description.Text = "*";
                lbl_time_creation_description.Text = "*";
                lbl_date_final_change_description.Text = "*";
                lbl_count_indescribable_binary_file.Text = "*";
                lbl_count_indescribable_executable_file.Text = "*";
                btn_check_description_and_files.Enabled = false;
                lbx_indescribable_executable_file.Items.Clear();
                lbx_indescribable_binary_file.Items.Clear();
                //tbx_FIO_for_indescribable.Clear();
                groupBox7.Visible = false;
                ClassCSharp.Flag_enter_descripion = false;

                if (groupBox2.Visible == false) lbl_rez.Visible = true;
            }/**/
        }

        //
        //*** Обработка нажатия на клавишу "Сохранить список неописанных файлов"
        //
        private void btn_save_indescribable_file_Click(object sender, EventArgs e)
        {
            if (!Flag_selected_FIO)
            {
                MessageBox.Show("Для формирования отчёта - введите ФИО сотрудника, проводившего анализ.\nПоле для ввода находится на вкладке \"Настройки\".");
                return;
            }

            string pathDir = Directory.GetCurrentDirectory() + "\\Отчёты о проверке директории " + name_of_the_audited_directory + "."; //---- Получаем путь до папки, из которой запускаемся
            if (!Directory.Exists(pathDir))//----- Если искомой папки нет - создаем
            {
                DirectoryInfo new_di = Directory.CreateDirectory(pathDir);
            }

            DateTime dt = DateTime.Now;
            string date = dt.Day.ToString("d2") + "." + dt.Month.ToString("d2") + "." + (dt.Year).ToString("d2");
            string time = dt.Hour.ToString("d2") + ":" + dt.Minute.ToString("d2");
            string pthDocument = pathDir + "\\Отчёт о наличии несоответствия №" + count_reports_indescribable + ".docx"; //--- Полный путь до файла
            count_reports_indescribable++;

            //--- Test Doc
            //*
            object fname = pthDocument;

            Word.Application wordapp = new Word.Application();
            wordapp.Visible = false;

            object missing = Type.Missing;
            Word.Document worddoc = wordapp.Documents.Add(ref missing, ref missing, ref missing, ref missing); //--- Создаём новый документ

            Word.Paragraph para = worddoc.Paragraphs.Add(ref missing); //--- Добавляем новый параграф
            para.Range.Text = "Отчёт от " + date + "г.";
            object style_name = "Заголовок 1";
            para.Range.set_Style(ref style_name); //--- Задаём стиль
            para.Range.Font.Name = "Times New Roman";
            para.Range.Font.Position = 20; //--- Задаём расстояние между заголовком и след. строкой
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; //--- Центруем
            para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
            para.Range.Font.Size = 20; //--- Выставили размер шрифта
            para.Range.Font.Bold = 1; //--- Показали, что строка будет жирной (очень жирной)
            para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

            para.Range.Font.Name = "Times New Roman";
            para.Range.Text = "«О наличии несоответствия»";
            para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
            para.Range.Font.Size = 14; //--- Выставили размер шрифта
            para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
            para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

            para.Range.Font.Name = "Times New Roman";
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            para.Range.Text = "\n\tСегодня, " + date + ", был произведён анализ директории (" + name_of_the_audited_directory + ") и документа (" + lbl_name_description.Text + ") на наличие несоответствия между файлами, находящимися в директории, и файлами, указанными в «Описании программы».\n\tВ ходе проверки были выявлены следующие несоответствия (приведён список файлов, отсутствующих в «Описании программы»).";
            para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
            para.Range.Font.Size = 14; //--- Выставили размер шрифта
            para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
            para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

            para.Range.Font.Name = "Times New Roman";
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            para.Range.Text = "\nИсполняемые файлы:";
            para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
            para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
            para.Range.Font.Size = 14; //--- Выставили размер шрифта
            para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
            para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

            //--- Тестово пробуем набирать пути до файлов
            int index = path_dir.LastIndexOf('\\'); //--- Нашли нужный нам разделитель

            string pth_indescribable_exe_one = lbx_indescribable_executable_file.Items[0].ToString();
            string mass_path_exe = "";
            int count_ind_exe = 0;
            for (int i = 0; i < lbx_list_all_files.Items.Count; i++)
            {
                string tmp = lbx_list_all_files.Items[i].ToString();
                FileInfo fit = new FileInfo(tmp);

                if (pth_indescribable_exe_one == fit.Name)
                {
                    string nameDir = fit.DirectoryName;
                    nameDir = nameDir.Substring(index);
                    mass_path_exe += nameDir + '%';
                    count_ind_exe++;
                    if (count_ind_exe == lbx_indescribable_executable_file.Items.Count) break;
                    pth_indescribable_exe_one = lbx_indescribable_executable_file.Items[count_ind_exe].ToString();
                }
            }
            pth_indescribable_exe = mass_path_exe.Split('%');


            string s_exe = "";
            for (int i = 0; i < lbx_indescribable_executable_file.Items.Count; i++) s_exe += Convert.ToString(i + 1) + ") " + lbx_indescribable_executable_file.Items[i] + ". \n    Находится: " + pth_indescribable_exe[i] + ";\n\n";
            
            para.Range.Font.Name = "Times New Roman";
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            para.Range.Text = s_exe;
            para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
            para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
            para.Range.Font.Size = 14; //--- Выставили размер шрифта
            para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
            para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

            para.Range.Font.Name = "Times New Roman";
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            para.Range.Text = "\n\nБинарные файлы:";
            para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
            para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
            para.Range.Font.Size = 14; //--- Выставили размер шрифта
            para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
            para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

            string pth_indescribable_bin_one = lbx_indescribable_binary_file.Items[0].ToString();
            string mass_path_bin = "";
            int count_ind_bin = 0;
            for (int i = 0; i < lbx_list_all_files.Items.Count; i++)
            {
                string tmp = lbx_list_all_files.Items[i].ToString();
                FileInfo fit = new FileInfo(tmp);

                if (pth_indescribable_bin_one == fit.Name)
                {
                    string nameDir = fit.DirectoryName;
                    nameDir = nameDir.Substring(index);
                    mass_path_bin += nameDir + '%';
                    count_ind_bin++;
                    if (count_ind_bin == lbx_indescribable_binary_file.Items.Count) break;
                    pth_indescribable_bin_one = lbx_indescribable_binary_file.Items[count_ind_bin].ToString();
                }
            }
            pth_indescribable_bin = mass_path_bin.Split('%');


            string s_bin = "";
            for (int i = 0; i < lbx_indescribable_binary_file.Items.Count; i++) s_bin += Convert.ToString(i + 1) + ") " + lbx_indescribable_binary_file.Items[i] + ". \n    Находится: " + pth_indescribable_bin[i] + ";\n\n";
            
            para.Range.Font.Name = "Times New Roman";
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            para.Range.Text = s_bin;
            para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
            para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
            para.Range.Font.Size = 14; //--- Выставили размер шрифта
            para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
            para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

            para.Range.Font.Name = "Times New Roman";
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            para.Range.Text = "\n\n\n\nАнализ проводил " + FIO + ".\nВ " + time + " " + date + "г.";
            para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
            para.Range.Font.Size = 14; //--- Выставили размер шрифта
            para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
            para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ


            worddoc.SaveAs(ref fname, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            
            object save_changes = false;
            worddoc.Close(ref save_changes, ref missing, ref missing);
            wordapp.Quit(ref save_changes, ref missing, ref missing);
        }

        //
        //*** Очистка всех полей (кроме вкладки "Настройки")
        //
        private void btn_Reset_Click(object sender, EventArgs e)
        {
            Flag_enter_Reset = true;

            lbx_list_all_files.Items.Clear();
            lbx_list_check_files.Items.Clear();
            lbx_list_bin_files.Items.Clear();
            lbx_check_connecting_header.Items.Clear();
            lbx_non_ligitimate_Dic.Items.Clear();
            lbx_non_ligitimate_Set.Items.Clear();
            lbx_False_header.Items.Clear();
            lbx_list_YP.Items.Clear();

            Set_files_describing_the_program.Clear();
            Dic_test.Clear();

            groupBox11.Visible = false;

            btn_check_connecting_header.BackColor = Color.Orange;
            btn_check_connecting_header.Text = "Проверить использование файлов с исполняемым кодом и заголовочных файлов в проекте";
            btn_check_connecting_header.Enabled = false;
            lbl_count_check_file.Text = "*";
            lbl_count_YP.Text = "*";
            lbl_bin_files.Text = "*";
            lbl_connect_header.Text = "*";
            lbl_d_size.Text = "*";
            lbl_count_file.Text = "*";
            lbl_date_in.Text = "*";
            lbl_time_in.Text = "*";
            lbl_name_dir.Text = "*";
            Flag_selected_dir = false;
            btn_start_inspect.Enabled = false;
            groupBox2.Enabled = false;
            groupBox4.Enabled = false;
            btn_Reset.Enabled = false;
            lbx_list_all_files.Enabled = false;
            label6.Enabled = false;
            btn_generate_a_report.Visible = false;
            rbn_non_ligitimate_and_indescribable.Checked = true;
            rbn_indescribable.Checked = false;
            rbn_non_ligitimate.Checked = false;

            chbx_Text_file_in_separate_folder.Checked = false;
            chbx_Text_file_in_separate_folder.Enabled = true;

            lbx_False_header.Visible = false;
            lbl_False_header.Visible = false;
            label28.Visible = false;
            label30.Visible = false;
            label30.Text = "";
            ClassCSharp.count_False_header = 0;
            ClassCSharp.Flag_visibl_lbx_false_header = false;
            ClassCSharp.Flag_False_header = false;

            count_YP = 0;
            count_check_files = 0;
            count_check_bin_files = 0;
            count_check_connecting_header = 0;

            lbl_count_non_ligitimate_Dic.Text = "*";
            lbl_count_non_ligitimate_Set.Text = "*";
            lbl_count_non_ligitimate_Sum.Text = "*";
            btn_check_connecting_header.Enabled = false;
            string[] mass_indescribable_header_file = new string[0];
            ClassCSharp.count_zero_in_dic_header_for_exe = 0;
            ClassCSharp.Flag_enter_non_ligitimacy = false;

            ClassCSharp.Set_connected_but_missing_header_files.Clear();

            btn_check_description_and_files.BackColor = Color.Orange;
            btn_check_description_and_files.Text = "Сверить описание и список файлов c исполняемым кодом";
            lbl_name_description.Text = "*";
            lbl_size_description.Text = "*";
            lbl_date_creation_description.Text = "*";
            lbl_time_creation_description.Text = "*";
            lbl_date_final_change_description.Text = "*";
            lbl_count_indescribable_binary_file.Text = "*";
            lbl_count_indescribable_executable_file.Text = "*";
            btn_check_description_and_files.Enabled = false;
            lbx_indescribable_executable_file.Items.Clear();
            lbx_indescribable_binary_file.Items.Clear();

            listBox1.Items.Clear();
            label8.Text = "Имя файла:";
            lbl_name_description.Visible = true;
            listBox1.Location = new Point(364, 13);
            listBox1.Size = new System.Drawing.Size(16, 21);
            listBox1.Visible = false;
            label12.Text = "Размер файла (байт):";
            lbl_size_description.Location = new Point(124, 79);
            label17.Text = "Дата создания файла:";
            label13.Visible = true;
            label16.Visible = true;
            lbl_date_creation_description.Location = new Point(134, 92);
            lbl_time_creation_description.Visible = true;
            lbl_date_final_change_description.Visible = true;


            groupBox7.Visible = false;
            string[] mass_indescribable_exe_file = new string[0];
            string[] mass_indescribable_bin_file = new string[0];
            ClassCSharp.count_exe = 0;
            ClassCSharp.count_bin = 0;

            ClassCSharp.Flag_enter_descripion = false;

            if (groupBox11.Visible == false && groupBox7.Visible == false) lbl_rez.Visible = true;
            btn_open_file.Enabled = true;
        }

        private void rbn_image_hex_CheckedChanged(object sender, EventArgs e)
        {
            Flag_image_hex = true;
            Flag_image_original = false;

            if (rbn_image_hex.Checked) MessageBox.Show("В данном случае, картинка будет сохранена в папке 'Folder_with_image_in_HEX' в формате '.hex'. \n\nРазделителем бежду байтами служит знак 'тире (-)'."); //--- Вывод информационного сообщения
        }

        private void rbn_image_original_CheckedChanged(object sender, EventArgs e)
        {
            Flag_image_original = true;
            Flag_image_hex = false;
        }

        private void btn_dash_Not_remove_CheckedChanged(object sender, EventArgs e)
        {
            Flag_Not_remove_dash = true;
            Flag_remove_dash = false;
        }

        private void btn_dash_remove_CheckedChanged(object sender, EventArgs e)
        {
            Flag_remove_dash = true;
            Flag_Not_remove_dash = false;

            MessageBox.Show("Разделитель будет заменён знаком пробела."); //--- Вывод информационного сообщения
        }

        private void tbx_FIO_TextChanged(object sender, EventArgs e)
        {
            FIO = tbx_FIO.Text;
            if (FIO.Length > 0) Flag_selected_FIO = true;
        }

        private void lbx_check_connecting_header_MouseDoubleClick(object sender, EventArgs e)
        {
            if (tbx_txt_exe.TextLength == 0 && tbx_bin_exe.TextLength == 0 && tbx_image_exe.TextLength == 0 && tbx_source_code_exe.TextLength == 0)  //************************Закончил здесь (04.09.2020) - ДОРАБОТАТЬЬЬЬЬЬ***************************
            {
                MessageBox.Show("Файла с сохранёнными путями не существует.\nУстановите пути до сторонних приложений во вкладке 'Настройки'.");
                return;
            }

            string fname = lbx_check_connecting_header.GetItemText(lbx_check_connecting_header.SelectedItem); //--- Получаем выбранное в данный момент имя файла в списке

            FileInfo fi = new FileInfo(fname);
            string ext_f = fi.Extension;
            string pth_tmp = "";

            if (TextExe.Contains(ext_f)) pth_tmp = path_txt_exe;
            else if (SourceCodeExe.Contains(ext_f)) pth_tmp = path_source_code_exe;
            else if (ImageExe.Contains(ext_f)) pth_tmp = path_image_exe;
            else if (BinExe.Contains(ext_f)) pth_tmp = path_txt_exe;
            else pth_tmp = path_txt_exe;

            if (ext_f == ".doc" || ext_f == ".docx")
            {
                Form1.wordapp = new Word.Application(); //--- Создали объект Ворд (открыли(запустили) Ворд)
                Form1.wordapp.Visible = true; //--- Сделали его видимым
                Object filename = fi.FullName;

                worddoc = wordapp.Documents.Open(ref filename); //--- Открыли документ
            }
            else
            {
                proc.StartInfo.FileName = pth_tmp;
                proc.StartInfo.Arguments = fname;
                proc.Start();
            }
        }

        private void lbx_False_header_MouseDoubleClick(object sender, EventArgs e)
        {
            if (tbx_txt_exe.TextLength == 0 && tbx_bin_exe.TextLength == 0 && tbx_image_exe.TextLength == 0 && tbx_source_code_exe.TextLength == 0)  //************************Закончил здесь (04.09.2020) - ДОРАБОТАТЬЬЬЬЬЬ***************************
            {
                MessageBox.Show("Файла с сохранёнными путями не существует.\nУстановите пути до сторонних приложений во вкладке 'Настройки'.");
                return;
            }

            string fname = lbx_False_header.GetItemText(lbx_False_header.SelectedItem); //--- Получаем выбранное в данный момент имя файла в списке

            FileInfo fi = new FileInfo(fname);
            string ext_f = fi.Extension;
            string pth_tmp = "";

            if (TextExe.Contains(ext_f)) pth_tmp = path_txt_exe;
            else if (SourceCodeExe.Contains(ext_f)) pth_tmp = path_source_code_exe;
            else if (ImageExe.Contains(ext_f)) pth_tmp = path_image_exe;
            else if (BinExe.Contains(ext_f)) pth_tmp = path_txt_exe;
            else pth_tmp = path_txt_exe;

            if (ext_f == ".doc" || ext_f == ".docx")
            {
                Form1.wordapp = new Word.Application(); //--- Создали объект Ворд (открыли(запустили) Ворд)
                Form1.wordapp.Visible = true; //--- Сделали его видимым
                Object filename = fi.FullName;

                worddoc = wordapp.Documents.Open(ref filename); //--- Открыли документ
            }
            else
            {
                proc.StartInfo.FileName = pth_tmp;
                proc.StartInfo.Arguments = fname;
                proc.Start();
            }
        }

        //
        //*** Проверка того: используются ли (подключаются ли) файлы с расширениями ".h" и ".cpp" в проверяемом проекте
        //
        private void btn_check_connecting_header_Click_1(object sender, EventArgs e)
        {
            if (Flag_start_inspect == false)
            {
                MessageBox.Show("Для проведения данной проверки необходимо провести анализ исходной директории, нажав на кнопку 'Анализ файлов'.");
                return;
            }

            if (btn_check_connecting_header.Text == "Проверить использование файлов с исполняемым кодом и заголовочных файлов в проекте")
            {
                ClassCSharp.CheckConnectingHeaderOrBinaryFiles(lbx_check_connecting_header, lbx_list_check_files, lbx_False_header);

                if (ClassCSharp.Set_connected_but_missing_header_files.Count == 0 && ClassCSharp.count_zero_in_dic_header_for_exe == 0)
                {
                    btn_check_connecting_header.BackColor = Color.Lime;
                    btn_check_connecting_header.Text = "Все заголовочные файлы проекта - легитимны";
                }
                else
                {
                    lbl_rez.Visible = false;
                    if (!btn_generate_a_report.Visible) btn_generate_a_report.Visible = true;

                    if (ClassCSharp.Flag_visibl_lbx_false_header)
                    {
                        lbx_False_header.Visible = true;
                        lbl_False_header.Visible = true;
                        label28.Visible = true;
                        label30.Visible = true;
                        label30.Text = ClassCSharp.count_False_header.ToString();
                    }

                    groupBox11.Visible = true;
                    btn_check_connecting_header.BackColor = Color.Red;

                    btn_check_connecting_header.Text = "Присутствуют нелигитимные заголовочные файлы";
                    MessageBox.Show("Для просмотра нелигитимных файлов перейти на вкладку 'Результаты проверки'.");

                    if (ClassCSharp.Set_connected_but_missing_header_files.Count != 0 && ClassCSharp.count_zero_in_dic_header_for_exe == 0) //--- Если есть заголовочные файлы, подключённые к исполняемым, но отсутствующие в проверяемой директории
                    {

                        foreach (string s_set in ClassCSharp.Set_connected_but_missing_header_files) lbx_non_ligitimate_Set.Items.Add(s_set);

                        lbl_count_non_ligitimate_Set.Text = ClassCSharp.Set_connected_but_missing_header_files.Count.ToString();
                    }
                    else if (ClassCSharp.Set_connected_but_missing_header_files.Count == 0 && ClassCSharp.count_zero_in_dic_header_for_exe != 0) //--- Если есть заголовочные файлы, находящиеся в проверяемой директории, но ни разу не подключённые к исполняемым файлам 
                    {
                        for (int i = 0; i < ClassCSharp.count_zero_in_dic_header_for_exe; i++) lbx_non_ligitimate_Dic.Items.Add(ClassCSharp.mass_indescribable_header_file[i]);

                        lbl_count_non_ligitimate_Dic.Text = ClassCSharp.count_zero_in_dic_header_for_exe.ToString();
                    }
                    else if (ClassCSharp.Set_connected_but_missing_header_files.Count != 0 && ClassCSharp.count_zero_in_dic_header_for_exe != 0)
                    {
                        foreach (string s_set in ClassCSharp.Set_connected_but_missing_header_files) lbx_non_ligitimate_Set.Items.Add(s_set);

                        lbl_count_non_ligitimate_Set.Text = ClassCSharp.Set_connected_but_missing_header_files.Count.ToString();

                        for (int i = 0; i < ClassCSharp.count_zero_in_dic_header_for_exe; i++) lbx_non_ligitimate_Dic.Items.Add(ClassCSharp.mass_indescribable_header_file[i]);

                        lbl_count_non_ligitimate_Dic.Text = ClassCSharp.count_zero_in_dic_header_for_exe.ToString();
                    }

                    //***Test
                    lbx_non_ligitimate_Set.DrawMode = DrawMode.OwnerDrawFixed;
                    lbx_non_ligitimate_Set.DrawItem += (sender_1, e_1) => {

                        e_1.DrawBackground();
                        Graphics g = e_1.Graphics;

                        var val = lbx_non_ligitimate_Set.Items[e_1.Index];

                        g.FillRectangle(new SolidBrush(ClassCSharp.mass_header_standart_lib.Contains(val) ? Color.LightGreen : Color.White), e_1.Bounds);
                        g.DrawString(val.ToString(), e_1.Font, new SolidBrush(e_1.ForeColor), e_1.Bounds);

                        e_1.DrawFocusRectangle();
                    };

                    lbl_count_non_ligitimate_Sum.Text = (ClassCSharp.count_zero_in_dic_header_for_exe + lbx_non_ligitimate_Set.Items.Count).ToString();

                    ClassCSharp.Flag_enter_non_ligitimacy = true;
                }
            }
            /*else
            {
                groupBox11.Visible = false;
                btn_check_connecting_header.BackColor = Color.Orange;
                btn_check_connecting_header.Text = "Проверить использование заголовочных и бинарных файлов в проекте";
                lbx_non_ligitimate_Dic.Items.Clear();
                lbx_non_ligitimate_Set.Items.Clear();
                lbl_count_non_ligitimate_Dic.Text = "*";
                lbl_count_non_ligitimate_Set.Text = "*";
                lbl_count_non_ligitimate_Sum.Text = "*";
                btn_check_connecting_header.Enabled = false;
                ClassCSharp.mass_indescribable_header_file = new string[ClassCSharp.count_zero_in_dic_header_for_exe];
                ClassCSharp.count_zero_in_dic_header_for_exe = 0;
                ClassCSharp.Flag_enter_non_ligitimacy = false;

                if (groupBox7.Visible == false) lbl_rez.Visible = true;
            }/**/
        }

        //
        //*** Обработка формирования отчёта по нелегитимным файлам
        //
        private void button1_Click(object sender, EventArgs e)
        {
            if (!Flag_selected_FIO)
            {
                MessageBox.Show("Для формирования отчёта - введите ФИО сотрудника, проводившего анализ.\nПоле для ввода находится на вкладке \"Настройки\".");
                return;
            }

            string pathDir = Directory.GetCurrentDirectory() + "\\Отчёты о проверке директории " + name_of_the_audited_directory + "."; //---- Получаем путь до папки, из которой запускаемся
            if (!Directory.Exists(pathDir))//----- Если искомой папки нет - создаем
            {
                DirectoryInfo new_di = Directory.CreateDirectory(pathDir);
            }

            DateTime dt = DateTime.Now;
            string date = dt.Day.ToString("d2") + "." + dt.Month.ToString("d2") + "." + (dt.Year).ToString("d2");
            string time = dt.Hour.ToString("d2") + ":" + dt.Minute.ToString("d2");
            string pthDocument = pathDir + "\\Отчёт о наличии нелегитимности №" + count_reports_non_legitimace + ".docx"; //--- Полный путь до файла
            count_reports_non_legitimace++;

            //--- Test Doc
            //*
            object fname = pthDocument;

            Word.Application wordapp = new Word.Application();
            wordapp.Visible = false;

            object missing = Type.Missing;
            Word.Document worddoc = wordapp.Documents.Add(ref missing, ref missing, ref missing, ref missing); //--- Создаём новый документ

            Word.Paragraph para = worddoc.Paragraphs.Add(ref missing); //--- Добавляем новый параграф
            para.Range.Text = "Отчёт от " + date + "г.";
            object style_name = "Заголовок 1";
            para.Range.set_Style(ref style_name); //--- Задаём стиль
            para.Range.Font.Name = "Times New Roman";
            para.Range.Font.Position = 20; //--- Задаём расстояние между заголовком и след. строкой
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; //--- Центруем
            para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
            para.Range.Font.Size = 20; //--- Выставили размер шрифта
            para.Range.Font.Bold = 1; //--- Показали, что строка будет жирной (очень жирной)
            para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

            para.Range.Font.Name = "Times New Roman";
            para.Range.Text = "«О наличии не легитимности»";
            para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
            para.Range.Font.Size = 14; //--- Выставили размер шрифта
            para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
            para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

            para.Range.Font.Name = "Times New Roman";
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            para.Range.Text = "\n\tСегодня, " + date + ", был произведён анализ директории (" + name_of_the_audited_directory + ") на наличие нелегитимных файлов, находящихся в директории.\n\tВ ходе проверки были выявлены следующие нелегитимные файлы (приведён список файлов).";
            para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
            para.Range.Font.Size = 14; //--- Выставили размер шрифта
            para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
            para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

            if (lbx_non_ligitimate_Dic.Items.Count != 0)
            {
                para.Range.Font.Name = "Times New Roman";
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                para.Range.Text = "\nФайлы, находящиеся в директории, но не подключаемые к исполняемым:";
                para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                para.Range.Font.Size = 14; //--- Выставили размер шрифта
                para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                int index = path_dir.LastIndexOf('\\'); //--- Нашли нужный нам разделитель

                string pth_non_legitimace_Dic_one = lbx_non_ligitimate_Dic.Items[0].ToString();
                string mass_path_Dic = "";
                int count_ind_Dic = 0;
                for (int i = 0; i < lbx_list_all_files.Items.Count; i++)
                {
                    string tmp = lbx_list_all_files.Items[i].ToString();
                    FileInfo fit = new FileInfo(tmp);

                    if (pth_non_legitimace_Dic_one == fit.Name)
                    {
                        string nameDir = fit.DirectoryName;
                        nameDir = nameDir.Substring(index);
                        mass_path_Dic += nameDir + '%';
                        count_ind_Dic++;
                        if (count_ind_Dic == lbx_non_ligitimate_Dic.Items.Count) break;
                        pth_non_legitimace_Dic_one = lbx_non_ligitimate_Dic.Items[count_ind_Dic].ToString();
                    }
                }
                pth_non_legitimace_Dic = mass_path_Dic.Split('%');


                string s_Dic = "";
                for (int i = 0; i < lbx_non_ligitimate_Dic.Items.Count; i++) s_Dic += Convert.ToString(i + 1) + ") " + lbx_non_ligitimate_Dic.Items[i] + ". \n    Находится: " + pth_non_legitimace_Dic[i] + ";\n\n";

                para.Range.Font.Name = "Times New Roman";
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                para.Range.Text = s_Dic;
                para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                para.Range.Font.Size = 14; //--- Выставили размер шрифта
                para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ
            }

            if (lbx_non_ligitimate_Set.Items.Count != 0)
            {
                para.Range.Font.Name = "Times New Roman";
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                para.Range.Text = "\n\nФайлы, подключаемые к исполняемым, но отсутствующие в директории:";
                para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                para.Range.Font.Size = 14; //--- Выставили размер шрифта
                para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                

                string s_Set = "";
                for (int i = 0; i < lbx_non_ligitimate_Set.Items.Count; i++)
                {
                    string name_head = lbx_non_ligitimate_Set.Items[i].ToString(); //--- Получаем заголовочный файл
                    HashSet<string> name_exe = ClassCSharp.dic_name_header_file_Set[name_head]; //--- Получаем список исполняемых файлов, к которым подключается текущий заголовочный
                    int count_name_exe = name_exe.Count; //--- Получаем кол-во исполняемых файлов
                    string s_name_exe = "";
                    int j = 1;

                    foreach (string s in name_exe)
                    {
                        s_name_exe += "    " + Convert.ToString(i + 1) + "." + Convert.ToString(j) + ") " + s + ";\n";
                        j++;
                    }

                    s_Set += Convert.ToString(i + 1) + ") " + name_head + ".\n    Подключается к файлам:\n" + s_name_exe + "\n";
                }
                para.Range.Font.Name = "Times New Roman";
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                para.Range.Text = s_Set;
                para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                para.Range.Font.Size = 14; //--- Выставили размер шрифта
                para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ
            }

            if (lbx_False_header.Items.Count != 0)
            {
                para.Range.Font.Name = "Times New Roman";
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                para.Range.Text = "\n\nФайлы с исполняемым кодом, замаскеруемые поз заголовочные файлы:";
                para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                para.Range.Font.Size = 14; //--- Выставили размер шрифта
                para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ


                string s = "";
                for (int i = 0; i < lbx_False_header.Items.Count; i++) s += Convert.ToString(i + 1) + ") " + lbx_False_header.Items[i] + ";\n\n";

                para.Range.Font.Name = "Times New Roman";
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                para.Range.Text = s;
                para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                para.Range.Font.Size = 14; //--- Выставили размер шрифта
                para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ
            }

            para.Range.Font.Name = "Times New Roman";
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            para.Range.Text = "\n\n\n\nАнализ проводил " + FIO + ".\nВ " + time + " " + date + "г.";
            para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
            para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
            para.Range.Font.Size = 14; //--- Выставили размер шрифта
            para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
            para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ


            worddoc.SaveAs(ref fname, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

            object save_changes = false;
            worddoc.Close(ref save_changes, ref missing, ref missing);
            wordapp.Quit(ref save_changes, ref missing, ref missing);
        }

        private void lbx_list_all_files_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lbx_list_all_files.SelectedItem != null) MessageBox.Show(lbx_list_all_files.SelectedItem.ToString());
        }

        //
        //*** Если выводим отчёт ТОЛЬКО с нелегитимными файлами
        //
        private void rbn_non_ligitimate_CheckedChanged(object sender, EventArgs e)
        {
            if (rbn_non_ligitimate.Checked) Flag_choice_radiobatton_non_ligitim = true;
            else                            Flag_choice_radiobatton_non_ligitim = false;
        }

        //
        //*** Если выводим отчёт ТОЛЬКО с неописанными файлами
        //
        private void rbn_indescribable_CheckedChanged(object sender, EventArgs e)
        {
            if (rbn_indescribable.Checked) Flag_choice_radiobatton_indescribable = true;
            else                           Flag_choice_radiobatton_indescribable = false;
        }

        //
        //*** Если выводим отчёт с нелегитимными И неописанными файлами
        //
        private void rbn_non_ligitimate_and_indescribable_CheckedChanged(object sender, EventArgs e)
        {
            if (rbn_non_ligitimate_and_indescribable.Checked) Flag_choice_radiobatton_non_ligitim_and_indescribable = true;
            else                                              Flag_choice_radiobatton_non_ligitim_and_indescribable = false;
        }

        //
        //*** Формирование отчёта
        //
        private void btn_generate_a_report_Click(object sender, EventArgs e)
        {
            if (!Flag_selected_FIO) //--- Если невведено ФИО сотрудника
            {
                MessageBox.Show("Для формирования отчёта - введите ФИО сотрудника, проводившего анализ.\nПоле для ввода находится на вкладке \"Настройки\".");
                return;
            }

            string pathDir = Directory.GetCurrentDirectory() + "\\Отчёты о проверке директории " + name_of_the_audited_directory + "."; //---- Получаем путь до папки, из которой запускаемся
            if (!Directory.Exists(pathDir))//----- Если искомой папки нет - создаем
            {
                Directory.CreateDirectory(pathDir);
            }

            DirectoryInfo di_tmp = new DirectoryInfo(pathDir);
            FileInfo[] fi_tmp = di_tmp.GetFiles();

            DateTime dt = DateTime.Now;
            string date = dt.Day.ToString("d2") + "." + dt.Month.ToString("d2") + "." + (dt.Year).ToString("d2");
            string time = dt.Hour.ToString("d2") + ":" + dt.Minute.ToString("d2");
            string pthDocument = pathDir + "\\Отчёт №" + (fi_tmp.Length + 1) + " от " + date + "г.docx"; //--- Полный путь до файла
            count_reports_indescribable++;

            object fname = pthDocument;

            Word.Application wordapp = new Word.Application();
            wordapp.Visible = false;

            object missing = Type.Missing;
            Word.Document worddoc = wordapp.Documents.Add(ref missing, ref missing, ref missing, ref missing); //--- Создаём новый документ

            Word.Paragraph para;
            Word.Paragraph para1;
            Word.Paragraph para2;

            if (Flag_choice_radiobatton_indescribable)
            {
                para = worddoc.Paragraphs.Add(ref missing); //--- Добавляем новый параграф

                para.Range.Text = "Отчёт от " + date + ".";
                object style_name = "Заголовок 1";
                para.Range.set_Style(ref style_name); //--- Задаём стиль
                para.Range.Font.Name = "Times New Roman";
                para.Range.Font.Position = 20; //--- Задаём расстояние между заголовком и след. строкой
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; //--- Центруем
                para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                para.Range.Font.Size = 20; //--- Выставили размер шрифта
                para.Range.Font.Bold = 1; //--- Показали, что строка будет жирной (очень жирной)
                para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                para.Range.Font.Name = "Times New Roman";
                para.Range.Text = "\n«О наличии несоответствия»";
                para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                para.Range.Font.Size = 14; //--- Выставили размер шрифта
                para.Range.Font.Bold = 1; //--- Показали, что строка будет жирной (очень жирной)
                para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                para.Range.Font.Name = "Times New Roman";
                para.Range.Text = "\n";
                para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                para.Range.Font.Size = 14; //--- Выставили размер шрифта
                para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                para.Range.Font.Name = "Times New Roman";
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                para.Range.Text = "\tСегодня, " + date + ", был произведён анализ директории (" + name_of_the_audited_directory + ") и документа (" + lbl_name_description.Text + ") на наличие несоответствия между файлами, находящимися в директории, и файлами, указанными в «Описании программы».\n\tВ ходе проверки были выявлены следующие несоответствия (приведён список файлов, отсутствующих в «Описании программы»).";
                para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                para.Range.Font.Size = 14; //--- Выставили размер шрифта
                para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                int index = path_dir.LastIndexOf('\\'); //--- Нашли нужный нам разделитель

                if (lbx_indescribable_executable_file.Items.Count != 0)
                {
                    para.Range.Font.Name = "Times New Roman";
                    para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    para.Range.Text = "\nИсполняемые файлы:";
                    para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                    para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                    para.Range.Font.Size = 14; //--- Выставили размер шрифта
                    para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                    para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ
                
                    string pth_indescribable_exe_one = lbx_indescribable_executable_file.Items[0].ToString();
                    string mass_path_exe = "";
                    int count_ind_exe = 0;
                    for (int i = 0; i < lbx_list_all_files.Items.Count; i++)
                    {
                        string tmp = lbx_list_all_files.Items[i].ToString();
                        FileInfo fit = new FileInfo(tmp);

                        if (pth_indescribable_exe_one == fit.Name)
                        {
                            string nameDir = fit.DirectoryName;
                            nameDir = nameDir.Substring(index);
                            mass_path_exe += nameDir + '%';
                            count_ind_exe++;
                            if (count_ind_exe == lbx_indescribable_executable_file.Items.Count) break;
                            pth_indescribable_exe_one = lbx_indescribable_executable_file.Items[count_ind_exe].ToString();
                        }
                    }
                    pth_indescribable_exe = mass_path_exe.Split('%');


                    string s_exe = "";
                    for (int i = 0; i < lbx_indescribable_executable_file.Items.Count; i++) s_exe += Convert.ToString(i + 1) + ") " + lbx_indescribable_executable_file.Items[i] + " \n    Находится: " + pth_indescribable_exe[i] + ";\n\n";

                    para.Range.Font.Name = "Times New Roman";
                    para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    para.Range.Text = s_exe;
                    para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                    para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                    para.Range.Font.Size = 14; //--- Выставили размер шрифта
                    para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                    para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ
                }

                if (lbx_indescribable_binary_file.Items.Count != 0)
                {
                    para.Range.Font.Name = "Times New Roman";
                    para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    para.Range.Text = "\n\nБинарные файлы:";
                    para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                    para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                    para.Range.Font.Size = 14; //--- Выставили размер шрифта
                    para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                    para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                    string pth_indescribable_bin_one = lbx_indescribable_binary_file.Items[0].ToString();
                    string mass_path_bin = "";
                    int count_ind_bin = 0;
                    for (int i = 0; i < lbx_list_all_files.Items.Count; i++)
                    {
                        string tmp = lbx_list_all_files.Items[i].ToString();
                        FileInfo fit = new FileInfo(tmp);

                        if (pth_indescribable_bin_one == fit.Name)
                        {
                            string nameDir = fit.DirectoryName;
                            nameDir = nameDir.Substring(index);
                            mass_path_bin += nameDir + '%';
                            count_ind_bin++;
                            if (count_ind_bin == lbx_indescribable_binary_file.Items.Count) break;
                            pth_indescribable_bin_one = lbx_indescribable_binary_file.Items[count_ind_bin].ToString();
                        }
                    }
                    pth_indescribable_bin = mass_path_bin.Split('%');


                    string s_bin = "";
                    for (int i = 0; i < lbx_indescribable_binary_file.Items.Count; i++) s_bin += Convert.ToString(i + 1) + ") " + lbx_indescribable_binary_file.Items[i] + " \n    Находится: " + pth_indescribable_bin[i] + ";\n\n";

                    para.Range.Font.Name = "Times New Roman";
                    para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    para.Range.Text = s_bin;
                    para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                    para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                    para.Range.Font.Size = 14; //--- Выставили размер шрифта
                    para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                    para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ
                }

                para.Range.Font.Name = "Times New Roman";
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                para.Range.Text = "\n\n\n\n";
                para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                para.Range.Font.Size = 14; //--- Выставили размер шрифта
                para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                para.Range.Font.Name = "Times New Roman";
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                para.Range.Text = "Анализ проводил " + FIO + ".\nВ " + time + " " + date + "г.";
                para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                para.Range.Font.Size = 14; //--- Выставили размер шрифта
                para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ


                worddoc.SaveAs(ref fname, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                object save_changes = false;
                worddoc.Close(ref save_changes, ref missing, ref missing);
                wordapp.Quit(ref save_changes, ref missing, ref missing);
            }
            else if (Flag_choice_radiobatton_non_ligitim)
            {
                para = worddoc.Paragraphs.Add(ref missing); //--- Добавляем новый параграф

                para.Range.Text = "Отчёт от " + date + "г.";
                object style_name = "Заголовок 1";
                para.Range.set_Style(ref style_name); //--- Задаём стиль
                para.Range.Font.Name = "Times New Roman";
                para.Range.Font.Position = 20; //--- Задаём расстояние между заголовком и след. строкой
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; //--- Центруем
                para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                para.Range.Font.Size = 20; //--- Выставили размер шрифта
                para.Range.Font.Bold = 1; //--- Показали, что строка будет жирной (очень жирной)
                para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                para.Range.Font.Name = "Times New Roman";
                para.Range.Text = "\n«О наличии не легитимности»";
                para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                para.Range.Font.Size = 14; //--- Выставили размер шрифта
                para.Range.Font.Bold = 1; //--- Показали, что строка будет жирной (очень жирной)
                para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                para.Range.Font.Name = "Times New Roman";
                para.Range.Text = "\n";
                para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                para.Range.Font.Size = 14; //--- Выставили размер шрифта
                para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                para.Range.Font.Name = "Times New Roman";
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                para.Range.Text = "\tСегодня, " + date + ", был произведён анализ директории (" + name_of_the_audited_directory + ") на наличие нелегитимных файлов, находящихся в директории.\n\tВ ходе проверки были выявлены следующие нелегитимные файлы (приведён список файлов).";
                para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                para.Range.Font.Size = 14; //--- Выставили размер шрифта
                para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                int index = path_dir.LastIndexOf('\\'); //--- Нашли нужный нам разделитель

                if (lbx_non_ligitimate_Dic.Items.Count != 0)
                {
                    para.Range.Font.Name = "Times New Roman";
                    para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    para.Range.Text = "\nФайлы, находящиеся в директории, но не подключаемые к исполняемым:";
                    para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                    para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                    para.Range.Font.Size = 14; //--- Выставили размер шрифта
                    para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                    para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                    string pth_non_legitimace_Dic_one = lbx_non_ligitimate_Dic.Items[0].ToString();
                    string mass_path_Dic = "";
                    int count_ind_Dic = 0;
                    for (int i = 0; i < lbx_list_all_files.Items.Count; i++)
                    {
                        string tmp = lbx_list_all_files.Items[i].ToString();
                        FileInfo fit = new FileInfo(tmp);

                        if (pth_non_legitimace_Dic_one == fit.Name)
                        {
                            string nameDir = fit.DirectoryName;
                            nameDir = nameDir.Substring(index);
                            mass_path_Dic += nameDir + '%';
                            count_ind_Dic++;
                            if (count_ind_Dic == lbx_non_ligitimate_Dic.Items.Count) break;
                            pth_non_legitimace_Dic_one = lbx_non_ligitimate_Dic.Items[count_ind_Dic].ToString();
                        }
                    }
                    pth_non_legitimace_Dic = mass_path_Dic.Split('%');


                    string s_Dic = "";
                    for (int i = 0; i < lbx_non_ligitimate_Dic.Items.Count; i++) s_Dic += Convert.ToString(i + 1) + ") " + lbx_non_ligitimate_Dic.Items[i] + " \n    Находится: " + pth_non_legitimace_Dic[i] + ";\n\n";

                    para.Range.Font.Name = "Times New Roman";
                    para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    para.Range.Text = s_Dic;
                    para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                    para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                    para.Range.Font.Size = 14; //--- Выставили размер шрифта
                    para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                    para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ
                }

                if (lbx_non_ligitimate_Set.Items.Count != 0)
                {
                    para.Range.Font.Name = "Times New Roman";
                    para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    para.Range.Text = "\n\nФайлы, подключаемые к исполняемым, но отсутствующие в директории:";
                    para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                    para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                    para.Range.Font.Size = 14; //--- Выставили размер шрифта
                    para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                    para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ



                    string s_Set = "";
                    for (int i = 0; i < lbx_non_ligitimate_Set.Items.Count; i++)
                    {
                        string name_head = lbx_non_ligitimate_Set.Items[i].ToString(); //--- Получаем заголовочный файл
                        HashSet<string> name_exe = ClassCSharp.dic_name_header_file_Set[name_head]; //--- Получаем список исполняемых файлов, к которым подключается текущий заголовочный
                        int count_name_exe = name_exe.Count; //--- Получаем кол-во исполняемых файлов
                        string s_name_exe = "";
                        int j = 1;

                        foreach (string s in name_exe)
                        {
                            s_name_exe += "    " + Convert.ToString(i + 1) + "." + Convert.ToString(j) + ") " + s + ";\n";
                            j++;
                        }

                        s_Set += Convert.ToString(i + 1) + ") " + name_head + "\n    Подключается к файлам:\n" + s_name_exe + "\n";
                    }
                    para.Range.Font.Name = "Times New Roman";
                    para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    para.Range.Text = s_Set;
                    para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                    para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                    para.Range.Font.Size = 14; //--- Выставили размер шрифта
                    para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                    para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ
                }

                if (lbx_False_header.Items.Count != 0)
                {
                    para.Range.Font.Name = "Times New Roman";
                    para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    para.Range.Text = "\n\nФайлы с исполняемым кодом, маскируемые под заголовочные файлы:";
                    para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                    para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                    para.Range.Font.Size = 14; //--- Выставили размер шрифта
                    para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                    para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ


                    string s = "";
                    for (int i = 0; i < lbx_False_header.Items.Count; i++) s += Convert.ToString(i + 1) + ") " + lbx_False_header.Items[i] + ";\n\n";

                    para.Range.Font.Name = "Times New Roman";
                    para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    para.Range.Text = s;
                    para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                    para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                    para.Range.Font.Size = 14; //--- Выставили размер шрифта
                    para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                    para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ
                }

                para.Range.Font.Name = "Times New Roman";
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                para.Range.Text = "\n\n\n\n";
                para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                para.Range.Font.Size = 14; //--- Выставили размер шрифта
                para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                para.Range.Font.Name = "Times New Roman";
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                para.Range.Text = "Анализ проводил " + FIO + ".\nВ " + time + " " + date + "г.";
                para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                para.Range.Font.Size = 14; //--- Выставили размер шрифта
                para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ


                worddoc.SaveAs(ref fname, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                object save_changes = false;
                worddoc.Close(ref save_changes, ref missing, ref missing);
                wordapp.Quit(ref save_changes, ref missing, ref missing);
            }
            else if (Flag_choice_radiobatton_non_ligitim_and_indescribable)
            {
                //*** Первый параграф - с неописанными файлами
                if (groupBox7.Visible)
                {
                    para1 = worddoc.Paragraphs.Add(ref missing); //--- Добавляем новый параграф

                    para1.Range.Text = "Отчёт от " + date + ".";
                    object style_name = "Заголовок 1";
                    para1.Range.set_Style(ref style_name); //--- Задаём стиль
                    para1.Range.Font.Name = "Times New Roman";
                    para1.Range.Font.Position = 20; //--- Задаём расстояние между заголовком и след. строкой
                    para1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; //--- Центруем
                    para1.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                    para1.Range.Font.Size = 20; //--- Выставили размер шрифта
                    para1.Range.Font.Bold = 1; //--- Показали, что строка будет жирной (очень жирной)
                    para1.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                    para1.Range.Font.Name = "Times New Roman";
                    para1.Range.Text = "\n«О наличии несоответствия»";
                    para1.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                    para1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    para1.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                    para1.Range.Font.Size = 14; //--- Выставили размер шрифта
                    para1.Range.Font.Bold = 1; //--- Показали, что строка будет жирной (очень жирной)
                    //para1.Range.Font.Underline = Word.WdUnderline.wdUnderlineDashHeavy; //--- Подчеркнули
                    para1.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                    para1.Range.Font.Name = "Times New Roman";
                    para1.Range.Text = "\n";
                    para1.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                    para1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    para1.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                    para1.Range.Font.Size = 14; //--- Выставили размер шрифта
                    para1.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                    //para1.Range.Font.Underline = Word.WdUnderline.wdUnderlineDashHeavy; //--- Подчеркнули
                    para1.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                    para1.Range.Font.Name = "Times New Roman";
                    para1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    para1.Range.Text = "\tСегодня, " + date + ", был произведён анализ директории (" + name_of_the_audited_directory + ") и документа (" + lbl_name_description.Text + ") на наличие несоответствия между файлами, находящимися в директории, и файлами, указанными в «Описании программы».\n\tВ ходе проверки были выявлены следующие несоответствия (приведён список файлов, отсутствующих в «Описании программы»).";
                    para1.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                    para1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    para1.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                    para1.Range.Font.Size = 14; //--- Выставили размер шрифта
                    para1.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                    para1.Range.Font.Underline = Word.WdUnderline.wdUnderlineNone;
                    para1.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                    int index = path_dir.LastIndexOf('\\'); //--- Нашли нужный нам разделитель

                    if (lbx_indescribable_executable_file.Items.Count != 0)
                    {
                        para1.Range.Font.Name = "Times New Roman";
                        para1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        para1.Range.Text = "\nИсполняемые файлы:";
                        para1.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                        para1.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                        para1.Range.Font.Size = 14; //--- Выставили размер шрифта
                        para1.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                        para1.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                        string pth_indescribable_exe_one = lbx_indescribable_executable_file.Items[0].ToString();
                        string mass_path_exe = "";
                        int count_ind_exe = 0;
                        for (int i = 0; i < lbx_list_all_files.Items.Count; i++)
                        {
                            string tmp = lbx_list_all_files.Items[i].ToString();
                            FileInfo fit = new FileInfo(tmp);

                            if (pth_indescribable_exe_one == fit.Name)
                            {
                                string nameDir = fit.DirectoryName;
                                nameDir = nameDir.Substring(index);
                                mass_path_exe += nameDir + '%';
                                count_ind_exe++;
                                if (count_ind_exe == lbx_indescribable_executable_file.Items.Count) break;
                                pth_indescribable_exe_one = lbx_indescribable_executable_file.Items[count_ind_exe].ToString();
                            }
                        }
                        pth_indescribable_exe = mass_path_exe.Split('%');


                        string s_exe = "";
                        for (int i = 0; i < lbx_indescribable_executable_file.Items.Count; i++) s_exe += Convert.ToString(i + 1) + ") " + lbx_indescribable_executable_file.Items[i] + " \n    Находится: " + pth_indescribable_exe[i] + ";\n\n";

                        para1.Range.Font.Name = "Times New Roman";
                        para1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        para1.Range.Text = s_exe;
                        para1.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                        para1.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                        para1.Range.Font.Size = 14; //--- Выставили размер шрифта
                        para1.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                        para1.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ
                    }

                    if (lbx_indescribable_binary_file.Items.Count != 0)
                    {
                        para1.Range.Font.Name = "Times New Roman";
                        para1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        para1.Range.Text = "\n\nБинарные файлы:";
                        para1.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                        para1.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                        para1.Range.Font.Size = 14; //--- Выставили размер шрифта
                        para1.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                        para1.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                        string pth_indescribable_bin_one = lbx_indescribable_binary_file.Items[0].ToString();
                        string mass_path_bin = "";
                        int count_ind_bin = 0;
                        for (int i = 0; i < lbx_list_all_files.Items.Count; i++)
                        {
                            string tmp = lbx_list_all_files.Items[i].ToString();
                            FileInfo fit = new FileInfo(tmp);

                            if (pth_indescribable_bin_one == fit.Name)
                            {
                                string nameDir = fit.DirectoryName;
                                nameDir = nameDir.Substring(index);
                                mass_path_bin += nameDir + '%';
                                count_ind_bin++;
                                if (count_ind_bin == lbx_indescribable_binary_file.Items.Count) break;
                                pth_indescribable_bin_one = lbx_indescribable_binary_file.Items[count_ind_bin].ToString();
                            }
                        }
                        pth_indescribable_bin = mass_path_bin.Split('%');


                        string s_bin = "";
                        for (int i = 0; i < lbx_indescribable_binary_file.Items.Count; i++) s_bin += Convert.ToString(i + 1) + ") " + lbx_indescribable_binary_file.Items[i] + " \n    Находится: " + pth_indescribable_bin[i] + ";\n\n";

                        para1.Range.Font.Name = "Times New Roman";
                        para1.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        para1.Range.Text = s_bin;
                        para1.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                        para1.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                        para1.Range.Font.Size = 14; //--- Выставили размер шрифта
                        para1.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                        para1.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ
                    }
                }

                if (groupBox11.Visible)
                {
                    para2 = worddoc.Paragraphs.Add(ref missing); //--- Добавляем новый параграф

                    //*** Добавляем разрыв страницы
                    para2.Range.InsertBreak(Word.WdBreakType.wdPageBreak);

                    if (!groupBox7.Visible) //--- Если нет данных по неописанным файлам - ставим заголовок
                    {
                        para2.Range.Text = "Отчёт от " + date + "г.";
                        object style_name = "Заголовок 1";
                        para2.Range.set_Style(ref style_name); //--- Задаём стиль
                        para2.Range.Font.Name = "Times New Roman";
                        para2.Range.Font.Position = 20; //--- Задаём расстояние между заголовком и след. строкой
                        para2.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; //--- Центруем
                        para2.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                        para2.Range.Font.Size = 20; //--- Выставили размер шрифта
                        para2.Range.Font.Bold = 1; //--- Показали, что строка будет жирной (очень жирной)
                        para2.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ
                    }

                    para2.Range.Font.Name = "Times New Roman";
                    if (!groupBox7.Visible) para2.Range.Text = "\n«О наличии не легитимности»";
                    else                    para2.Range.Text = "\n\n«О наличии не легитимности»";
                    para2.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                    para2.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    para2.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                    para2.Range.Font.Size = 14; //--- Выставили размер шрифта
                    para2.Range.Font.Bold = 1; //--- Показали, что строка будет жирной (очень жирной)
                    para2.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                    para2.Range.Font.Name = "Times New Roman";
                    para2.Range.Text = "\n";
                    para2.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                    para2.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    para2.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                    para2.Range.Font.Size = 14; //--- Выставили размер шрифта
                    para2.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                    para2.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                    para2.Range.Font.Name = "Times New Roman";
                    para2.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    if (!groupBox7.Visible)  para2.Range.Text = "\tСегодня, " + date + ", был произведён анализ директории (" + name_of_the_audited_directory + ") на наличие нелегитимных файлов, находящихся в директории.\n\tВ ходе проверки были выявлены следующие нелегитимные файлы (приведён список файлов).";
                    else                     para2.Range.Text = "\tБыл произведён анализ директории (" + name_of_the_audited_directory + ") на наличие нелегитимных файлов, находящихся в директории.\n\tВ ходе проверки были выявлены следующие нелегитимные файлы (приведён список файлов).";
                    para2.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                    para2.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    para2.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                    para2.Range.Font.Size = 14; //--- Выставили размер шрифта
                    para2.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                    para2.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                    int index = path_dir.LastIndexOf('\\'); //--- Нашли нужный нам разделитель

                    if (lbx_non_ligitimate_Dic.Items.Count != 0)
                    {
                        para2.Range.Font.Name = "Times New Roman";
                        para2.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        para2.Range.Text = "\nФайлы, находящиеся в директории, но не подключаемые к исполняемым:";
                        para2.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                        para2.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                        para2.Range.Font.Size = 14; //--- Выставили размер шрифта
                        para2.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                        para2.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                        string pth_non_legitimace_Dic_one = lbx_non_ligitimate_Dic.Items[0].ToString();
                        string mass_path_Dic = "";
                        int count_ind_Dic = 0;
                        for (int i = 0; i < lbx_list_all_files.Items.Count; i++)
                        {
                            string tmp = lbx_list_all_files.Items[i].ToString();
                            FileInfo fit = new FileInfo(tmp);

                            if (pth_non_legitimace_Dic_one == fit.Name)
                            {
                                string nameDir = fit.DirectoryName;
                                nameDir = nameDir.Substring(index);
                                mass_path_Dic += nameDir + '%';
                                count_ind_Dic++;
                                if (count_ind_Dic == lbx_non_ligitimate_Dic.Items.Count) break;
                                pth_non_legitimace_Dic_one = lbx_non_ligitimate_Dic.Items[count_ind_Dic].ToString();
                            }
                        }
                        pth_non_legitimace_Dic = mass_path_Dic.Split('%');


                        string s_Dic = "";
                        for (int i = 0; i < lbx_non_ligitimate_Dic.Items.Count; i++) s_Dic += Convert.ToString(i + 1) + ") " + lbx_non_ligitimate_Dic.Items[i] + " \n    Находится: " + pth_non_legitimace_Dic[i] + ";\n\n";

                        para2.Range.Font.Name = "Times New Roman";
                        para2.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        para2.Range.Text = s_Dic;
                        para2.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                        para2.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                        para2.Range.Font.Size = 14; //--- Выставили размер шрифта
                        para2.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                        para2.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ
                    }

                    if (lbx_non_ligitimate_Set.Items.Count != 0)
                    {
                        para2.Range.Font.Name = "Times New Roman";
                        para2.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        para2.Range.Text = "\n\nФайлы, подключаемые к исполняемым, но отсутствующие в директории:";
                        para2.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                        para2.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                        para2.Range.Font.Size = 14; //--- Выставили размер шрифта
                        para2.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                        para2.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ



                        string s_Set = "";
                        for (int i = 0; i < lbx_non_ligitimate_Set.Items.Count; i++)
                        {
                            string name_head = lbx_non_ligitimate_Set.Items[i].ToString(); //--- Получаем заголовочный файл
                            HashSet<string> name_exe = ClassCSharp.dic_name_header_file_Set[name_head]; //--- Получаем список исполняемых файлов, к которым подключается текущий заголовочный
                            int count_name_exe = name_exe.Count; //--- Получаем кол-во исполняемых файлов
                            string s_name_exe = "";
                            int j = 1;

                            foreach (string s in name_exe)
                            {
                                s_name_exe += "    " + Convert.ToString(i + 1) + "." + Convert.ToString(j) + ") " + s + ";\n";
                                j++;
                            }

                            s_Set += Convert.ToString(i + 1) + ") " + name_head + "\n    Подключается к файлам:\n" + s_name_exe + "\n";
                        }
                        para2.Range.Font.Name = "Times New Roman";
                        para2.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        para2.Range.Text = s_Set;
                        para2.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                        para2.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                        para2.Range.Font.Size = 14; //--- Выставили размер шрифта
                        para2.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                        para2.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ
                    }

                    if (lbx_False_header.Items.Count != 0)
                    {
                        para2.Range.Font.Name = "Times New Roman";
                        para2.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        para2.Range.Text = "\n\nФайлы с исполняемым кодом, маскируемые под заголовочные файлы:";
                        para2.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                        para2.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                        para2.Range.Font.Size = 14; //--- Выставили размер шрифта
                        para2.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                        para2.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ


                        string s = "";
                        for (int i = 0; i < lbx_False_header.Items.Count; i++) s += Convert.ToString(i + 1) + ") " + lbx_False_header.Items[i] + ";\n\n";

                        para2.Range.Font.Name = "Times New Roman";
                        para2.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        para2.Range.Text = s;
                        para2.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                        para2.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                        para2.Range.Font.Size = 14; //--- Выставили размер шрифта
                        para2.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                        para2.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ
                    }
                }

                para = worddoc.Paragraphs.Add(ref missing); //--- Добавляем новый параграф

                para.Range.Font.Name = "Times New Roman";
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                para.Range.Text = "\n\n\n\n";
                para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                para.Range.Font.Size = 14; //--- Выставили размер шрифта
                para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                para = worddoc.Paragraphs.Add(ref missing); //--- Добавляем новый параграф

                para.Range.Font.Name = "Times New Roman";
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                para.Range.Text = "Анализ проводил " + FIO + ".\nВ " + time + " " + date + "г.";
                para.Range.Font.Position = 0; //--- Задаём расстояние между заголовком и след. строкой
                para.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                para.Range.Font.Color = Word.WdColor.wdColorBlack; //--- Выставляем чёрный цвет
                para.Range.Font.Size = 14; //--- Выставили размер шрифта
                para.Range.Font.Bold = 0; //--- Показали, что строка будет жирной (очень жирной)
                para.Range.InsertParagraphAfter(); //--- Добавляем параграф в документ

                worddoc.SaveAs(ref fname, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                object save_changes = false;
                worddoc.Close(ref save_changes, ref missing, ref missing);
                wordapp.Quit(ref save_changes, ref missing, ref missing);
            }
        }

        //
        //*** Обработка галки, полволяющей вынести все текстовые файлы, содержащиеся в проверяемой директории, в отдельную папку
        //
        private void chbx_Text_file_in_separate_folder_CheckedChanged(object sender, EventArgs e)
        {
            if (chbx_Text_file_in_separate_folder.Checked) Flag_text_files_in_new_folder = true;
            else                                           Flag_text_files_in_new_folder = false;
        }
    }
}

//
//******  Класс ClassCSharp - реализация поиска исходных текстов на языке C#.
//
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Diplom;
using Word = Microsoft.Office.Interop.Word;
using Xceed.Words.NET;
using Xceed.Document.NET;


namespace _CSharp
{
    public class ClassCSharp
    {
        //--- Множество, по которому будет осуществляться поиск фрагментов кода Regex
        public static HashSet<String> strSet = new HashSet<String> { "using", "System;", "namespace", "public", "class", "InitializeComponent()", "private", "void", "//", "/**/", 
                                                                     "static", "break", "for (", "for(", "object", "sender", "struct", "~", "catch (IOException", "Main(", "args", "argv", "#include", "std", "<<", ">>", "\n", "close(", "{", "}", "return", "ret", "out" };

        //--- Массив с именами заголовочных файлов из стандартной библиотеки С и С++
        public static HashSet<String> mass_header_standart_lib = new HashSet<String> {"assert.h",  "ctupe.h",  "errno.h",    "float.h",   "iso646.h",      "limits.h",  "locale.h", "math.h",
                                                                                      "setjmp.h",  "signal.h", "stdarg.h",   "stddef.h",  "stdlib.h",      "string.h",  "time.h",   "wctype.h", "wchar.h",
                                                                                      "complex.h", "fenv.h",   "inttypes.h", "stdbool.h", "stdint.h",      "tgmath.h",  "stdio.h",  "malloc.h", "windows.h", "algorithm", "bitset",
                                                                                      "complex",   "deque",    "exception",  "fstrearn",  "functional",    "iomanip",   "ios",      "iosfwd",   "iostream",
                                                                                      "istream",   "iterator", "limits",     "list",      "locale",        "map",       "memory",   "new",      "numeric",   "ostream",
                                                                                      "queue",     "set",      "sstream",    "stack",     "stdexcept",     "streambuf", "string",   "typeinfo", "utility",
                                                                                      "valarray",  "vector",   "cassert",    "cctype",    "cerrno",        "cfloat",    "ciso646",  "climits",  "clocale",
                                                                                      "cmath",     "csetjmp",  "csignal",    "cstdarg",   "cstddef",       "cstdio",    "cstdlib",  "cstring",  "ctime",     "cwctype", "cwchar"};

        public static int count_coincidences = 0; //--- Счётчик совпадений для выявления потенциальных вхождений в текстовые файлы фрагментов исполняемого кода
        public static string[] smass; //--- Массив слов из считанной строки
        public static bool Flag_add_file = false; //--- Флаг, разрешающий добавление файла в список "К проверке"
        public static bool Flag_open_descripion = false; //---
        public static bool Flag_non_ligitimacy = false;
        public static bool Flag_False_header = false; //--- Флаг, сигнализирующий о том, что найден щпиён среди заголовочных файлов
        public static bool Flag_visibl_lbx_false_header = false;
        public static int count_False_header = 0;

        public static bool Flag_enter_descripion = false;
        public static bool Flag_enter_non_ligitimacy = false;

        public static bool Flag_found = false; //--- Флаг, сигнализирующий, что найден "#include"

        public static Dictionary<string, int> dic_executable = new Dictionary<string, int>(); //--- Словарь, в который будут записаны имена файлов с исходниками из полученной директирии и кол-во их упоминаний в "Описании" самого приложения
        public static Dictionary<string, int> dic_binary = new Dictionary<string, int>();
        public static Dictionary<string, int> dic_header_for_exe = new Dictionary<string, int>();
        public static Dictionary<string, HashSet<string>> dic_name_header_file_Set = new Dictionary<string, HashSet<string>>(); //--- Словарь, в которов ключём будет имя заголовочного файла, а значением - имена исполняемых файлов, в которых подключается данный заголовочный файл

        public static int count_zero_in_dic_header_for_exe = 0;
        public static int count_bin = 0;
        public static int count_exe = 0;

        public static HashSet<String> Set_connected_but_missing_header_files = new HashSet<String>(); //--- Множество, в которое будут заноситься заголовочные файлы, подключённые к исполняемым файлам, но отсутствующие в проверяемой директории

        public static string[] mass_indescribable_exe_file = new string[count_exe];
        public static string[] mass_indescribable_bin_file = new string[count_bin];
        public static string[] mass_indescribable_header_file = new string[count_zero_in_dic_header_for_exe];

        public static string[] massWord;
        public static string[] massTestDoc;
        public static string[] massWordHeader;

        public ClassCSharp()
        {

        }

        //
        //*** Функция по поиску вероятных вхождений фрагментов исполняемого кода на C# в текстовом файл Contains
        //
        public static bool Check_CS_InTxtFile(FileStream f_stream)
        {
            string s = ""; //--- Строка для считывания в него  
            StreamReader sr = new StreamReader(f_stream); //--- Поток на чтение
            string[] Separator = new string[] { " ", "<", "::", "\"", ".", "\t" }; //--- Разделитель в строке

            while ((s = sr.ReadLine()) != null) //--- Пока файл не кончился
            {
                smass = s.Split(Separator, StringSplitOptions.RemoveEmptyEntries);
                foreach (string i in smass)
                {
                    if (strSet.Contains(i)) count_coincidences++;

                    if (count_coincidences == 5)
                    {
                        Flag_add_file = true;
                        count_coincidences = 0;
                        break;
                    }
                }
                if (Flag_add_file) break;
            }
            return Flag_add_file;
        }

        //
        //*** Поиск исполняемого кода в документе Word doc
        //
        public static bool Check_CS_InDocFile(Word.Application wapp, string path) 
        {
            count_coincidences = 0;
            wapp.Visible = false; //--- Делаем приложение невидимым
            Object readOnly = true;

            Word.Document doc = wapp.Documents.Open(path, ref readOnly);
            string allWords = doc.Content.Text; //--- Считали содержимое файла в строковую переменную
            doc.Close();
            wapp.Quit();

            string[] Separator = new string[] { " ", "\r", "\t", "\n", "\f", "\b", "<", "::", "\"", "." };

            massWord = allWords.Split(Separator, StringSplitOptions.RemoveEmptyEntries); //--- получаем массив всех слов в документе
            foreach (string s in massWord)
            {
                if (strSet.Contains(s)) count_coincidences++; //--- Если текущее слово принадлежит множеству

                if (count_coincidences == 5)
                {
                    Flag_add_file = true;
                    count_coincidences = 0;
                    break;
                }
                if (Flag_add_file) break;
            }
            return Flag_add_file;
        }

        //
        //*** Поиск исполняемого кода в документе Word docx
        //
        public static bool Check_CS_InDocXFile(string fname)//(Word.Application wappX, string pth)
        {
            DocX docxLoad = DocX.Load(fname); //--- Загрузили себе файл
            count_coincidences = 0;

            string[] Separator = new string[] { " ", "\r", "\t", "\n", "\f", "\b", "<", "::", "\"", "." };

            var paragraphsList = docxLoad.Paragraphs; //--- Считали ко-во параграфов (каждый параграф - начинается с перевода на новую строку с помощью нажатия на "Enter")
            //var paragraphsText = paragraphsList[2]; //--- Запомнили третий параграф

            for (int i = 0; i < paragraphsList.Count; i++)
            {
                massTestDoc = paragraphsList[i].Text.Split(Separator, StringSplitOptions.RemoveEmptyEntries);

                for (int j = 0; j < massTestDoc.Length; j++)
                {
                    if (strSet.Contains(massTestDoc[j])) count_coincidences++;
                    if (count_coincidences == 5)
                    {
                        Flag_add_file = true;
                        break;
                    }
                }
                if (count_coincidences == 7) break;
            }

            //int noln = 0;

            return Flag_add_file;
        }

        //
        //*** Функция по проверке файлов, упомянутых в "Описании приложения", с файлами, находящимися непосредственно в дериктории 
        //
        public static void CheckDescriptionAndFiles(HashSet<string> deskript_name, ListBox LB_all)
        {
            dic_executable.Clear();
            dic_binary.Clear();

            //--- Проходим по ЛистБоксу
            for (int i = 0; i < LB_all.Items.Count; i++)
            {
                string fname_tmp = LB_all.Items[i].ToString();
                FileInfo fiLB = new FileInfo(fname_tmp);

                if (fiLB.Extension == ".c" || fiLB.Extension == ".cpp" || fiLB.Extension == ".h" || fiLB.Extension == ".cs")           dic_executable[fiLB.Name] = 0; //--- Делаем начальное заполнение словаря
                else if (fiLB.Extension == ".lib" || fiLB.Extension == ".exe" || fiLB.Extension == ".dll" || fiLB.Extension == ".bin") dic_binary[fiLB.Name] = 0; //--- Делаем начальное заполнение словаря
            }

            if (deskript_name.Count > 1)
            {
                foreach (string name in deskript_name)
                {
                    FileInfo fi = new FileInfo(name);

                    //*** Заполняем словарь значениями
                    if (fi.Extension == ".doc")
                    {
                        Word.Application wapp = new Word.Application();
                        wapp.Visible = false; //--- Делаем приложение невидимым
                        Object readOnly = true;

                        Word.Document doc = wapp.Documents.Open(name, ref readOnly);
                        string allWords = doc.Content.Text; //--- Считали содержимое файла в строковую переменную
                        doc.Close();
                        wapp.Quit();

                        massWord = allWords.Split(new char[] { ' ', ',', '?', ':', ';', '=', '+', '*', '\\', '|', '/', '\"', '<', '>', ']', '[', '\r', '\t', '\n', '\f', '\b' }, StringSplitOptions.RemoveEmptyEntries); //--- получаем массив всех слов в документе
                        foreach (string s in massWord)
                        {
                            if (dic_executable.ContainsKey(s)) //--- Если слово из файла есть в словаре
                            {
                                int tmp_exe = dic_executable[s]; //--- Получаем значение по ключу
                                tmp_exe++; //--- Увеличиваем число вхождений на 1
                                dic_executable[s] = tmp_exe;
                            }
                            else if (dic_binary.ContainsKey(s))
                            {
                                int tmp_bin = dic_binary[s]; //--- Получаем значение по ключу
                                tmp_bin++; //--- Увеличиваем число вхождений на 1
                                dic_binary[s] = tmp_bin;
                            }
                        }
                    }
                    else if (fi.Extension == ".docx")
                    {
                        DocX docxLoad = DocX.Load(name);

                        var paragraphsList = docxLoad.Paragraphs;

                        for (int i = 0; i < paragraphsList.Count; i++)
                        {
                            massTestDoc = paragraphsList[i].Text.Split(new char[] { ' ', ',', '?', ':', ';', '=', '+', '*', '\\', '|', '/', '\"', '<', '>', ']', '[', '\r', '\t', '\n', '\f', '\b' }, StringSplitOptions.RemoveEmptyEntries);

                            for (int j = 0; j < massTestDoc.Length; j++)
                            {
                                if (dic_executable.ContainsKey(massTestDoc[j])) //--- Если слово из файла есть в словаре
                                {
                                    int tmp_exe = dic_executable[massTestDoc[j]]; //--- Получаем значение по ключу
                                    tmp_exe++; //--- Увеличиваем число вхождений на 1
                                    dic_executable[massTestDoc[j]] = tmp_exe;
                                }
                                else if (dic_binary.ContainsKey(massTestDoc[j]))
                                {
                                    int tmp_bin = dic_binary[massTestDoc[j]]; //--- Получаем значение по ключу
                                    tmp_bin++; //--- Увеличиваем число вхождений на 1
                                    dic_binary[massTestDoc[j]] = tmp_bin;
                                }
                            }
                        }
                    }
                    else if (fi.Extension == ".txt")
                    {
                        FileStream fs = new FileStream(name, FileMode.Open, FileAccess.Read);
                        StreamReader sr = new StreamReader(fs);
                        string s;

                        while ((s = sr.ReadLine()) != null)
                        {
                            massWord = s.Split(new char[] { ' ', ',', '?', ':', ';', '=', '+', '*', '\\', '|', '/', '\"', '<', '>', ']', '[', '\r', '\t', '\n', '\f', '\b' }, StringSplitOptions.RemoveEmptyEntries);

                            for (int i = 0; i < massWord.Length; i++)
                            {
                                if (dic_executable.ContainsKey(massWord[i])) //--- Если слово из файла есть в словаре
                                {
                                    int tmp_exe = dic_executable[massWord[i]]; //--- Получаем значение по ключу
                                    tmp_exe++; //--- Увеличиваем число вхождений на 1
                                    dic_executable[massWord[i]] = tmp_exe;
                                }
                                else if (dic_binary.ContainsKey(massWord[i]))
                                {
                                    int tmp_bin = dic_binary[massWord[i]]; //--- Получаем значение по ключу
                                    tmp_bin++; //--- Увеличиваем число вхождений на 1
                                    dic_binary[massWord[i]] = tmp_bin;
                                }
                            }
                        }

                        fs.Close();
                        sr.Close();
                    }
                }
            }
            else if (deskript_name.Count == 1)
            {
                var tmp_list = new List<string>(deskript_name);
                FileInfo fi = new FileInfo(tmp_list[0]);

                //*** Заполняем словарь значениями
                if (fi.Extension == ".doc")
                {
                    Word.Application wapp = new Word.Application();
                    wapp.Visible = false; //--- Делаем приложение невидимым
                    Object readOnly = true;

                    Word.Document doc = wapp.Documents.Open(tmp_list[0], ref readOnly);
                    string allWords = doc.Content.Text; //--- Считали содержимое файла в строковую переменную
                    doc.Close();
                    wapp.Quit();

                    massWord = allWords.Split(new char[] { ' ', ',', '?', ':', ';', '=', '+', '*', '\\', '|', '/', '\"', '<', '>', ']', '[', '\r', '\t', '\n', '\f', '\b' }, StringSplitOptions.RemoveEmptyEntries); //--- получаем массив всех слов в документе
                    foreach (string s in massWord)
                    {
                        if (dic_executable.ContainsKey(s)) //--- Если слово из файла есть в словаре
                        {
                            int tmp_exe = dic_executable[s]; //--- Получаем значение по ключу
                            tmp_exe++; //--- Увеличиваем число вхождений на 1
                            dic_executable[s] = tmp_exe;
                        }
                        else if (dic_binary.ContainsKey(s))
                        {
                            int tmp_bin = dic_binary[s]; //--- Получаем значение по ключу
                            tmp_bin++; //--- Увеличиваем число вхождений на 1
                            dic_binary[s] = tmp_bin;
                        }
                    }
                }
                else if (fi.Extension == ".docx")
                {
                    DocX docxLoad = DocX.Load(tmp_list[0]);

                    var paragraphsList = docxLoad.Paragraphs;

                    for (int i = 0; i < paragraphsList.Count; i++)
                    {
                        massTestDoc = paragraphsList[i].Text.Split(new char[] { ' ', ',', '?', ':', ';', '=', '+', '*', '\\', '|', '/', '\"', '<', '>', ']', '[', '\r', '\t', '\n', '\f', '\b' }, StringSplitOptions.RemoveEmptyEntries);

                        for (int j = 0; j < massTestDoc.Length; j++)
                        {
                            if (dic_executable.ContainsKey(massTestDoc[j])) //--- Если слово из файла есть в словаре
                            {
                                int tmp_exe = dic_executable[massTestDoc[j]]; //--- Получаем значение по ключу
                                tmp_exe++; //--- Увеличиваем число вхождений на 1
                                dic_executable[massTestDoc[j]] = tmp_exe;
                            }
                            else if (dic_binary.ContainsKey(massTestDoc[j]))
                            {
                                int tmp_bin = dic_binary[massTestDoc[j]]; //--- Получаем значение по ключу
                                tmp_bin++; //--- Увеличиваем число вхождений на 1
                                dic_binary[massTestDoc[j]] = tmp_bin;
                            }
                        }
                    }
                }
                else if (fi.Extension == ".txt")
                {
                    FileStream fs = new FileStream(tmp_list[0], FileMode.Open, FileAccess.Read);
                    StreamReader sr = new StreamReader(fs);
                    string s;

                    while ((s = sr.ReadLine()) != null)
                    {
                        massWord = s.Split(new char[] { ' ', ',', '?', ':', ';', '=', '+', '*', '\\', '|', '/', '\"', '<', '>', ']', '[', '\r', '\t', '\n', '\f', '\b' }, StringSplitOptions.RemoveEmptyEntries);

                        for (int i = 0; i < massWord.Length; i++)
                        {
                            if (dic_executable.ContainsKey(massWord[i])) //--- Если слово из файла есть в словаре
                            {
                                int tmp_exe = dic_executable[massWord[i]]; //--- Получаем значение по ключу
                                tmp_exe++; //--- Увеличиваем число вхождений на 1
                                dic_executable[massWord[i]] = tmp_exe;
                            }
                            else if (dic_binary.ContainsKey(massWord[i]))
                            {
                                int tmp_bin = dic_binary[massWord[i]]; //--- Получаем значение по ключу
                                tmp_bin++; //--- Увеличиваем число вхождений на 1
                                dic_binary[massWord[i]] = tmp_bin;
                            }
                        }
                    }

                    fs.Close();
                    sr.Close();
                }
            }
            

            //*** Обход словаря в поиске нулевых значений
            foreach (string key in dic_executable.Keys)
            {
                int int_tmp = dic_executable[key];
                if (int_tmp == 0)
                {
                    Form1.Flag_not_converge = true; //--- Как только получили нулевое значение (хотябы одно) - выставляем флаг (Закончил здесь - 15.09.2020 16:55
                    mass_indescribable_exe_file = dic_executable.Where(r => r.Value == 0).Select(r => r.Key).ToArray();
                    count_exe++;
                }
            }


            //*** Обход словаря в поиске нулевых значений
            foreach (string key in dic_binary.Keys)
            {
                int int_tmp = dic_binary[key];
                if (int_tmp == 0)
                {
                    Form1.Flag_not_converge = true; //--- Как только получили нулевое значение (хотябы одно) - выставляем флаг (Закончил здесь - 15.09.2020 16:55
                    mass_indescribable_bin_file = dic_binary.Where(r => r.Value == 0).Select(r => r.Key).ToArray();
                    count_bin++;
                }
            }

            if (Form1.Flag_not_converge == false) Form1.Flag_converge = true; //--- Если всё совпадает
        }

        //
        //*** Функция, проверяющая: описаны ли заголовочные (и в дальнейшем - ещё и бинарные) файлы, которые подключаются в исполняемых
        //
        public static void CheckConnectingHeaderOrBinaryFiles(ListBox LB_Header, ListBox LB_Executable, ListBox LB_False_H)
        {
            dic_header_for_exe.Clear();
            Set_connected_but_missing_header_files.Clear();
            dic_name_header_file_Set.Clear();

            //--- Проходим по ЛистБоксу c заголовочными файлами
            for (int i = 0; i < LB_Header.Items.Count; i++)
            {
                string fname_tmp = LB_Header.Items[i].ToString();
                FileInfo fiLB = new FileInfo(fname_tmp);
                if (dic_header_for_exe.ContainsKey(fiLB.Name)) //--- Если найден файл, который уже находится в словаре (но имеет иную родительскую поддиректорию) 
                {
                    //var t = fiLB.Directory.Name;
                    //string tmp_s = fiLB.Name + "@" + t + Convert.ToString(i);
                    //dic_header_for_exe[tmp_s] = 0;
                }
                else dic_header_for_exe[fiLB.Name] = 0; //--- Делаем начальное заполнение словаря


                //--- Проверка на то - является ли файл в действительности заголовочным
                string[] mass_line = File.ReadAllLines(LB_Header.Items[i].ToString());

                foreach (string s in mass_line)
                {
                    string[] m_str = s.Split(' ', '#', ',');

                    foreach (string s_in in m_str)
                    {
                        if (s_in == "main(argv" || s_in == "main(args" || s_in == "main()" || s_in == "main(argc") //--- Если присутствует точка входа
                        {
                            Flag_False_header = true;
                            Flag_visibl_lbx_false_header = true;

                            LB_False_H.Items.Add(LB_Header.Items[i]);
                            count_False_header++;
                            break;
                        }
                    }

                    if (Flag_False_header)
                    {
                        Flag_False_header = false;
                        break;
                    }
                }
            }

            HashSet<string> set = new HashSet<string>();

            //--- Проходим по ЛистБоксу c исполняемыми файлами
            for (int i = 0; i < LB_Executable.Items.Count; i++)
            {
                string fname_tmp = LB_Executable.Items[i].ToString();
                FileInfo fiLB = new FileInfo(fname_tmp);
                if (fiLB.Extension == ".cpp" || fiLB.Extension == ".c")
                {
                    massWordHeader = File.ReadAllLines(fname_tmp);

                    foreach (string s_outer in massWordHeader)
                    {
                        string[] mass_str = s_outer.Split(' ');

                        foreach (string s_inner in mass_str)
                        {
                            if (s_inner == "#include")
                            {
                                Flag_found = true;
                                continue;
                            }

                            if (Flag_found) //--- Если выставлен флаг - получаем след. слово
                            {
                                //*** Блок проверки на то - действительно ли полученное слово явл. заголовочным
                                string tmp = s_inner; //--- Получаем тек. слово
                                int lenght = tmp.Length; //--- Получаем длинну
                                if ((tmp[0] == '"' || tmp[0] == '<') && (tmp[lenght - 1] == '>' || tmp[lenght - 1] == '"')) //--- Если полученное слово начинается на признак объявления подключаемого файла 
                                {
                                    string[] @out = s_inner.Split('"', '<', '>'); //--- Отделяем от ковычек (наше сокровище - на позиции 1)

                                    if (@out[1].Contains('/')) //--- Если есть признак присутствия поддиректории
                                    {
                                        int index = @out[1].LastIndexOf('/'); //--- Находим индекс последнего вхождения обратного слэша
                                        string tmp_name_without_slash = @out[1].Substring(index + 1); //--- Обрезаем ненужное - оставляем только само имя подключаемого файла с расширением
                                        @out[1] = tmp_name_without_slash;
                                    }

                                    if (dic_header_for_exe.ContainsKey(@out[1])) //--- Если заголовочный файл есть в словаре, а значит - есть в директории
                                    {
                                        int tmp_hed = dic_header_for_exe[@out[1]]; //--- Получаем значение по ключу
                                        tmp_hed++; //--- Увеличиваем число вхождений на 1
                                        dic_header_for_exe[@out[1]] = tmp_hed;
                                    }
                                    else
                                    {
                                        if (Set_connected_but_missing_header_files.Contains(@out[1]) == false) //--- Если заголовочный файл подключён к исполняемому, при этом - его нет в проверяемой директории и он ещё не внесён в множество
                                        {
                                            Set_connected_but_missing_header_files.Add(@out[1]);
                                            Flag_non_ligitimacy = true;
                                        }

                                        if (!dic_name_header_file_Set.ContainsKey(@out[1]))  //--- Если заголовочного файла нет в словаре <имя заголовочного файла - список исполняемых файлов (в которых он подключён)> - Создаём новую пару <имя - множество>  
                                        {
                                            dic_name_header_file_Set[@out[1]] = new HashSet<string>();
                                        }

                                        int index = Form1.path_dir.LastIndexOf('\\'); //--- Нашли нужный нам разделитель
                                        string nameF = fname_tmp.Substring(index); //--- Удалили всё лишнее - теперь путь до файла начинается с имени директории, которую мы проверяем

                                        set = dic_name_header_file_Set[@out[1]]; //--- получаем множество по ключу  
                                        set.Add(nameF); //--- добавляем в множество новый элемент
                                        dic_name_header_file_Set[@out[1]] = set; //--- заносим изменённое множество обратно в словарь
                                    }
                                    Flag_found = false;
                                }
                                else
                                {
                                    Flag_found = false;
                                    continue; 
                                } 
                            }
                        }
                    }
                }
            }

            //*** Обход словаря в поиске нулевых значений
            foreach (string key in dic_header_for_exe.Keys)
            {
                int int_tmp = dic_header_for_exe[key];
                if (int_tmp == 0)
                {
                    Flag_non_ligitimacy = true;
                    mass_indescribable_header_file = dic_header_for_exe.Where(r => r.Value == 0).Select(r => r.Key).ToArray();
                    count_zero_in_dic_header_for_exe++;
                }
            }
        }
    }
}

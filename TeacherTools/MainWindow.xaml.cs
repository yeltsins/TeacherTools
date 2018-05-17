using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;

namespace TeacherTools
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        List<(int,int,string)> Questions = new List<(int,int,string)>();
        List<(string, string)> Bilets = new List<(string, string)>();
        static Random rnd = new Random();

        public MainWindow()
        {

            InitializeComponent();

            xlsPath.Text = @"C:\Users\MediaMarkt\Google Диск\Навигация роботов\Рабочая программа\Вопросы к экзамену.xlsx";
            wordPath.Text = @"C:\Users\MediaMarkt\Google Диск\Навигация роботов\Рабочая программа\Экзаменационные билеты.docx";


        }


        private void xlsFileChoose(object sender, RoutedEventArgs e)
        {
            OpenFileDialog OPF = new OpenFileDialog();
            OPF.Title = "Выбор таблицы с вопросами";
            OPF.Filter = "Файлы xlsx|*.xlsx|Файлы xls|*.xls";
            if (OPF.ShowDialog() == true)
            {
                xlsPath.Text = OPF.FileName;
            }
        }

        private void createQuestions(object sender, RoutedEventArgs e)
        {
            Questions.Clear();
            Bilets.Clear();
            //Создаём приложение.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(xlsPath.Text, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
            //Выбираем область таблицы. (в нашем случае просто ячейку)
            int i = 1;
            string cellStr = ObjWorkSheet.get_Range("B2", "B2").Text.ToString();
            int cellInt;

            //Если игнорируем темы
            if (ignoreTopic.IsChecked.Value)
            {
                cellInt = 1;
            }
            else
            {
                cellInt = Convert.ToInt32(ObjWorkSheet.get_Range("A2", "A2").Text.ToString());
            }
            
            while (cellStr != "")
            {

                Questions.Add((i, cellInt, cellStr));
                if (ignoreTopic.IsChecked.Value)
                {
                    cellInt = i+1;
                }
                else
                {
                    Int32.TryParse(ObjWorkSheet.get_Range("A" + (i + 2).ToString(), "A" + (i + 2).ToString()).Text.ToString(), out cellInt);
                }
                cellStr = ObjWorkSheet.get_Range("B" + (i + 2).ToString(), "B" + (i + 2).ToString()).Text.ToString();
                i++;
            }

                            
    
            //Удаляем приложение (выходим из экселя) - ато будет висеть в процессах!
            ObjExcel.Quit();

            //Найдем количество билетов
            int numQuest = (int)Math.Floor(Questions.Count/2.0);



            //Определяем комбинации билетов
            for(i=1;i<= numQuest; i++)
            {
                var firstQ = Questions[rnd.Next(Questions.Count)];
                var secondQ = firstQ;
                Questions.RemoveAll(a => a.Item1 == firstQ.Item1);
                if (Questions.FindAll(a => a.Item2 != firstQ.Item2).Count == 0)
                {
                    secondQ = Questions[rnd.Next(Questions.Count)];
                }
                else
                {
                    secondQ = Questions.FindAll(a => a.Item2 != firstQ.Item2)[rnd.Next(Questions.FindAll(b => b.Item2 != firstQ.Item2).Count())];
                }
                Questions.RemoveAll(a => a.Item1 == secondQ.Item1);
                Bilets.Add((firstQ.Item3, secondQ.Item3));
                
            }
            //MessageBox.Show("Номер билета: " + i + "\n" + Bilets.Last().Item1 + "\n" + Bilets.Last().Item2 + "\n осталось вопросов: " + Questions.Count);
            
            //Открыли шаблон
            Word.Application app = new Word.Application();            
            app.Documents.Open(wordPath.Text);


            //Выделили и скопировали таблицу
            app.ActiveDocument.Tables[1].Range.Copy();

            Object missing = Type.Missing;
            Object wrap = Word.WdFindWrap.wdFindContinue;
            Object replace = Word.WdReplace.wdReplaceAll;
            Object nullobj = System.Reflection.Missing.Value;
            Object objBreak = Word.WdBreakType.wdPageBreak;
            Object objUnit = Word.WdUnits.wdStory;
            Word.Find find = app.Selection.Find;

            for (i = 1; i <= numQuest; i++)
            {
                if (i != 1)
                {
                    //Вставляем новую таблицу в конец                    
                    app.Selection.EndKey(ref objUnit, ref nullobj);
                    app.Selection.Paste();
                }                
                
                if (i % 2 == 1)
                {
                    if (i != 1)
                    {
                        //Вставляем в конец новый параграф
                        app.ActiveDocument.SelectAllEditableRanges();
                        app.Selection.EndKey();
                        app.Selection.InsertParagraphAfter();
                        app.Selection.InsertParagraphAfter();
                    }

                }
                else
                {
                    if (i != numQuest)
                    {
                        //Вставляем в конец новый параграф
                        app.Selection.EndKey(ref objUnit, ref nullobj);
                        app.Selection.InsertParagraphAfter();
                    }
                }

                //Заменили вопрос 1 и 2

                find = app.Selection.Find;

                find.Text = "Вопрос 1";
                find.Replacement.Text = Bilets[i-1].Item1;

                find.Execute(FindText: Type.Missing,
                    MatchCase: false,
                    MatchWholeWord: false,
                    MatchWildcards: false,
                    MatchSoundsLike: missing,
                    MatchAllWordForms: false,
                    Forward: true,
                    Wrap: wrap,
                    Format: false,
                    ReplaceWith: missing, Replace: replace);

                find = app.Selection.Find;

                find.Text = "Вопрос 2";
                find.Replacement.Text = Bilets[i-1].Item2;

                find.Execute(FindText: Type.Missing,
                    MatchCase: false,
                    MatchWholeWord: false,
                    MatchWildcards: false,
                    MatchSoundsLike: missing,
                    MatchAllWordForms: false,
                    Forward: true,
                    Wrap: wrap,
                    Format: false,
                    ReplaceWith: missing, Replace: replace);

                find = app.Selection.Find;

                find.Text = "№ 01";
                find.Replacement.Text = "№ "+i;

                find.Execute(FindText: Type.Missing,
                    MatchCase: false,
                    MatchWholeWord: false,
                    MatchWildcards: false,
                    MatchSoundsLike: missing,
                    MatchAllWordForms: false,
                    Forward: true,
                    Wrap: wrap,
                    Format: false,
                    ReplaceWith: missing, Replace: replace);

            }

            //Закрываем ворд
            app.ActiveDocument.SaveAs2(wordPath.Text.Remove(wordPath.Text.Count()-5,5) +"_generated.docx");
            app.ActiveDocument.Close();
            app.Quit();

            MessageBox.Show("Файл с билетами сгенерирован.");
            
        }

        private void wordFileChoose(object sender, RoutedEventArgs e)
        {
            OpenFileDialog OPF = new OpenFileDialog();
            OPF.Title = "Выбор файла с шаблоном билета";
            OPF.Filter = "Файлы docx|*.docx|Файлы doc|*.doc";
            if (OPF.ShowDialog() == true)
            {
                wordPath.Text = OPF.FileName;
            }
        }

        private void wordReplaceText(string find, string replace)
        {

        }

    }
}

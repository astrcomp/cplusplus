using System;
using System.IO;
using System.Xml.Linq;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;


namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            form2 = new Form2(this);
        }
        Form2 form2;

        private void Form1_Load(object sender, EventArgs e)
        {
            string[] text = File.ReadAllLines("data\\0.txt");
           comboBox1.Items.AddRange(text);
           comboBox1.SelectedIndex = 0;
            dataGridView1.AutoGenerateColumns = true;
            
            string []text1 = File.ReadAllLines("data\\16.txt");
            
/*0*/       dataGridView1.Columns.Add("Cat", "0. Каталог");
/*1*/       dataGridView1.Columns.Add("Name", "1. Копии доков");
/*2*/       dataGridView1.Columns.Add("ves", "2. Ед.");
/*3*/       dataGridView1.Columns.Add("Kol", "3. Тех. описание");
/*4*/       dataGridView1.Columns.Add("Teh", "4. Производитель");
/*5*/       dataGridView1.Columns.Add("Product", "5. Товарный знак");
/*6*/       dataGridView1.Columns.Add("Kod", "6. работах, услугах");
/*7*/       dataGridView1.Columns.Add("TR", "7. Колличество.");
            dataGridView1.Columns.Add("Instr", "8. ТР.ТС.");
            dataGridView1.Columns.Add("Zarub", "9. Копии разрешит.");
            dataGridView1.Columns.Add("Vubor", "10. Сертификаты");
            dataGridView1.Columns.Add("Rus", "11. Заруб. аналоги");
            dataGridView1.Columns.Add("sved", "12. Обоснование выбора");
            dataGridView1.Columns.Add("analiz", "13. Закупочная документация");
            dataGridView1.Columns.Add("sogl", "14. Отечественный произв.");
            dataGridView1.Columns.Add("status", "15. ОНМ/ПЭН");
            dataGridView1.Columns.Add("status1", "16. Коды");
            dataGridView1.Columns[1].Width= 80;
            dataGridView1.Columns[2].Width = 80;
            dataGridView1.Columns[3].Width = 100;
            for (int i = 0; i <text.Length-1;i++)
            {
                dataGridView1.Rows.Add();
            }
            for (int i = 0; i < 18; ++i)
            {
                text = File.ReadAllLines("data\\" + i.ToString() + ".txt");                
                for (int y = 0; y < text.Length; ++y)
                {                    
                    dataGridView1.Rows[y].Cells[i].Value = text[y];
                }
            }


            



        }

        private void button1_Click(object sender, EventArgs e)
        {
            string CurrentDir = AppDomain.CurrentDomain.BaseDirectory;
            if (!Directory.Exists(CurrentDir + "Заявки"))
            {
                Directory.CreateDirectory(CurrentDir + "Заявки");
            }          

            if (!Directory.Exists(CurrentDir + "Заявки\\" + comboBox1.Text))// dataGridView1.Rows[0].Cells[0].Value.ToString()))
            {
                Directory.CreateDirectory(CurrentDir + "Заявки\\" + comboBox1.Text); // dataGridView1.Rows[0].Cells[0].Value.ToString());
            }

            if (!Directory.Exists(CurrentDir + "Заявки\\" + comboBox1.Text + "\\Приложение"))// dataGridView1.Rows[0].Cells[0].Value.ToString()))
            {
                Directory.CreateDirectory(CurrentDir + "Заявки\\" + comboBox1.Text + "\\Приложение"); // dataGridView1.Rows[0].Cells[0].Value.ToString());
            }

            for (int i = 0; i < 14; i++)
            {
                i++;
                if (!Directory.Exists(CurrentDir + "Заявки\\" + comboBox1.Text + "\\Приложение\\" + i.ToString()))
                {                   
                    Directory.CreateDirectory(CurrentDir + "Заявки\\" + comboBox1.Text  + "/Приложение/" + i.ToString());                
                }
                i--;
            }

            Word.Application word = new Word.Application(); //создаем COM-объект Word   
            Object missingObj = System.Reflection.Missing.Value;
            Object trueObj = true;
            Object falseObj = false;
            word.Visible = true;
            Word.Document document = new Word.Document();            
            Object templatePathObj = CurrentDir + "shablon1.dotx";

            try
            {
                document = word.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
            }
            catch (Exception error)
            {
                document.Close(ref falseObj, ref missingObj, ref missingObj);
                word.Quit(ref missingObj, ref missingObj, ref missingObj);
                document = null;
                word = null;
                throw error;
            }

            int index = comboBox1.SelectedIndex;
            object strToFindObj = "name";
//            string str  = dataGridView1.Rows[index].Cells[1].Value.ToString();
 //           str = str.Remove(0, 10);
            object replaceStrObj = dataGridView1.Rows[index].Cells[1].Value.ToString(); ;
            object replaceTypeObj;
            Word.Range wordRange;            
            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            for (int i =1; i<= document.Sections.Count;i++)
            {
                wordRange = document.Sections[i].Range;
                Word.Find wordFindObj = wordRange.Find;
                object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceStrObj, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);
            }

            strToFindObj = "kod";
            replaceStrObj = dataGridView1.Rows[index].Cells[6].Value.ToString();
            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            for (int i = 1; i <= document.Sections.Count; i++)
            {
                wordRange = document.Sections[i].Range;
                Word.Find wordFindObj = wordRange.Find;
                object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceStrObj, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);
            }

            strToFindObj = "prod";
            replaceStrObj = dataGridView1.Rows[index].Cells[4].Value.ToString();
            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            for (int i = 1; i <= document.Sections.Count; i++)
            {
                wordRange = document.Sections[i].Range;
                Word.Find wordFindObj = wordRange.Find;
                object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceStrObj, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);
            }

            strToFindObj = "kol";
            replaceStrObj = dataGridView1.Rows[index].Cells[3].Value.ToString();
            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            for (int i = 1; i <= document.Sections.Count; i++)
            {
                wordRange = document.Sections[i].Range;
                Word.Find wordFindObj = wordRange.Find;
                object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceStrObj, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);
            }
            document.SaveAs(CurrentDir + "Заявки\\" + comboBox1.Text +"\\"+ comboBox1.Text+ ".docx");
            document.Close();

            // Второй файл ТАБЛИЦА
            templatePathObj = CurrentDir + "shablon.dotx";
            try
            {
                document = word.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
            }
            catch (Exception error)
            {
                document.Close(ref falseObj, ref missingObj, ref missingObj);
                word.Quit(ref missingObj, ref missingObj, ref missingObj);
                document = null;
                word = null;
                throw error;
            }

            strToFindObj = "name";
            replaceStrObj = dataGridView1.Rows[index].Cells[1].Value.ToString();
            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            for (int i = 1; i <= document.Sections.Count; i++)
            {
                wordRange = document.Sections[i].Range;
                Word.Find wordFindObj = wordRange.Find;
                object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceStrObj, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);
            }
            strToFindObj = "kol";
            replaceStrObj = dataGridView1.Rows[index].Cells[3].Value.ToString();
            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            for (int i = 1; i <= document.Sections.Count; i++)
            {
                wordRange = document.Sections[i].Range;
                Word.Find wordFindObj = wordRange.Find;
                object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceStrObj, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);
            }
            strToFindObj = "prod";
            replaceStrObj = dataGridView1.Rows[index].Cells[5].Value.ToString();
            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            for (int i = 1; i <= document.Sections.Count; i++)
            {
                wordRange = document.Sections[i].Range;
                Word.Find wordFindObj = wordRange.Find;
                object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceStrObj, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);
            }

            strToFindObj = "teh";
            replaceStrObj = dataGridView1.Rows[index].Cells[4].Value.ToString();
            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            for (int i = 1; i <= document.Sections.Count; i++)
            {
                wordRange = document.Sections[i].Range;
                Word.Find wordFindObj = wordRange.Find;
                object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceStrObj, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);
            }


            strToFindObj = "img";
            replaceStrObj = dataGridView1.Rows[index].Cells[5].Value.ToString();
            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            Word.Tables tables = document.Tables;
            Word.Table table = tables[1];
            wordRange = table.Rows[8].Cells[3].Range;
            wordRange.Text = "";

            string strvr = comboBox1.Text;  
            strvr = strvr.Replace("Заявка", "Приложение к Заявке");
            templatePathObj = CurrentDir + "Новая папка\\" + comboBox1.Text +  "\\" +strvr +".docx";

            //if (!File.Exists(templatePathObj.ToString()))
            //{
            //    Word.Document doc1 = word.Documents.Add();
            //    doc1.SaveAs(templatePathObj.ToString());
            //    doc1.Close();
            //}
                

            Word.Document doc = word.Documents.Open(templatePathObj);
            if (doc.InlineShapes.Count > 0)
            {
                Word.InlineShape pic = doc.InlineShapes[1];
                pic.Height = 100;
                pic.Width = 100;
                pic.Select();
                doc.ActiveWindow.Selection.CopyAsPicture();
                wordRange.Paste();
           //     Clipboard.Clear();
            }

            if (doc.Shapes.Count > 0)
            {
                Word.Shape pic = doc.Shapes[1];
                pic.Height = 100;
                pic.Width = 100;
                pic.Select();
                doc.ActiveWindow.Selection.CopyAsPicture();
                wordRange.Paste();
               // Clipboard.Clear();
                document.Shapes[1].WrapFormat.Type =  Word.WdWrapType.wdWrapInline;
                
            }



            doc.Close();


            wordRange = table.Rows[9].Cells[3].Range;
            wordRange.Text = dataGridView1.Rows[comboBox1.SelectedIndex].Cells[7].Value.ToString();

            wordRange = table.Rows[10].Cells[3].Range;
            wordRange.Text = dataGridView1.Rows[comboBox1.SelectedIndex].Cells[8].Value.ToString();

            wordRange = table.Rows[11].Cells[3].Range;
            wordRange.Text = dataGridView1.Rows[comboBox1.SelectedIndex].Cells[9].Value.ToString();

            wordRange = table.Rows[12].Cells[3].Range;
            wordRange.Text = dataGridView1.Rows[comboBox1.SelectedIndex].Cells[10].Value.ToString();


            wordRange = table.Rows[13].Cells[3].Range;
            wordRange.Text = dataGridView1.Rows[comboBox1.SelectedIndex].Cells[11].Value.ToString();

            wordRange = table.Rows[14].Cells[3].Range;
            wordRange.Text = dataGridView1.Rows[comboBox1.SelectedIndex].Cells[12].Value.ToString();

            wordRange = table.Rows[15].Cells[3].Range;
            wordRange.Text = dataGridView1.Rows[comboBox1.SelectedIndex].Cells[13].Value.ToString();

            wordRange = table.Rows[16].Cells[3].Range;
            wordRange.Text = dataGridView1.Rows[comboBox1.SelectedIndex].Cells[14].Value.ToString();

            //Bitmap img = new Bitmap( dataGridView1.Rows[1].Cells[6].Value);


            // Clipboard.SetImage(Image.FromStream(new MemoryStream((byte[])dataGridView1.Rows[0].Cells[6].Value)));
            // Image img = Image.FromStream(new MemoryStream((byte[])dataGridView1.Rows[0].Cells[6].Value));
            document.SaveAs(CurrentDir + "Заявки\\" + comboBox1.Text + "\\Таблица " + comboBox1.Text + ".docx"); ;
            document.Close();
            
            

            ///сохранить в 1 папке 
            string[] str = File.ReadAllLines("vopros.txt");
            for (int i = 1; i<15; i++)
            {
                templatePathObj = CurrentDir + "Заявки\\" + comboBox1.Text + "\\Приложение\\" + i;
                string filename = "Пояснение " + i + ".docx";
                
                if (!File.Exists(templatePathObj+"\\"+filename))
                {                    
                    Word.Document vremdoc = word.Documents.Add();
                    vremdoc.Content.SetRange(0, 0);
                    vremdoc.Content.Text = i.ToString()+". " +str[i-1];
                    vremdoc.Paragraphs.Add();
                    if (dataGridView1.Rows[comboBox1.SelectedIndex].Cells[i].Value != null)
                    {
                        vremdoc.Paragraphs[2].Range.Text = dataGridView1.Rows[comboBox1.SelectedIndex].Cells[i].Value.ToString();
                    }
                    vremdoc.SaveAs(templatePathObj + "\\" + filename);
                    vremdoc.Close();
                }
                else
                {
                    File.Delete(templatePathObj + "\\" + filename);
                    Word.Document vremdoc = word.Documents.Add();
                    vremdoc.Content.SetRange(0, 0);
                    vremdoc.Content.Text = i.ToString() + ". " + str[i - 1];
                    vremdoc.Paragraphs.Add();
                    if (dataGridView1.Rows[comboBox1.SelectedIndex].Cells[i].Value != null)
                    {
                        vremdoc.Paragraphs[2].Range.Text = dataGridView1.Rows[comboBox1.SelectedIndex].Cells[i].Value.ToString();
                    }
                    vremdoc.SaveAs(templatePathObj + "\\" + filename);
                    vremdoc.Close();
                }
            }

            ///////Втсавляем картинку
            templatePathObj = CurrentDir + "Заявки\\" + comboBox1.Text + "\\Приложение\\" + "6";
            string filename1 = "Пояснение " + "6" + ".docx";
            File.Delete(templatePathObj + "\\" + filename1);
            Word.Document vremdoc1 = word.Documents.Add();
            vremdoc1.Content.SetRange(0, 0);
            vremdoc1.Content.Text = 6.ToString() + ". " + str[6 - 1];
            vremdoc1.Paragraphs.Add();

            vremdoc1.Paragraphs[2].Range.Paste();
            vremdoc1.SaveAs(templatePathObj + "\\" + filename1);
            vremdoc1.Close();





            templatePathObj = CurrentDir + "Заявки\\" + comboBox1.Text + "\\Приложение\\" + "3";
            filename1 = "Пояснение " + "3" + ".docx";
            File.Delete(templatePathObj + "\\" + filename1);
            vremdoc1 = word.Documents.Add();
            vremdoc1.Content.SetRange(0, 0);
            vremdoc1.Content.Text = 3.ToString() + ". " + str[3 - 1];
            vremdoc1.Paragraphs.Add();
            if (dataGridView1.Rows[comboBox1.SelectedIndex].Cells[3].Value != null)
            {
                vremdoc1.Paragraphs[2].Range.Text = dataGridView1.Rows[comboBox1.SelectedIndex].Cells[1].Value.ToString()+ " - "+dataGridView1.Rows[comboBox1.SelectedIndex].Cells[3].Value.ToString()+" шт.";
            }
            vremdoc1.SaveAs(templatePathObj + "\\" + filename1);
            vremdoc1.Close();


            Clipboard.Clear();
            word.Quit();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string s = textBox1.Text;
            File.AppendAllText("text.txt", s + Environment.NewLine);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            string[] text = File.ReadAllLines("text.txt");
            comboBox1.Items.AddRange(text);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Text = comboBox1.Text;        
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }
        
        private void button5_Click(object sender, EventArgs e)
        {
            string path = null;
            using (var dialog = new FolderBrowserDialog())
            if (dialog.ShowDialog() == DialogResult.OK) path = dialog.SelectedPath;
            String[] dir = Directory.GetDirectories(path);
            int size = dir.Length;
            for (int i = 0; i< size; i++)
            {
                dir[i] = dir[i].Replace(path + "\\","");
            }
            File.AppendAllLines("dir.txt",dir);
            form2.textBox2.Text = "HI!";
            form2.Show();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Form2 newform = new Form2(this);
            if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
            {
                newform.textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                
            }
            newform.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string strvr = comboBox1.Text;  // + "Приложение к Заявке";
            strvr = strvr.Replace("Заявка", "Приложение к Заявке");
            MessageBox.Show(strvr);
            return;
            Word.Application word = new Word.Application();
            word.Visible = true;
            Word.Document doc = word.Documents.Open("C:\\Users\\Администратор\\Desktop\\WindowsFormsApplication1\\Пояснение 10.docx");
            doc.Content.Delete();
            word.Quit();            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string CurrentDir = AppDomain.CurrentDomain.BaseDirectory;
            string PathData = CurrentDir + "data\\";
            
            if (File.Exists(PathData + "prilojenie.txt"))
            {
                string[] stroka = File.ReadAllLines(PathData + "prilojenie.txt");
                Word.Application AppWord = new Word.Application();
                Word.Document doc = new Word.Document();
                doc = AppWord.Documents.Add();
                AppWord.Visible = true;
                doc.Paragraphs.Add();
                doc.Paragraphs[1].Range.Text = stroka[0];
                doc.Paragraphs[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                doc.Paragraphs[1].Range.Bold=-1;
                doc.Paragraphs[1].Range.Font.Size = 14;
                doc.Paragraphs[1].Range.Font.Name = "Times New Roman";
                
                int j = 2;

                for (int i = 1; i < stroka.Length; i++)
                {
                    doc.Paragraphs.Add();
                    
                    doc.Paragraphs[j].Range.ParagraphFormat.CharacterUnitFirstLineIndent = 1;
                    doc.Paragraphs[j].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                    doc.Paragraphs[j].Range.Bold = 0;
                    doc.Paragraphs[j].Range.Text = stroka[i];
                    j++;

                    doc.Paragraphs.Add();
                    if (dataGridView1.Rows[comboBox1.SelectedIndex].Cells[i].Value == null) dataGridView1.Rows[comboBox1.SelectedIndex].Cells[i].Value = "a";


                    doc.Paragraphs[j].Range.Text = dataGridView1.Rows[comboBox1.SelectedIndex].Cells[i].Value.ToString();


                    if (j == 11)
                    {
                        string strvr = comboBox1.Text;
                        strvr = strvr.Replace("Заявка", "Приложение к Заявке");
                        string templatePathObj = CurrentDir + "Новая папка\\" + comboBox1.Text + "\\" + strvr + ".docx";
                        Word.Document doc1 = AppWord.Documents.Open(templatePathObj);
                        if (doc1.InlineShapes.Count > 0)
                        {
                            Word.InlineShape pic = doc1.InlineShapes[1];
                            pic.Height = 100;
                            pic.Width = 100;
                            pic.Select();
                            doc1.ActiveWindow.Selection.CopyAsPicture();
                            doc.Paragraphs[11].Range.Paste();
                            Clipboard.Clear();
                        }

                        if (doc1.Shapes.Count > 0)
                        {
                            Word.Shape pic = doc1.Shapes[1];
                            pic.Height = 100;
                            pic.Width = 100;
                            pic.Select();
                            doc1.ActiveWindow.Selection.CopyAsPicture();
                            doc.Paragraphs[11].Range.Paste();
                            Clipboard.Clear();
                            doc.Shapes[1].WrapFormat.Type = Word.WdWrapType.wdWrapInline;

                        }
                        doc1.Close();
                    }

                    j++;
                }

                doc.Paragraphs.Add();
                doc.Paragraphs.Add();
                doc.Paragraphs.Add();
                doc.Paragraphs.Add();
                doc.Paragraphs.Add();
                doc.Tables.Add(doc.Paragraphs[doc.Paragraphs.Count].Range,3,2);
                doc.Tables[1].Rows[1].Cells[1].Range.Paragraphs[1].Range.Bold = -1;
                doc.Tables[1].Rows[1].Cells[1].Range.Paragraphs[1].LineUnitAfter = 0;
                doc.Tables[1].Rows[1].Cells[1].Range.Text = "Начальник Службы";
                doc.Tables[1].Rows[1].Cells[1].Range.Paragraphs.Add();
                
                
                doc.Tables[1].Rows[1].Cells[1].Width = 270;
                doc.Tables[1].Rows[1].Cells[2].Width = 230;
                doc.Tables[1].Rows[2].Cells[1].Width = 270;
                doc.Tables[1].Rows[2].Cells[2].Width = 230;
                doc.Tables[1].Rows[3].Cells[1].Width = 270;
                doc.Tables[1].Rows[3].Cells[2].Width = 230;
                doc.Tables[1].Rows[1].Cells[1].Range.Paragraphs[2].Range.Text = "информационно - управляющих систем" ;
                doc.Tables[1].Rows[1].Cells[1].Range.Paragraphs.Add();
                doc.Tables[1].Rows[1].Cells[1].Range.Paragraphs[3].Range.Text = " ";
                doc.Tables[1].Rows[1].Cells[2].Range.Paragraphs.Add();
                doc.Tables[1].Rows[1].Cells[2].Range.Paragraphs[2].Range.Bold = -1;
               doc.Tables[1].Rows[1].Cells[2].Range.Paragraphs[2].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                doc.Tables[1].Rows[1].Cells[2].Range.Paragraphs[2].Range.Text = "Д.Р.Юсупов";
               doc.Tables[1].Rows[2].Cells[1].Range.Paragraphs[1].Range.Bold = -1;
                doc.Tables[1].Rows[2].Cells[1].Range.Text = "Главный инженер - заместитель";
                doc.Tables[1].Rows[2].Cells[1].Range.Paragraphs.Add();
          
                doc.Tables[1].Rows[2].Cells[1].Range.Paragraphs[2].Range.Text = "генерального директора";
                doc.Tables[1].Rows[2].Cells[1].Range.Paragraphs.Add();
                doc.Tables[1].Rows[2].Cells[1].Range.Paragraphs[3].Range.Text = " ";
                doc.Tables[1].Rows[2].Cells[2].Range.Paragraphs.Add();
                doc.Tables[1].Rows[2].Cells[2].Range.Paragraphs[2].Range.Bold = -1;
                doc.Tables[1].Rows[2].Cells[2].Range.Paragraphs[2].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                doc.Tables[1].Rows[2].Cells[2].Range.Paragraphs[2].Range.Text = "Н.Ф.Низамов";

                doc.Tables[1].Rows[3].Cells[1].Range.Paragraphs[1].Range.Bold = -1;
                doc.Tables[1].Rows[3].Cells[1].Range.Text = "Заместитель генерального директора";
                doc.Tables[1].Rows[3].Cells[1].Range.Paragraphs.Add();

           

                doc.Tables[1].Rows[3].Cells[1].Range.Paragraphs[2].Range.Text = "по общим вопросам";
                doc.Tables[1].Rows[3].Cells[1].Range.Paragraphs.Add();
               // doc.Tables[1].Rows[3].Cells[1].Range.Paragraphs[3].Range.Text = " ";
                doc.Tables[1].Rows[3].Cells[2].Range.Paragraphs.Add();
                doc.Tables[1].Rows[3].Cells[2].Range.Paragraphs[2].Range.Bold = -1;
                doc.Tables[1].Rows[3].Cells[2].Range.Paragraphs[2].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                doc.Tables[1].Rows[3].Cells[2].Range.Paragraphs[2].Range.Text = " С.Ю.Сергеев";





             

           

                CurrentDir = creatdir(CurrentDir);
                doc.SaveAs(CurrentDir + "\\Приложение к Заявке на Сервер Lenovo x3550 M5 80Gb.docx");
              //  AppWord.Quit();
            }
            else
            {
                MessageBox.Show("Нет файла приложение в папке data");
            }



            
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string[] text = new string[dataGridView1.RowCount - 1] ;
            for (int i = 0; i < 16; i++)
            {
                
                for (int z = 0; z < dataGridView1.RowCount-1; ++z)
                {
                    if (dataGridView1.Rows[z].Cells[i].Value == null) { text[z] = ""; }
                    else text[z] = dataGridView1.Rows[z].Cells[i].Value.ToString();
                }

                File.WriteAllLines(("data\\" + i + ".txt"), text);

 
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            string path = null;
            using (var dialog = new FolderBrowserDialog())
                if (dialog.ShowDialog() == DialogResult.OK) { 
            path = dialog.SelectedPath;
            MessageBox.Show(new DirectoryInfo(path).Name);
            string dirName = new DirectoryInfo(path).Name;
            string[] Direct = Directory.GetDirectories(path);
                    if (dirName == "ОНМ СИУС")
                    {
                        int z = dataGridView1.RowCount-1;
                        for (int i = 0; i < Direct.Length; i++)
                        {
                            dataGridView1.Rows.Add();
                        }

                        for (int i = 0; i < Direct.Length; i++)
                        {
                            dataGridView1.Rows[z].Cells[0].Value = new DirectoryInfo(Direct[i]).Name;
                            comboBox1.Items.Add(new DirectoryInfo(Direct[i]).Name);
                            dataGridView1.Rows[z].Cells[1].Value = "Не требуется.";
                            dataGridView1.Rows[z].Cells[2].Value = "Не требуется.";
                            dataGridView1.Rows[z].Cells[6].Value = "Не требуется.";
                            dataGridView1.Rows[z].Cells[13].Value = "Проект закупочной документации(Приложение № 7).";
                            dataGridView1.Rows[z].Cells[15].Value = "1";
                            z++;
                        }

                    }

                    if (dirName == "ПЭН СИУС")
                    {
                        int z = dataGridView1.RowCount - 1;
                        for (int i = 0; i < Direct.Length; i++)
                        {
                            dataGridView1.Rows.Add();
                        }

                        for (int i = 0; i < Direct.Length; i++)
                        {
                            dataGridView1.Rows[z].Cells[0].Value = new DirectoryInfo(Direct[i]).Name;
                            comboBox1.Items.Add(new DirectoryInfo(Direct[i]).Name);
                            dataGridView1.Rows[z].Cells[1].Value = "Не требуется.";
                            dataGridView1.Rows[z].Cells[2].Value = "Не требуется.";
                            dataGridView1.Rows[z].Cells[6].Value = "Не требуется.";
                            dataGridView1.Rows[z].Cells[13].Value = "Проект закупочной документации(Приложение № 7).";
                            dataGridView1.Rows[z].Cells[15].Value = "2";
                            z++;
                        }

                    }






                }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            string CurrentDir = AppDomain.CurrentDomain.BaseDirectory;


                    Word.Application word = new Word.Application(); //создаем COM-объект Word   
            Object missingObj = System.Reflection.Missing.Value;
            Object trueObj = true;
            Object falseObj = false;
            word.Visible = true;
            Word.Document document = new Word.Document();
            Object templatePathObj = CurrentDir + "shablon1.dotx";
            CurrentDir = creatdir(CurrentDir);

            
            try
            {
                document = word.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
            }
            catch (Exception error)
            {
                document.Close(ref falseObj, ref missingObj, ref missingObj);
                word.Quit(ref missingObj, ref missingObj, ref missingObj);
                document = null;
                word = null;
                throw error;
            }

            int index = comboBox1.SelectedIndex;
            object strToFindObj = "name";
            string Str = dataGridView1.Rows[index].Cells[0].Value.ToString();
            Str = Str.Remove(0, 10);
            object replaceStrObj = Str;// dataGridView1.Rows[index].Cells[1].Value.ToString();
            object replaceTypeObj;
            Word.Range wordRange;
            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            for (int i = 1; i <= document.Sections.Count; i++)
            {
                wordRange = document.Sections[i].Range;
                Word.Find wordFindObj = wordRange.Find;
                object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceStrObj, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);
            }

            strToFindObj = "kod";
            replaceStrObj = dataGridView1.Rows[index].Cells[16].Value.ToString();
            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            for (int i = 1; i <= document.Sections.Count; i++)
            {
                wordRange = document.Sections[i].Range;
                Word.Find wordFindObj = wordRange.Find;
                object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceStrObj, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);
            }

            strToFindObj = "prod";
            replaceStrObj = dataGridView1.Rows[index].Cells[4].Value.ToString();
            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            for (int i = 1; i <= document.Sections.Count; i++)
            {
                wordRange = document.Sections[i].Range;
                Word.Find wordFindObj = wordRange.Find;
                object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceStrObj, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);
            }

            strToFindObj = "kol";
            replaceStrObj = dataGridView1.Rows[index].Cells[7].Value.ToString();
            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            for (int i = 1; i <= document.Sections.Count; i++)
            {
                wordRange = document.Sections[i].Range;
                Word.Find wordFindObj = wordRange.Find;
                object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceStrObj, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);
            }

            strToFindObj = "teh";
            Str = dataGridView1.Rows[index].Cells[3].Value.ToString();
            replaceStrObj = Str;// dataGridView1.Rows[index].Cells[4].Value.ToString();
            MessageBox.Show(Str.Length.ToString());
            replaceTypeObj = Word.WdReplace.wdReplaceAll;
            for (int i = 1; i <= document.Sections.Count; i++)
            {
                wordRange = document.Sections[i].Range;
                Word.Find wordFindObj = wordRange.Find;
                object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceStrObj, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };
                wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);
            }





            document.SaveAs(CurrentDir + "\\" + comboBox1.Text + ".docx");
            document.Close();
            word.Quit();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            string CurrentDir = AppDomain.CurrentDomain.BaseDirectory;
            int selectItem = comboBox1.SelectedIndex;
            //string CurrentDir = AppDomain.CurrentDomain.BaseDirectory;
            if (!Directory.Exists("Заявки"))
            {
                Directory.CreateDirectory("Заявки");
            }

            if (!Directory.Exists("Заявки\\ОНМ СИУС"))// dataGridView1.Rows[0].Cells[0].Value.ToString()))
            {
                Directory.CreateDirectory("Заявки\\ОНМ СИУС"); // dataGridView1.Rows[0].Cells[0].Value.ToString());
            }

            if (!Directory.Exists("Заявки\\ПЭН СИУС"))// dataGridView1.Rows[0].Cells[0].Value.ToString()))
            {
                Directory.CreateDirectory("Заявки\\ПЭН СИУС"); // dataGridView1.Rows[0].Cells[0].Value.ToString());
            }


            string path;
            if (dataGridView1.Rows[selectItem].Cells[15].Value.ToString() == "1")
            {
                if (!Directory.Exists("Заявки\\ОНМ СИУС\\" + comboBox1.Text))// dataGridView1.Rows[0].Cells[0].Value.ToString()))
                {
                    Directory.CreateDirectory("Заявки\\ОНМ СИУС\\" + comboBox1.Text); // dataGridView1.Rows[0].Cells[0].Value.ToString());
                }
                path = "Заявки\\ОНМ СИУС\\" + comboBox1.Text;
                CurrentDir = CurrentDir + path;
            }
            else {
                if (!Directory.Exists("Заявки\\ПЭН СИУС\\" + comboBox1.Text))// dataGridView1.Rows[0].Cells[0].Value.ToString()))
                {
                    Directory.CreateDirectory("Заявки\\ПЭН СИУС\\" + comboBox1.Text); // dataGridView1.Rows[0].Cells[0].Value.ToString());
                }
                path = "Заявки\\ПЭН СИУС\\" + comboBox1.Text;
                CurrentDir = CurrentDir + path;
            }



            if (!Directory.Exists(path + "\\Приложение"))// dataGridView1.Rows[0].Cells[0].Value.ToString()))
            {
                Directory.CreateDirectory(path+ "\\Приложение"); // dataGridView1.Rows[0].Cells[0].Value.ToString());
            }


            Word.Application AppWord = new Word.Application();
            Word.Document doc = new Word.Document();

            if (!File.Exists("data\\prilojenie.txt")) return;
            string[] stroka = File.ReadAllLines("data\\prilojenie.txt");
            for (int i = 0; i < 14; i++)
            {
                i++;
                if (!Directory.Exists(path + "\\Приложение\\" + i.ToString()))Directory.CreateDirectory(path + "\\Приложение\\" + i.ToString());
                
                    doc = AppWord.Documents.Add();
                    doc.Paragraphs.Add();
                doc.Paragraphs[1].Range.Font.Size = 14;
                    doc.Paragraphs[1].Range.Text = stroka[i];
                doc.Paragraphs.Add();
                doc.Paragraphs[2].Range.Font.Size = 14;
                if (dataGridView1[i, comboBox1.SelectedIndex].Value == null) dataGridView1[i, comboBox1.SelectedIndex].Value = " ";
                doc.Paragraphs[2].Range.Text = dataGridView1[i, comboBox1.SelectedIndex].Value.ToString();
                doc.SaveAs(CurrentDir + "\\Приложение\\"+i.ToString()+"\\п."+i.ToString()+" приложения к заявке.docx");
                doc.Close();
                i--;
            }
            AppWord.Quit();
            MessageBox.Show("Выполнено");
        }

        public string creatdir (string CurrentDir)
        {
            //string CurrentDir = AppDomain.CurrentDomain.BaseDirectory;
            int selectItem = comboBox1.SelectedIndex;
            if (!Directory.Exists("Заявки"))
            {
                Directory.CreateDirectory("Заявки");
            }

            if (!Directory.Exists("Заявки\\ОНМ СИУС"))// dataGridView1.Rows[0].Cells[0].Value.ToString()))
            {
                Directory.CreateDirectory("Заявки\\ОНМ СИУС"); // dataGridView1.Rows[0].Cells[0].Value.ToString());
            }

            if (!Directory.Exists("Заявки\\ПЭН СИУС"))// dataGridView1.Rows[0].Cells[0].Value.ToString()))
            {
                Directory.CreateDirectory("Заявки\\ПЭН СИУС"); // dataGridView1.Rows[0].Cells[0].Value.ToString());
            }


            string path;
            if (dataGridView1.Rows[selectItem].Cells[15].Value.ToString() == "1")
            {
                if (!Directory.Exists("Заявки\\ОНМ СИУС\\" + comboBox1.Text))// dataGridView1.Rows[0].Cells[0].Value.ToString()))
                {
                    Directory.CreateDirectory("Заявки\\ОНМ СИУС\\" + comboBox1.Text); // dataGridView1.Rows[0].Cells[0].Value.ToString());
                }
                path = "Заявки\\ОНМ СИУС\\" + comboBox1.Text;
                return CurrentDir = CurrentDir + path;
            }
            else {
                if (!Directory.Exists("Заявки\\ПЭН СИУС\\" + comboBox1.Text))// dataGridView1.Rows[0].Cells[0].Value.ToString()))
                {
                    Directory.CreateDirectory("Заявки\\ПЭН СИУС\\" + comboBox1.Text); // dataGridView1.Rows[0].Cells[0].Value.ToString());
                }
                path = "Заявки\\ПЭН СИУС\\" + comboBox1.Text;
                return CurrentDir = CurrentDir + path;
            }

        }

        private void button13_Click(object sender, EventArgs e)
        {
            string[] text = File.ReadAllLines("data\\0.txt");
            comboBox1.Items.AddRange(text);
            comboBox1.SelectedIndex = 0;
            dataGridView1.AutoGenerateColumns = true;

          //  string[] text1 = File.ReadAllLines("data\\16.txt");

           
            for (int i = 0; i < 18; ++i)
            {
                //if (!File.Exists("data\\" + i + ".txt"))
                //{
                //    MessageBox.Show("Нет файла "+i+".txt в папк data");
                //    File.Create("data\\" + i + ".txt");
                //    return;
                //}
                text = File.ReadAllLines("data\\" + i.ToString() + ".txt");
                for (int z = 0; z < text.Length-1; ++z)
                {
                    dataGridView1.Rows[z].Cells[i].Value = text[z];
          //          dataGridView1.Rows[z].Cells[16].Value = text1[z];
                }
            }
        }
    }
}

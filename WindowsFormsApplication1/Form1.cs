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
            form2 = new Form2();
        }
        Form2 form2;

        private void Form1_Load(object sender, EventArgs e)
        {



           // dataGridView1.VirtualMode = true;
            string[] text = File.ReadAllLines("dir.txt");
           comboBox1.Items.AddRange(text);
           comboBox1.SelectedIndex = 0;
            text = File.ReadAllLines("text.txt");
/*0*/         dataGridView1.Columns.Add("Cat", "0. Каталог");
/*1*/       dataGridView1.Columns.Add("Name", "1. Наименование");
/*2*/       dataGridView1.Columns.Add("ves", "2. Ед.");
/*3*/       dataGridView1.Columns.Add("Kol", "3. шт");
/*4*/       dataGridView1.Columns.Add("Teh", "4. Тех. Описание");
/*5*/       dataGridView1.Columns.Add("Product", "5. Производитель");
/*6*/       dataGridView1.Columns.Add("Kod", "6. Коды");




         //   DataGridViewImageColumn iconColumn = new DataGridViewImageColumn();


       //     dataGridView1.Columns.Add(iconColumn);
            /*7*/
            dataGridView1.Columns.Add("TR", "7. ТР.ТС.");
            dataGridView1.Columns.Add("Instr", "8. Инстр.");
            dataGridView1.Columns.Add("Zarub", "9. Заруб.");
            dataGridView1.Columns.Add("Vubor", "10. Обоснование");
            dataGridView1.Columns.Add("Rus", "11. Рус");
            dataGridView1.Columns.Add("sved", "12. Сведения");
            dataGridView1.Columns.Add("analiz", "13. Ср. анализ");
            dataGridView1.Columns.Add("sogl", "14. Согласование");
            dataGridView1.Columns[1].Width= 150;
            dataGridView1.Columns[2].Width = 25;
            dataGridView1.Columns[3].Width = 25;
            int matrix;
           matrix = text.Length;
            int sum = 0;
            for (int i = 0; i < matrix; ++i)
            {
                dataGridView1.Rows.Add();
                dataGridView1.Rows[i].Cells[2].Value = "шт";
            }

          //dir 1
                for (int i = 0; i < matrix; i++)
                {
                    try
                    {
                        dataGridView1.Rows[i].Cells[1].Value = text[sum];
                        sum++;
                    }
                    catch (Exception e1)
                    {
                        MessageBox.Show(e1.Message);
                    }
                }



            text = File.ReadAllLines("rus.txt");
            matrix = text.Length;
            sum = 0;
            for (int i = 0; i < matrix; i++)
            {
                try
                {
                    dataGridView1.Rows[i].Cells[11].Value = text[sum];
                    sum++;
                }
                catch (Exception e1)
                {
                    MessageBox.Show(e1.Message);
                }
            }

            text = File.ReadAllLines("odobren.txt");
            matrix = text.Length;
            sum = 0;
            for (int i = 0; i < matrix; i++)
            {
                try
                {
                    dataGridView1.Rows[i].Cells[14].Value = text[sum];
                    sum++;
                }
                catch (Exception e1)
                {
                    MessageBox.Show(e1.Message);
                }
            }



            text = File.ReadAllLines("sranaliz.txt");
            matrix = text.Length;
            sum = 0;
            for (int i = 0; i < matrix; i++)
            {
                try
                {
                    dataGridView1.Rows[i].Cells[13].Value = text[sum];
                    sum++;
                }
                catch (Exception e1)
                {
                    MessageBox.Show(e1.Message);
                }
            }




            text = File.ReadAllLines("svedenia.txt");
            matrix = text.Length;
            sum = 0;
            for (int i = 0; i < matrix; i++)
            {
                try
                {
                    dataGridView1.Rows[i].Cells[12].Value = text[sum];
                    sum++;
                }
                catch (Exception e1)
                {
                    MessageBox.Show(e1.Message);
                }
            }


            text = File.ReadAllLines("obosn.txt");
            matrix = text.Length;
            sum = 0;
            for (int i = 0; i < matrix; i++)
            {
                try
                {
                    dataGridView1.Rows[i].Cells[10].Value = text[sum];
                    sum++;
                }
                catch (Exception e1)
                {
                    MessageBox.Show(e1.Message);
                }
            }

            text = File.ReadAllLines("zarub.txt");
            matrix = text.Length;
            sum = 0;
            for (int i = 0; i < matrix; i++)
            {
                try
                {
                    dataGridView1.Rows[i].Cells[9].Value = text[sum];
                    sum++;
                }
                catch (Exception e1)
                {
                    MessageBox.Show(e1.Message);
                }
            }
            
            text = File.ReadAllLines("instr.txt");
            matrix = text.Length;
            sum = 0;
            for (int i = 0; i < matrix; i++)
            {
                try
                {
                    dataGridView1.Rows[i].Cells[8].Value = text[sum];
                    sum++;
                }
                catch (Exception e1)
                {
                    MessageBox.Show(e1.Message);
                }
            }



            text = File.ReadAllLines("trts.txt");
            matrix = text.Length;
            sum = 0;
            for (int i = 0; i < matrix; i++)
            {
                try
                {
                    dataGridView1.Rows[i].Cells[7].Value = text[sum];
                    sum++;
                }
                catch (Exception e1)
                {
                    MessageBox.Show(e1.Message);
                }
            }

            text = File.ReadAllLines("teh.txt");
            matrix = text.Length;
            sum = 0;
                for (int i = 0; i < matrix; i++)
                {
                    try
                    {
                        dataGridView1.Rows[i].Cells[4].Value = text[sum];
                        sum++;
                    }
                    catch (Exception e1)
                    {
                        MessageBox.Show(e1.Message);
                    }
                }
         
            
            text = File.ReadAllLines("product.txt");            
            matrix = text.Length;
            sum = 0;
           //4
                for (int i = 0; i < matrix; i++)
                {
                    try
                    {
                        dataGridView1.Rows[i].Cells[5].Value = text[sum];
                        sum++;
                    }
                    catch (Exception e1)
                    {
                        MessageBox.Show(e1.Message);
                    }
                }



            text = File.ReadAllLines("kod.txt");

            matrix = text.Length;
            sum = 0;

            for (int i = 0; i < matrix; i++)
            {
                try
                {
                    dataGridView1.Rows[i].Cells[6].Value = text[sum];
                    sum++;
                }
                catch (Exception e1)
                {
                    MessageBox.Show(e1.Message);
                }
            }




            /////////////////////////BITMAP???????????????????

            //Word.Application word = new Word.Application();
            //string strvr = comboBox1.Text;
            //strvr = strvr.Replace("Заявка", "Приложение к Заявке");
            //string CurrentDir = AppDomain.CurrentDomain.BaseDirectory;
            //string templatePathObj = CurrentDir + "Новая папка\\" + comboBox1.Text + "\\" + strvr + ".docx";


            //Bitmap img;
            //Word.Document doc = word.Documents.Open(templatePathObj);
            //if (doc.InlineShapes.Count > 0)
            //{
            //    Word.InlineShape pic = doc.InlineShapes[1];
            //    pic.Height = 100;
            //    pic.Width = 100;
            //    pic.Select();
            //    doc.ActiveWindow.Selection.CopyAsPicture();
            //  //  iconColumn.Image = new Bitmap(Clipboard.GetImage());
            //   dataGridView1.Rows[0].Cells[6].Value = new Bitmap(Clipboard.GetImage());
            //    Clipboard.Clear();
            //}

            //doc.Close();
            //word.Quit();






            ////////////////////END BITMAP///////////////

            //           MemoryStream ms = new MemoryStream((byte[])dataGridView1.Rows[0].Cells[6].Value);
            //           pictureBox1.Image = Image.FromStream(ms);

            // Image imge = Image.FromStream(new MemoryStream((byte[])dataGridView1.Rows[0].Cells[6].Value));
            //  pictureBox1.Image = Image.FromStream(new MemoryStream((byte[])dataGridView1.Rows[0].Cells[6].Value));



            text = File.ReadAllLines("kol.txt");            
            matrix = text.Length;
            sum = 0;
            for (int j=3;j < 4;j++)
            {
                for (int i = 0; i< matrix; i++)
                {
                    try
                    {
                        dataGridView1.Rows[i].Cells[j].Value = text[sum];
                        sum++;
                    }
                    catch (Exception e1)
                    {
                        MessageBox.Show(e1.Message);
                    }
                }

            }


            text = File.ReadAllLines("dir.txt");
            matrix = text.Length;
            sum = 0;
            for (int j = 0; j < 1; j++)
            {
                for (int i = 0; i < matrix; i++)
                {
                    try
                    {
                        dataGridView1.Rows[i].Cells[j].Value = text[sum];
                        sum++;
                    }
                    catch (Exception e1)
                    {
                        MessageBox.Show(e1.Message);
                    }
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
            object replaceStrObj = dataGridView1.Rows[index].Cells[1].Value.ToString();
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
            replaceStrObj = dataGridView1.Rows[index].Cells[1].Value.ToString();
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
            Form2 newform = new Form2();
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
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using COWTranslation;
using System.IO;
namespace Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //TranslationToWord.UpadteTextAfterBookmarke("a11", "محمد هشام خالد المختار", "c:\\1.docx");
            //TranslationToWord.UpadteTextAfterBookmarke("a22", "محمد هشام خالد المختار", "c:\\1.docx");
            //TranslationToWord.UpadteTextAfterBookmarke("box", "محمد هشام خالد المختار", "c:\\1.docx");
            OpenFileDialog open = new OpenFileDialog();
            open.ShowDialog(this);
            Image img = Image.FromFile(open.FileName);
            System.IO.MemoryStream strm = new System.IO.MemoryStream();
            img.Save(strm, System.Drawing.Imaging.ImageFormat.Jpeg);
            TranslationToWord Trw = new TranslationToWord("c:\\1.docx", "d:\\1-3.docx", true);
            DateTime st = DateTime.Now;
            for (int i = 0; i <= 50; i++)
            {
             //   تحميل الصفحات داخل مصفوة الكلاس                
                MemoryStream mem = new MemoryStream();                
                img.Save(mem,System.Drawing.Imaging.ImageFormat.Jpeg );
                Dictionary<string,byte[]> picts = new Dictionary<string,byte[]>();
                picts.Add("picture",mem.ToArray());
                Dictionary<string,string> strs = new Dictionary<string,string>();
                strs.Add("firstname","Hesham " + i.ToString());
                Trw.AddPage(i,picts,strs);
            }
            Trw.DoIt();
            TimeSpan  en = DateTime.Now.Subtract(st);
            MessageBox.Show(en.TotalSeconds.ToString());
        }
    }
}

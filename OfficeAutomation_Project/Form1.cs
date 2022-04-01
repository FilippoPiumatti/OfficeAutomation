using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using OfficeAutomation_Project.Controller;

namespace OfficeAutomation_Project
{
    public partial class Form1 : Form
    {
        Microsoft.Office.Interop.Word.Application MyWord; // DICHIARAZIONE PER POI ISTANZIARE WORD
        Microsoft.Office.Interop.Word.Document MyDoc; // GRAZIE A QUESTA DICHIARAZIONE, POSSIAMO ANDARE A LAVORARE SUL DOCUMENTO WORD
        Microsoft.Office.Interop.Word.Range MyRange;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //CARICO I FONT NELLA COMBO

            foreach (FontFamily family in FontFamily.Families)
            {
                cmbFont.Items.Add(family.Name);
            }
            cmbFont.SelectedIndex = 0;

            // CARICO COMBO SIZE

            for (int i = 10; i <= 80 ; i += 2)
            {
                cmbSize.Items.Add(i);
            }
            cmbSize.SelectedIndex = 0;

            //CARICO COMBO ALLIENAMENTI
            
            AllineamentoController alCtrl = new AllineamentoController();

            cmbAllineamento.DataSource = alCtrl.GetAllineamenti();
            cmbAllineamento.DisplayMember = "StrAllineamento";
            cmbAllineamento.ValueMember = "ValAllineamento";
            cmbAllineamento.SelectedIndex = 0;

            //CARICO COMBO SOTTOLINEATURA

            SottolineaturaController sotCtrl = new SottolineaturaController();

            cmbSottolineato.DataSource = sotCtrl.GetSottolineaturas();
            cmbSottolineato.DisplayMember = "StrSottolineatura";
            cmbSottolineato.ValueMember = "ValSottolineatura";
            cmbSottolineato.SelectedIndex = 0;

            //CARICO COMBO COLORI

            ColoreController ColoreCtrl = new ColoreController();

            cmbColore.DataSource = ColoreCtrl.GetColores();
            cmbColore.DisplayMember = "StrColore";
            cmbColore.ValueMember = "ValColore";
            cmbColore.SelectedIndex = 0;

            //CARICO DIMENSIONI TABELLA

            for (int i = 2; i <= 10; i++)
            {
                cmbRighe.Items.Add(i);
                cmbColonne.Items.Add(i);
            }

            cmbRighe.SelectedIndex = -1;
            cmbColonne.SelectedIndex = -1;

            //ALL AVVIO DEL PROGRAMMA, DISATTIVO I SEGUENTI BOTTONI: 

            btnSalva.Enabled = false; 
            btnInserisciTabella.Enabled = false;
            btnInserisciTesto.Enabled = false;
        }

        private void btnCrea_Click(object sender, EventArgs e)
        {
            // CREO ISTANZA DELL APPLICAZIONE WORD

            MyWord = new Microsoft.Office.Interop.Word.Application();
            MyWord.Visible = true;

            // AGGIUNGIAMO L ISTANZA DI UN DOCUMENTO

            MyDoc = new Document();
            MyDoc = MyWord.Documents.Add();

            btnSalva.Enabled = true;
            btnInserisciTabella.Enabled = true;
            btnInserisciTesto.Enabled = true;
        }

        //VADO A GESTIRE LA TEXTBOX DOVE COMPARIRA' IL TESTO DEL DOCUMENTO WORD
        private void txtTesto_Enter(object sender, EventArgs e)
        {
            if (txtTesto.ForeColor == Color.Gray)
            {
                txtTesto.ForeColor = Color.Black;
                txtTesto.Text = "";
            }
            
        }

        private void txtTesto_Leave(object sender, EventArgs e)
        {
            btnInserisciTesto.Enabled = true;
            if (txtTesto.Text.Trim() == "")
            {
                txtTesto.Text = "Inserire il testo da scrivere sul file Word...";
                txtTesto.ForeColor = Color.Gray;
                btnInserisciTesto.Enabled = false;
            }
        }

        private void btnInserisciTesto_Click(object sender, EventArgs e)
        {
            if (txtTesto.ForeColor == Color.Silver || txtTesto.Text.Trim() == "Inserire il testo da scrivere sul file Word...")
            {
                MessageBox.Show("Inserire il testo da scrivere sul documento word...");
            }
            else
            {
                //FORMATTIAMO IL TESTO CHE C'E' NELLA TEXTBOX 

                ImpostaRange();

                MyRange.Text = txtTesto + "\n";
                MyRange.Font.Name = cmbFont.Text;
                MyRange.Font.Size = Convert.ToInt32(cmbSize.Text);

                if (chkBold.Checked)
                {
                    MyRange.Bold = 1;
                }
                else
                {
                    MyRange.Bold = 0;
                }
                if (chkItalic.Checked)
                {
                    MyRange.Italic = 1;
                }
                else
                {
                    MyRange.Italic = 0;
                }

                WdUnderline wUline;
                Enum.TryParse(cmbSottolineato.SelectedValue.ToString(), out wUline);
                MyRange.Font.Underline = wUline;

                WdParagraphAlignment wPar;
                Enum.TryParse(cmbAllineamento.SelectedValue.ToString(), out wPar);
                MyRange.ParagraphFormat.Alignment = wPar;

                WdColor wColor;
                Enum.TryParse(cmbColore.SelectedValue.ToString(), out wColor);
                MyRange.Font.Color = wColor;
            }
        }

        //IMPOSTA IL PUNTO DEL DODUMENTO WORD IN CUI VERRA' INSERITO IL TESTO
        public void ImpostaRange()
        {
            object Start;
            object End;

            Start = MyDoc.Sentences[MyDoc.Sentences.Count].End - 1;
            End = MyDoc.Sentences[MyDoc.Sentences.Count].End;

            MyRange = MyDoc.Range(ref Start, ref End); // IL TESTO VERRA' INSERITO DA QUI
        }

        //GESTIAMO LA FINESTRA DEI FONT ATTRAVERSO LA DIALOG RESULT

        private void btnFontDialog_Click(object sender, EventArgs e)
        {
            DialogResult DlgFont = fontDialog1.ShowDialog();

            if (DlgFont == DialogResult.OK)
            {
                cmbFont.Text = fontDialog1.Font.Name;
                cmbSize.Text = Math.Round(Convert.ToSingle(fontDialog1.Font.Size)).ToString();
                chkItalic.Checked = fontDialog1.Font.Italic; // FONT DIALOG HA DEI BOOL CHE SE SON CHECCATI LI METTIAMO NEI CHECKBOX
                chkBold.Checked = fontDialog1.Font.Bold;

                if (fontDialog1.Font.Underline)// STESSA COSA DI PRIMA CON I CHECKBOX MA GIOCHIAMO CON GLI INDICI DELLA COMBO
                {
                    cmbSottolineato.SelectedIndex = 1;
                }
                else
                {
                    cmbSottolineato.SelectedIndex = 0;
                }
            }
            else if(DlgFont == DialogResult.Cancel)
            {
                MessageBox.Show("Azione Annullata!" , "Attenzione!", MessageBoxButtons.OK,MessageBoxIcon.Warning);
            }
        }
    }
}

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
    #region Initialize
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
        #endregion
        #region Events
        //CREO IL DOCUMENTO DA ZERO CON I VARI PARAMETRI
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

                MyRange.Text = txtTesto.Text + "\n";
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
        private void btnSalva_Click(object sender, EventArgs e)
        {
            MyDoc.SaveAs2(System.Windows.Forms.Application.StartupPath + @"\prova.docx");// salvo il doc in bin/debug
            MyDoc.Close();//chiudo il doc
            MyWord.Quit();//killo le istanze di word
        }

        // GESTIAMO L' INSERIMENTO DI UNA TABELLA WORD
        private void btnInserisciTabella_Click(object sender, EventArgs e)
        {
            int r, c, i, j;
            Table MyTable;

            r = Convert.ToInt32(cmbRighe.Text);
            c = Convert.ToInt32(cmbColonne.Text);

            ImpostaRange(); // aggiunge in coda al documento

            MyTable = MyDoc.Tables.Add(MyRange, r, c);
            MyTable.Borders.Enable = 1; //abilitato

            for  (i = 1; i <= r; i++) // riempo la tabella
            {
                for (j = 1; j <= c; j++)
                {
                    MyTable.Cell(i, j).Range.Text = "R" + i.ToString() + "-C" + j.ToString();//non centra con myrange
                    MyTable.Cell(i, j).Width = 100;
                    MyTable.Cell(i, j).Height = 100;
                    MyTable.Cell(i, j).Range.ParagraphFormat.SpaceAfter = 30; // come il padding su css
                    MyTable.Cell(i, j).Range.ParagraphFormat.SpaceBefore = 30; //stessa cosa

                    if ((i + j) % 2 == 0) //se è pari
                    {
                        MyTable.Cell(i, j).Range.Shading.BackgroundPatternColorIndex = WdColorIndex.wdYellow;
                        MyTable.Cell(i, j).Range.Font.ColorIndex = WdColorIndex.wdBlue;
                    }
                    else
                    {
                        MyTable.Cell(i, j).Range.Shading.BackgroundPatternColorIndex = WdColorIndex.wdBrightGreen; //back color del doc
                        MyTable.Cell(i, j).Range.Font.ColorIndex = WdColorIndex.wdRed; // fore color del doc
                    }

                    MyTable.Cell(i, j).Range.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphCenter; // allinneameoto orizziontale
                    MyTable.Cell(i, j).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;// alinemaneto verticale
                    MyTable.Cell(i, j).Range.Bold = 1; //grassetto abilitato
                }
            }
            ImpostaRange();
            MyRange.Text = "\n";
        }

        // GESTIAMO TUTTI GLI EVENTI DI UN CLASSICO EDITOR WORD ATTRRAVERSO DEI BOTTONO (Es: CERCA E SOSTITUISCI)
        private void btnSelezionaTesto_Click(object sender, EventArgs e)
        {
            object Start, end;

            Start = (object)(Interaction.InputBox("Inserisci l' indice Iniziale: "));
            end = (object)(Interaction.InputBox("Inserisci l' indice Finale: "));

            try
            {
                MyRange = MyDoc.Range(ref Start, ref end);
                MyRange.Select();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Selezione non valida: " + ex.Message," Attenzione " , MessageBoxButtons.OK,MessageBoxIcon.Warning);
            }

        }
        //VEDIAO LA SELEZIONE DEI CARATTERI
        private void btnVediSelezione_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Testo Selezionato : " + MyWord.Selection.Text + " - indice inizio: " + MyWord.Selection.Start.ToString() + " indice fine: " + MyWord.Selection.End.ToString());
        }
        //CERCA E SOSTITUISCI
        private void btnSearchAndFind_Click(object sender, EventArgs e)
        {
            object Start, End;
            object TextToFind = Interaction.InputBox("Inserisci il Testo da ricercare");

            MyWord.Selection.Find.ClearFormatting();

            if (MyWord.Selection.Find.Execute(ref TextToFind))
            {
                Start = MyWord.Selection.Start;
                End = MyWord.Selection.End;

                MyRange = MyDoc.Range(ref Start, ref End);
                MyRange.Text = Interaction.InputBox("Inserisci il nuovo Testo: ");

            }
            else
            {
                MessageBox.Show("Testo non Trovato: ", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning); ;
            }
        }
        //SOSTITUISCI TUTTO
        private void btnSubAll_Click(object sender, EventArgs e)
        {
            object Start, End;
            object MissingType = Type.Missing; //omissione di parametro
            object TextToFind = Interaction.InputBox("Inserisci il Testo da ricercare");
            string NewText = Interaction.InputBox("Inserisci il Nuvo testo da Aggiungere: ");
            int init;

            MyWord.Selection.ClearFormatting();

            while (MyWord.Selection.Find.Execute(ref TextToFind,ref MissingType,true))//il true cerca per parola
            {
                Start = MyWord.Selection.Start;
                End = MyWord.Selection.End;

                MyRange = MyDoc.Range(ref Start, ref End);
                MyRange.Text = NewText;

                init = Convert.ToInt32(Start) + NewText.Length + 1; //sposto il cursore a capo di uno cosi a evitare loop parola uguale
                Start = (object)init;
                End = MyDoc.Content.End;

                MyRange = MyDoc.Range(ref Start, ref End);
                MyRange.Select();
            }

            MyRange = MyDoc.Range(0,0);//mi sposto all inizio
            MyRange.Select();
        }
        //2.0 SOSTITUISCI TUTTO
        private void btnSubAll2_Click(object sender, EventArgs e)
        {
            object Start, End;
            object MissingType = Type.Missing; //omissione di parametro
            object TextToFind = Interaction.InputBox("Inserisci il Testo da ricercare");
            object NewText = Interaction.InputBox("Inserisci il Nuvo testo da Aggiungere: ");

            MyWord.Selection.ClearFormatting();
            MyWord.Selection.Find.Execute(

                //PARAMETRI

                ref TextToFind, MissingType,
                true, MissingType, MissingType,
                MissingType, MissingType,
                MissingType, MissingType,
                ref NewText, WdReplace.wdReplaceAll
                );
        }
        //ESPORTAZIONE IN PDF
        private void BtnCreatePdf_Click(object sender, EventArgs e)
        {
            try
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK) // se clicckiamo ok:
                {
                    Microsoft.Office.Interop.Word.Application WordApp = new Microsoft.Office.Interop.Word.Application();
                    Document WordDoc = new Document();
                   
                    WordApp.Visible = false;
                    WordDoc = WordApp.Documents.Open(openFileDialog1.FileName); //uso la dialog per sfogliare i file

                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        WordDoc.ExportAsFixedFormat(saveFileDialog1.FileName, WdExportFormat.wdExportFormatPDF, true);
                        //esportiamo in pdf, il true mi apre direttamente il pdf
                    }

                    WordDoc.Close();
                    WordApp.Quit();

                    MessageBox.Show("Creazione File PDF Terminata con Successo...","Success",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Creazione PDF Fallita : " + ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }
        //GESTIONE DELLA STAMPA SU VERA STAMPANTE
        private void btnStampa_Click(object sender, EventArgs e)
        {
            try
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    Microsoft.Office.Interop.Word.Application WordApp = new Microsoft.Office.Interop.Word.Application();
                    Document WordDoc = new Document();

                    //N.B LE FINESTRE DI DIALOGO FUNZIONANO TUTTE ALLO STESSO MODO!

                    WordApp.Visible = false;
                    WordDoc = WordApp.Documents.Open(openFileDialog1.FileName); //passiamo al doc il nome del file che vogliamo aprire

                    if (printDialog1.ShowDialog() == DialogResult.OK) 
                    {
                        WordDoc.PrintOut(); //va a stampare sulla stampante correttamente selezionata
                        MessageBox.Show("Documento Stampato Correttamente","Attenzione",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                    }
                    else
                    {
                        MessageBox.Show("Operazione Annullata", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    WordDoc.Close();
                    WordApp.Quit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore: " + ex.Message, "FATAL", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion
        // GESTISCO UN FORM DI REGISTRAZIONE ATTRAVERSO I SEGNALIBRI DI WORD
        private void btnSubmit_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application WordVerbale = new Microsoft.Office.Interop.Word.Application();
            Document DocVerbale = new Document();
            object TemplateName = System.Windows.Forms.Application.StartupPath + @"\reg.docx";

            DocVerbale = WordVerbale.Documents.Open(TemplateName);
            try
            {
                WordVerbale.Visible = true;

                //segnalibro cog
                if (DocVerbale.Bookmarks.Exists("cog")) // controllo se il segnalibro di nome cog esiste
                {
                    object wBookmarks = "cog";//puntiamo il bookmark cog

                    DocVerbale.Bookmarks.get_Item(ref wBookmarks).Range.Text = txtCognome.Text.Trim();// lo passiamo alla getitem per referenza cosi da posizionarci li successivamente
                }
                else
                {
                    throw new Exception("Eccezione Generata Dalla Mancanza Del Segnalibro Cognome");
                }

                //segnalibro nome
                if (DocVerbale.Bookmarks.Exists("nome")) // controllo se il segnalibro di nome cog esiste
                {
                    object wBookmarks = "nome";//puntiamo il bookmark cog

                    DocVerbale.Bookmarks.get_Item(ref wBookmarks).Range.Text = txtNome.Text.Trim();// lo passiamo alla getitem per referenza cosi da posizionarci li successivamente
                }
                else
                {
                    throw new Exception("Eccezione Generata Dalla Mancanza Del Segnalibro Nome");
                }

                //segnalibro dataN
                if (DocVerbale.Bookmarks.Exists("dataN")) // controllo se il segnalibro di nome cog esiste
                {
                    object wBookmarks = "dataN";//puntiamo il bookmark cog

                    DocVerbale.Bookmarks.get_Item(ref wBookmarks).Range.Text = dateTimePicker1.Value.ToShortDateString();// lo passiamo alla getitem per referenza cosi da posizionarci li successivamente
                }
                else
                {
                    throw new Exception("Eccezione Generata Dalla Mancanza Del Segnalibro dataN");
                }

                //segnalibro mail
                if (DocVerbale.Bookmarks.Exists("mail")) // controllo se il segnalibro di nome cog esiste
                {
                    object wBookmarks = "mail";//puntiamo il bookmark cog

                    DocVerbale.Bookmarks.get_Item(ref wBookmarks).Range.Text = txtEmail.Text.Trim();// lo passiamo alla getitem per referenza cosi da posizionarci li successivamente
                }
                else
                {
                    throw new Exception("Eccezione Generata Dalla Mancanza Del Segnalibro e-mail");
                }

                //segnalibro pwd
                if (DocVerbale.Bookmarks.Exists("pwd")) // controllo se il segnalibro di nome cog esiste
                {
                    object wBookmarks = "pwd";//puntiamo il bookmark cog

                    DocVerbale.Bookmarks.get_Item(ref wBookmarks).Range.Text = txtPassword.Text.Trim();// lo passiamo alla getitem per referenza cosi da posizionarci li successivamente
                }
                else
                {
                    throw new Exception("Eccezione Generata Dalla Mancanza Del Segnalibro pwd");
                }

                //segnalibro dataOggi
                if (DocVerbale.Bookmarks.Exists("dataOggi")) // controllo se il segnalibro di nome cog esiste
                {
                    object wBookmarks = "dataOggi";//puntiamo il bookmark cog

                    DocVerbale.Bookmarks.get_Item(ref wBookmarks).Range.Text = DateTime.Today.ToShortDateString();// lo passiamo alla getitem per referenza cosi da posizionarci li successivamente
                }
                else
                {
                    throw new Exception("Eccezione Generata Dalla Mancanza Del Segnalibro dataOggi");
                }

                //segnalibro dataCitta
                if (DocVerbale.Bookmarks.Exists("dataCitta")) // controllo se il segnalibro di nome cog esiste
                {
                    object wBookmarks = "dataCitta";//puntiamo il bookmark cog

                    DocVerbale.Bookmarks.get_Item(ref wBookmarks).Range.Text = DateTime.Today.ToShortDateString();// lo passiamo alla getitem per referenza cosi da posizionarci li successivamente
                }
                else
                {
                    throw new Exception("Eccezione Generata Dalla Mancanza Del Segnalibro dataCitta");
                }

                DocVerbale.SaveAs2(System.Windows.Forms.Application.StartupPath + @"\reg_" + txtCognome.Text.Trim() +  "_" + txtNome.Text.Trim() + ".docx");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "FATAL" ,MessageBoxButtons.OK,MessageBoxIcon.Error);
            }

            /*DocVerbale.Close();
            WordVerbale.Quit();*/
        }
    }
}

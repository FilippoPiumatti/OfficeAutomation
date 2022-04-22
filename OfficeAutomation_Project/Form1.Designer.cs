namespace OfficeAutomation_Project
{
    partial class Form1
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Pulire le risorse in uso.
        /// </summary>
        /// <param name="disposing">ha valore true se le risorse gestite devono essere eliminate, false in caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Codice generato da Progettazione Windows Form

        /// <summary>
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnCrea = new System.Windows.Forms.Button();
            this.btnSalva = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbFont = new System.Windows.Forms.ComboBox();
            this.cmbSize = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.cmbSottolineato = new System.Windows.Forms.ComboBox();
            this.cmbAllineamento = new System.Windows.Forms.ComboBox();
            this.cmbColore = new System.Windows.Forms.ComboBox();
            this.txtTesto = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.btnInserisciTesto = new System.Windows.Forms.Button();
            this.cmbRighe = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.cmbColonne = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.btnInserisciTabella = new System.Windows.Forms.Button();
            this.chkBold = new System.Windows.Forms.CheckBox();
            this.chkItalic = new System.Windows.Forms.CheckBox();
            this.fontDialog1 = new System.Windows.Forms.FontDialog();
            this.btnFontDialog = new System.Windows.Forms.Button();
            this.btnSelezionaTesto = new System.Windows.Forms.Button();
            this.btnStampa = new System.Windows.Forms.Button();
            this.BtnCreatePdf = new System.Windows.Forms.Button();
            this.btnSubAll = new System.Windows.Forms.Button();
            this.btnSearchAndFind = new System.Windows.Forms.Button();
            this.btnVediSelezione = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnSubmit = new System.Windows.Forms.Button();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.label14 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.txtEmail = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.txtNome = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.txtCognome = new System.Windows.Forms.TextBox();
            this.btnSubAll2 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.printDialog1 = new System.Windows.Forms.PrintDialog();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnCrea
            // 
            this.btnCrea.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCrea.Location = new System.Drawing.Point(46, 12);
            this.btnCrea.Name = "btnCrea";
            this.btnCrea.Size = new System.Drawing.Size(235, 45);
            this.btnCrea.TabIndex = 0;
            this.btnCrea.Text = "Crea documento word";
            this.btnCrea.UseVisualStyleBackColor = true;
            this.btnCrea.Click += new System.EventHandler(this.btnCrea_Click);
            // 
            // btnSalva
            // 
            this.btnSalva.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSalva.Location = new System.Drawing.Point(308, 12);
            this.btnSalva.Name = "btnSalva";
            this.btnSalva.Size = new System.Drawing.Size(159, 45);
            this.btnSalva.TabIndex = 1;
            this.btnSalva.Text = "Salva e Chiudi";
            this.btnSalva.UseVisualStyleBackColor = true;
            this.btnSalva.Click += new System.EventHandler(this.btnSalva_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(46, 101);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 20);
            this.label1.TabIndex = 2;
            this.label1.Text = "FONT";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(46, 156);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(50, 20);
            this.label2.TabIndex = 3;
            this.label2.Text = "SIZE";
            // 
            // cmbFont
            // 
            this.cmbFont.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbFont.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbFont.FormattingEnabled = true;
            this.cmbFont.Location = new System.Drawing.Point(137, 95);
            this.cmbFont.Name = "cmbFont";
            this.cmbFont.Size = new System.Drawing.Size(121, 32);
            this.cmbFont.TabIndex = 4;
            // 
            // cmbSize
            // 
            this.cmbSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSize.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbSize.FormattingEnabled = true;
            this.cmbSize.Location = new System.Drawing.Point(137, 150);
            this.cmbSize.Name = "cmbSize";
            this.cmbSize.Size = new System.Drawing.Size(121, 32);
            this.cmbSize.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(46, 194);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(142, 20);
            this.label3.TabIndex = 6;
            this.label3.Text = "SOTTOLINEATO";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(46, 272);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(144, 20);
            this.label4.TabIndex = 7;
            this.label4.Text = "ALLINEAMENTO";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(46, 348);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(82, 20);
            this.label5.TabIndex = 8;
            this.label5.Text = "COLORE";
            // 
            // cmbSottolineato
            // 
            this.cmbSottolineato.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSottolineato.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbSottolineato.FormattingEnabled = true;
            this.cmbSottolineato.Location = new System.Drawing.Point(46, 227);
            this.cmbSottolineato.Name = "cmbSottolineato";
            this.cmbSottolineato.Size = new System.Drawing.Size(421, 32);
            this.cmbSottolineato.TabIndex = 9;
            // 
            // cmbAllineamento
            // 
            this.cmbAllineamento.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbAllineamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbAllineamento.FormattingEnabled = true;
            this.cmbAllineamento.Location = new System.Drawing.Point(46, 295);
            this.cmbAllineamento.Name = "cmbAllineamento";
            this.cmbAllineamento.Size = new System.Drawing.Size(421, 32);
            this.cmbAllineamento.TabIndex = 10;
            // 
            // cmbColore
            // 
            this.cmbColore.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbColore.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbColore.FormattingEnabled = true;
            this.cmbColore.Location = new System.Drawing.Point(46, 371);
            this.cmbColore.Name = "cmbColore";
            this.cmbColore.Size = new System.Drawing.Size(421, 32);
            this.cmbColore.TabIndex = 11;
            // 
            // txtTesto
            // 
            this.txtTesto.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtTesto.ForeColor = System.Drawing.Color.Gray;
            this.txtTesto.Location = new System.Drawing.Point(50, 443);
            this.txtTesto.Name = "txtTesto";
            this.txtTesto.Size = new System.Drawing.Size(417, 29);
            this.txtTesto.TabIndex = 12;
            this.txtTesto.Text = "Inserire il testo da scrivere sul file Word...";
            this.txtTesto.Enter += new System.EventHandler(this.txtTesto_Enter);
            this.txtTesto.Leave += new System.EventHandler(this.txtTesto_Leave);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(46, 420);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(320, 20);
            this.label6.TabIndex = 13;
            this.label6.Text = "TESTO CHE SI DESIDERA INSERIRE";
            // 
            // btnInserisciTesto
            // 
            this.btnInserisciTesto.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnInserisciTesto.Location = new System.Drawing.Point(46, 490);
            this.btnInserisciTesto.Name = "btnInserisciTesto";
            this.btnInserisciTesto.Size = new System.Drawing.Size(421, 38);
            this.btnInserisciTesto.TabIndex = 14;
            this.btnInserisciTesto.Text = "INSERISCI TESTO";
            this.btnInserisciTesto.UseVisualStyleBackColor = true;
            this.btnInserisciTesto.Click += new System.EventHandler(this.btnInserisciTesto_Click);
            // 
            // cmbRighe
            // 
            this.cmbRighe.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbRighe.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbRighe.FormattingEnabled = true;
            this.cmbRighe.Location = new System.Drawing.Point(120, 556);
            this.cmbRighe.Name = "cmbRighe";
            this.cmbRighe.Size = new System.Drawing.Size(68, 32);
            this.cmbRighe.TabIndex = 16;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(59, 565);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(55, 16);
            this.label7.TabIndex = 15;
            this.label7.Text = "RIGHE";
            // 
            // cmbColonne
            // 
            this.cmbColonne.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbColonne.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbColonne.FormattingEnabled = true;
            this.cmbColonne.Location = new System.Drawing.Point(385, 556);
            this.cmbColonne.Name = "cmbColonne";
            this.cmbColonne.Size = new System.Drawing.Size(68, 32);
            this.cmbColonne.TabIndex = 18;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(299, 565);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(80, 16);
            this.label8.TabIndex = 17;
            this.label8.Text = "COLONNE";
            // 
            // btnInserisciTabella
            // 
            this.btnInserisciTabella.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnInserisciTabella.Location = new System.Drawing.Point(100, 627);
            this.btnInserisciTabella.Name = "btnInserisciTabella";
            this.btnInserisciTabella.Size = new System.Drawing.Size(309, 36);
            this.btnInserisciTabella.TabIndex = 19;
            this.btnInserisciTabella.Text = "INSERISCI TABELLA";
            this.btnInserisciTabella.UseVisualStyleBackColor = true;
            this.btnInserisciTabella.Click += new System.EventHandler(this.btnInserisciTabella_Click);
            // 
            // chkBold
            // 
            this.chkBold.AutoSize = true;
            this.chkBold.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkBold.Location = new System.Drawing.Point(308, 101);
            this.chkBold.Name = "chkBold";
            this.chkBold.Size = new System.Drawing.Size(101, 17);
            this.chkBold.TabIndex = 20;
            this.chkBold.Text = "GRASSETTO";
            this.chkBold.UseVisualStyleBackColor = true;
            // 
            // chkItalic
            // 
            this.chkItalic.AutoSize = true;
            this.chkItalic.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkItalic.Location = new System.Drawing.Point(424, 101);
            this.chkItalic.Name = "chkItalic";
            this.chkItalic.Size = new System.Drawing.Size(65, 17);
            this.chkItalic.TabIndex = 21;
            this.chkItalic.Text = "ITALIC";
            this.chkItalic.UseVisualStyleBackColor = true;
            // 
            // btnFontDialog
            // 
            this.btnFontDialog.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFontDialog.Location = new System.Drawing.Point(308, 150);
            this.btnFontDialog.Name = "btnFontDialog";
            this.btnFontDialog.Size = new System.Drawing.Size(159, 32);
            this.btnFontDialog.TabIndex = 22;
            this.btnFontDialog.Text = "FONT DIALOG";
            this.btnFontDialog.UseVisualStyleBackColor = true;
            this.btnFontDialog.Click += new System.EventHandler(this.btnFontDialog_Click);
            // 
            // btnSelezionaTesto
            // 
            this.btnSelezionaTesto.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSelezionaTesto.Location = new System.Drawing.Point(515, 12);
            this.btnSelezionaTesto.Name = "btnSelezionaTesto";
            this.btnSelezionaTesto.Size = new System.Drawing.Size(398, 32);
            this.btnSelezionaTesto.TabIndex = 23;
            this.btnSelezionaTesto.Text = "Seleziona Testo";
            this.btnSelezionaTesto.UseVisualStyleBackColor = true;
            this.btnSelezionaTesto.Click += new System.EventHandler(this.btnSelezionaTesto_Click);
            // 
            // btnStampa
            // 
            this.btnStampa.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnStampa.Location = new System.Drawing.Point(515, 354);
            this.btnStampa.Name = "btnStampa";
            this.btnStampa.Size = new System.Drawing.Size(398, 32);
            this.btnStampa.TabIndex = 24;
            this.btnStampa.Text = "Stampa";
            this.btnStampa.UseVisualStyleBackColor = true;
            this.btnStampa.Click += new System.EventHandler(this.btnStampa_Click);
            // 
            // BtnCreatePdf
            // 
            this.BtnCreatePdf.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BtnCreatePdf.Location = new System.Drawing.Point(515, 297);
            this.BtnCreatePdf.Name = "BtnCreatePdf";
            this.BtnCreatePdf.Size = new System.Drawing.Size(398, 32);
            this.BtnCreatePdf.TabIndex = 25;
            this.BtnCreatePdf.Text = "Stampa In PDF";
            this.BtnCreatePdf.UseVisualStyleBackColor = true;
            this.BtnCreatePdf.Click += new System.EventHandler(this.BtnCreatePdf_Click);
            // 
            // btnSubAll
            // 
            this.btnSubAll.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSubAll.Location = new System.Drawing.Point(515, 183);
            this.btnSubAll.Name = "btnSubAll";
            this.btnSubAll.Size = new System.Drawing.Size(178, 32);
            this.btnSubAll.TabIndex = 26;
            this.btnSubAll.Text = "Sostituisci Tutto";
            this.btnSubAll.UseVisualStyleBackColor = true;
            this.btnSubAll.Click += new System.EventHandler(this.btnSubAll_Click);
            // 
            // btnSearchAndFind
            // 
            this.btnSearchAndFind.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSearchAndFind.Location = new System.Drawing.Point(515, 126);
            this.btnSearchAndFind.Name = "btnSearchAndFind";
            this.btnSearchAndFind.Size = new System.Drawing.Size(398, 32);
            this.btnSearchAndFind.TabIndex = 27;
            this.btnSearchAndFind.Text = "Trova - Sostituisci";
            this.btnSearchAndFind.UseVisualStyleBackColor = true;
            this.btnSearchAndFind.Click += new System.EventHandler(this.btnSearchAndFind_Click);
            // 
            // btnVediSelezione
            // 
            this.btnVediSelezione.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnVediSelezione.Location = new System.Drawing.Point(515, 69);
            this.btnVediSelezione.Name = "btnVediSelezione";
            this.btnVediSelezione.Size = new System.Drawing.Size(398, 32);
            this.btnVediSelezione.TabIndex = 28;
            this.btnVediSelezione.Text = "Vedi Selezione";
            this.btnVediSelezione.UseVisualStyleBackColor = true;
            this.btnVediSelezione.Click += new System.EventHandler(this.btnVediSelezione_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnSubmit);
            this.groupBox1.Controls.Add(this.dateTimePicker1);
            this.groupBox1.Controls.Add(this.label14);
            this.groupBox1.Controls.Add(this.comboBox1);
            this.groupBox1.Controls.Add(this.txtPassword);
            this.groupBox1.Controls.Add(this.label13);
            this.groupBox1.Controls.Add(this.txtEmail);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.label11);
            this.groupBox1.Controls.Add(this.txtNome);
            this.groupBox1.Controls.Add(this.label12);
            this.groupBox1.Controls.Add(this.txtCognome);
            this.groupBox1.Location = new System.Drawing.Point(511, 440);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(398, 247);
            this.groupBox1.TabIndex = 29;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Form Di Registrazione";
            // 
            // btnSubmit
            // 
            this.btnSubmit.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSubmit.Location = new System.Drawing.Point(73, 209);
            this.btnSubmit.Name = "btnSubmit";
            this.btnSubmit.Size = new System.Drawing.Size(255, 32);
            this.btnSubmit.TabIndex = 31;
            this.btnSubmit.Text = "Registrati Ora!!";
            this.btnSubmit.UseVisualStyleBackColor = true;
            this.btnSubmit.Click += new System.EventHandler(this.btnSubmit_Click);
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(125, 96);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(240, 20);
            this.dateTimePicker1.TabIndex = 41;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.Location = new System.Drawing.Point(18, 168);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(76, 16);
            this.label14.TabIndex = 35;
            this.label14.Text = "Password";
            // 
            // comboBox1
            // 
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(-130, 181);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(68, 32);
            this.comboBox1.TabIndex = 18;
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(125, 164);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '*';
            this.txtPassword.Size = new System.Drawing.Size(240, 20);
            this.txtPassword.TabIndex = 40;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(18, 133);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(52, 16);
            this.label13.TabIndex = 34;
            this.label13.Text = "E-mail";
            // 
            // txtEmail
            // 
            this.txtEmail.Location = new System.Drawing.Point(125, 130);
            this.txtEmail.Name = "txtEmail";
            this.txtEmail.Size = new System.Drawing.Size(240, 20);
            this.txtEmail.TabIndex = 39;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(18, 28);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(74, 16);
            this.label10.TabIndex = 31;
            this.label10.Text = "Cognome";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(18, 63);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(49, 16);
            this.label11.TabIndex = 32;
            this.label11.Text = "Nome";
            // 
            // txtNome
            // 
            this.txtNome.Location = new System.Drawing.Point(125, 62);
            this.txtNome.Name = "txtNome";
            this.txtNome.Size = new System.Drawing.Size(240, 20);
            this.txtNome.TabIndex = 37;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(18, 98);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(98, 16);
            this.label12.TabIndex = 33;
            this.label12.Text = "Data Nascita";
            // 
            // txtCognome
            // 
            this.txtCognome.Location = new System.Drawing.Point(125, 28);
            this.txtCognome.Name = "txtCognome";
            this.txtCognome.Size = new System.Drawing.Size(240, 20);
            this.txtCognome.TabIndex = 36;
            // 
            // btnSubAll2
            // 
            this.btnSubAll2.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSubAll2.Location = new System.Drawing.Point(731, 240);
            this.btnSubAll2.Name = "btnSubAll2";
            this.btnSubAll2.Size = new System.Drawing.Size(182, 32);
            this.btnSubAll2.TabIndex = 30;
            this.btnSubAll2.Text = "Sostituisci Tutto 2";
            this.btnSubAll2.UseVisualStyleBackColor = true;
            this.btnSubAll2.Click += new System.EventHandler(this.btnSubAll2_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // printDialog1
            // 
            this.printDialog1.UseEXDialog = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(949, 699);
            this.Controls.Add(this.btnSubAll2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnVediSelezione);
            this.Controls.Add(this.btnSearchAndFind);
            this.Controls.Add(this.btnSubAll);
            this.Controls.Add(this.BtnCreatePdf);
            this.Controls.Add(this.btnStampa);
            this.Controls.Add(this.btnSelezionaTesto);
            this.Controls.Add(this.btnFontDialog);
            this.Controls.Add(this.chkItalic);
            this.Controls.Add(this.chkBold);
            this.Controls.Add(this.btnInserisciTabella);
            this.Controls.Add(this.cmbColonne);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.cmbRighe);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.btnInserisciTesto);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtTesto);
            this.Controls.Add(this.cmbColore);
            this.Controls.Add(this.cmbAllineamento);
            this.Controls.Add(this.cmbSottolineato);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cmbSize);
            this.Controls.Add(this.cmbFont);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnSalva);
            this.Controls.Add(this.btnCrea);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCrea;
        private System.Windows.Forms.Button btnSalva;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmbFont;
        private System.Windows.Forms.ComboBox cmbSize;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cmbSottolineato;
        private System.Windows.Forms.ComboBox cmbAllineamento;
        private System.Windows.Forms.ComboBox cmbColore;
        private System.Windows.Forms.TextBox txtTesto;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnInserisciTesto;
        private System.Windows.Forms.ComboBox cmbRighe;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cmbColonne;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button btnInserisciTabella;
        private System.Windows.Forms.CheckBox chkBold;
        private System.Windows.Forms.CheckBox chkItalic;
        private System.Windows.Forms.FontDialog fontDialog1;
        private System.Windows.Forms.Button btnFontDialog;
        private System.Windows.Forms.Button btnSelezionaTesto;
        private System.Windows.Forms.Button btnStampa;
        private System.Windows.Forms.Button BtnCreatePdf;
        private System.Windows.Forms.Button btnSubAll;
        private System.Windows.Forms.Button btnSearchAndFind;
        private System.Windows.Forms.Button btnVediSelezione;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnSubmit;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox txtEmail;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox txtNome;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox txtCognome;
        private System.Windows.Forms.Button btnSubAll2;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.PrintDialog printDialog1;
    }
}


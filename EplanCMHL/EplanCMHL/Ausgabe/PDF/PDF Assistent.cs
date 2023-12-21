// PDF Assistent.cs
//
// Integriert einen Dialog um automatisiert PDF-Dateien in definierte Ordner abzulegen.
// Der Aufruf erfolgt entweder über Menüpunkt (unter Werkzeuge > Erweiterungen > PDF Assistent)
// oder automatisch beim Schließen oder Beenden.
//
// Der Ursprung für diesen Assistenten war das Skript "PDFbyProjectClose"
//
// Copyright by Frank Schöneck, 2014-2023
// letzte Änderung:
// V1.0.0, 27.01.2014, Frank Schöneck,	Projektbeginn
// V1.1.0, 05.03.2014, Frank Schöneck,	Ausgabe Farbe/Schwarz-Weiß ergänzt
//										Seitenfilter Aktivierung ergänzt
// V1.1.1, 03.07.2014, Frank Schöneck,	Fehler wenn kein geöffnetes Projekt vorhanden ist behoben
// V1.2.0, 09.07.2014, Frank Schöneck,	Neue Auswahlmöglichkeit für User-Speicherort hinzugefügt
//										Datum-/Zeit-Stempel kann nun auch vor den Dateinamen gesetzt werden
// V1.3.0, 04.02.2015, Frank Schöneck,	Datei Handling verbessert (Abfrage ob Datei überschrieben werden kann, Prüfung ob Datei in Verwendung)
// V1.4.0, 11.01.2016, Frank Schöneck,	Ausgabe in Farbig/Schwarz-Weiß/Graustufen erweitert
// V1.5.0, 26.04.2016, Frank Schöneck,	Die Anzeige des Speicherort hat nun ein Kontextmenü, damit kann man diesen "im Windows-Explorer öffnen"
// V1.5.0, 26.04.2016, Frank Schöneck,	Die Anzeige des Speicherort hat nun ein Kontextmenü, damit kann man diesen "im Windows-Explorer öffnen"
// V1.6.0, 13.09.2016, Frank Schöneck,	Möglichkeit den Dateinamen um Präfix und Suffix zu ergänzen hinzugefügt, Einstellungen speichern verbessert
// V1.6.1, 28.11.2016, Frank Schöneck,	Fehler in Einstellungen speichern behoben, Ausgabe in ... wurde nicht richtig eingelesen
// V1.7.0, 28.11.2016, Frank Schöneck,	Neue Checkboxen für "Anzeigen" und "Ordner öffnen" hinzugefügt
// V1.8.0, 17.01.2017, Frank Schöneck,	Neue Checkbox für "Auswertungen aktualisieren" hinzugefügt
// V2.0.0, 16.04.2018, Frank Schöneck,	Es dürfen nun mehrere Projekte markiert sein und werden exportiert (ab Eplan V2.6)
// V2.1.0, 30.04.2018, Frank Schöneck,	Ausgabe nach erweitert um die Möglichkeit das Ausgabeverzeichnis aus Datei im Projektordner auszulesen
// V2.1.1, 25.05.2018, Frank Schöneck,	Neue Syntax für Settings verwendet
// V2.1.2, 28.02.2019, Frank Schöneck,	Fehler in Einstellungen speichern behoben, Projektbezogener-Speicherort wurde nicht gespeichert
// V2.1.3, 06.05.2020, Frank Schöneck,	Fehler in Einstellungen speichern behoben, es muss das Script "PDF Assistent_Einstellungen löschen.cs" ausgeführt werden.
// V2.2.0, 19.05.2020, Frank Schöneck,	Begriffe in der Auswahlliste bei "Ausgabe nach" überarbeitet bzw. an Windows 10 angepasst
// V3.0.0, 17.10.2023, Frank Schöneck,	Angepasst an die Ribbon-Technik der EPLAN Plattform 2022.
// V3.1.0, 17.11.2023, Frank Schöneck,	Möglichkeit des Ausführens ohne Dialog hinzugefügt, Aktion: PDFAssistent_Start_Ohne_Dialog
//
// für Eplan Electric P8, ab V2022 Update 2
//

/*
The following compiler directive is necessary to enable editing scripts
within Visual Studio.

It requires that the "Conditional compilation symbol" SCRIPTENV be defined 
in the Visual Studio project properties

This is because EPLAN's internal scripting engine already adds "using directives"
when you load the script in EPLAN. Having them twice would cause errors.
*/
#if SCRIPTENV
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Scripting;
using Eplan.EplApi.Base;
using System.Linq;
#endif

/*
On the other hand, some namespaces are not automatically added by EPLAN when
you load a script. Those have to be outside of the previous conditional compiler directive
*/
using System;
using System.Windows.Forms;
using System.Xml;
using System.IO;
using System.Diagnostics;

public partial class frmPDFAssistent : System.Windows.Forms.Form
{
	private Button btnAbbrechen;
	private Button btnOK;
	private CheckBox chkEinstellungSpeichern;
	private TabControl tabControl1;
	private TabPage tabPage1;
	private Button btnOrdnerAuswahl;
	private TextBox txtSpeicherort;
	private TextBox txtDateiname;
	private CheckBox chkDatumStempel;
	private CheckBox chkZeitStempel;
	private ComboBox cboAusgabeNach;
	private Label label3;
	private Label label2;
	private Label label1;
	private TabPage tabPage2;
	private GroupBox groupBox1;
	private CheckBox chkByEplanEnd;
	private CheckBox chkByProjectClose;
	private Label label4;
	private CheckBox chkOhneNachfrage;
	private Button btnSpeichern;
	private GroupBox groupBox2;
	private Button btnProjektOrdnerAuswahl;
	private TextBox txtProjektGespeichertInOrdner;
	private CheckBox chkIstInProjektOrdner;
	private CheckBox chkPageFilter;
	private CheckBox chkStempelVorne;
	private ComboBox comboBoxAusgabe;
	private Label label5;
	private ContextMenuStrip contextMenuSpeicherort;
	private ToolStripMenuItem toolStripMenuItem1;
	private TextBox txtPräfix;
	private Label label6;
	private CheckBox chkOrdnerÖffnen;
	private CheckBox chkOpenPDF;
	private CheckBox chkAuswertungenAktualisieren;

	#region Vom Windows Form-Designer generierter Code

	/// <summary>
	/// Erforderliche Designervariable.
	/// </summary>
	private System.ComponentModel.IContainer components = null;

	/// <summary>
	/// Verwendete Ressourcen bereinigen.
	/// </summary>
	/// <param name="disposing">True, wenn verwaltete Ressourcen
	/// gelöscht werden sollen; andernfalls False.</param>
	protected override void Dispose(bool disposing)
	{
		if (disposing && (components != null))
		{
			components.Dispose();
		}
		base.Dispose(disposing);
	}

	/// <summary>
	/// Erforderliche Methode für die Designerunterstützung.
	/// Der Inhalt der Methode darf nicht mit dem Code-Editor
	/// geändert werden.
	/// </summary>
	private void InitializeComponent()
	{
		this.components = new System.ComponentModel.Container();
		this.btnAbbrechen = new System.Windows.Forms.Button();
		this.btnOK = new System.Windows.Forms.Button();
		this.chkEinstellungSpeichern = new System.Windows.Forms.CheckBox();
		this.tabControl1 = new System.Windows.Forms.TabControl();
		this.tabPage1 = new System.Windows.Forms.TabPage();
		this.chkOpenPDF = new System.Windows.Forms.CheckBox();
		this.chkOrdnerÖffnen = new System.Windows.Forms.CheckBox();
		this.btnOrdnerAuswahl = new System.Windows.Forms.Button();
		this.txtSpeicherort = new System.Windows.Forms.TextBox();
		this.contextMenuSpeicherort = new System.Windows.Forms.ContextMenuStrip(this.components);
		this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
		this.txtPräfix = new System.Windows.Forms.TextBox();
		this.txtDateiname = new System.Windows.Forms.TextBox();
		this.chkDatumStempel = new System.Windows.Forms.CheckBox();
		this.chkZeitStempel = new System.Windows.Forms.CheckBox();
		this.cboAusgabeNach = new System.Windows.Forms.ComboBox();
		this.label6 = new System.Windows.Forms.Label();
		this.label4 = new System.Windows.Forms.Label();
		this.label3 = new System.Windows.Forms.Label();
		this.label2 = new System.Windows.Forms.Label();
		this.label1 = new System.Windows.Forms.Label();
		this.tabPage2 = new System.Windows.Forms.TabPage();
		this.label5 = new System.Windows.Forms.Label();
		this.comboBoxAusgabe = new System.Windows.Forms.ComboBox();
		this.chkAuswertungenAktualisieren = new System.Windows.Forms.CheckBox();
		this.chkStempelVorne = new System.Windows.Forms.CheckBox();
		this.chkPageFilter = new System.Windows.Forms.CheckBox();
		this.groupBox2 = new System.Windows.Forms.GroupBox();
		this.btnProjektOrdnerAuswahl = new System.Windows.Forms.Button();
		this.txtProjektGespeichertInOrdner = new System.Windows.Forms.TextBox();
		this.chkIstInProjektOrdner = new System.Windows.Forms.CheckBox();
		this.chkOhneNachfrage = new System.Windows.Forms.CheckBox();
		this.btnSpeichern = new System.Windows.Forms.Button();
		this.groupBox1 = new System.Windows.Forms.GroupBox();
		this.chkByEplanEnd = new System.Windows.Forms.CheckBox();
		this.chkByProjectClose = new System.Windows.Forms.CheckBox();
		this.tabControl1.SuspendLayout();
		this.tabPage1.SuspendLayout();
		this.contextMenuSpeicherort.SuspendLayout();
		this.tabPage2.SuspendLayout();
		this.groupBox2.SuspendLayout();
		this.groupBox1.SuspendLayout();
		this.SuspendLayout();
		// 
		// btnAbbrechen
		// 
		this.btnAbbrechen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
		this.btnAbbrechen.DialogResult = System.Windows.Forms.DialogResult.Cancel;
		this.btnAbbrechen.Location = new System.Drawing.Point(748, 639);
		this.btnAbbrechen.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.btnAbbrechen.Name = "btnAbbrechen";
		this.btnAbbrechen.Size = new System.Drawing.Size(174, 44);
		this.btnAbbrechen.TabIndex = 0;
		this.btnAbbrechen.Text = "Abbrechen";
		this.btnAbbrechen.UseVisualStyleBackColor = true;
		this.btnAbbrechen.Click += new System.EventHandler(this.btnAbbrechen_Click);
		// 
		// btnOK
		// 
		this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
		this.btnOK.Location = new System.Drawing.Point(546, 639);
		this.btnOK.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.btnOK.Name = "btnOK";
		this.btnOK.Size = new System.Drawing.Size(174, 44);
		this.btnOK.TabIndex = 0;
		this.btnOK.Text = "OK";
		this.btnOK.UseVisualStyleBackColor = true;
		this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
		// 
		// chkEinstellungSpeichern
		// 
		this.chkEinstellungSpeichern.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
		this.chkEinstellungSpeichern.AutoSize = true;
		this.chkEinstellungSpeichern.Location = new System.Drawing.Point(17, 650);
		this.chkEinstellungSpeichern.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.chkEinstellungSpeichern.Name = "chkEinstellungSpeichern";
		this.chkEinstellungSpeichern.Size = new System.Drawing.Size(245, 29);
		this.chkEinstellungSpeichern.TabIndex = 3;
		this.chkEinstellungSpeichern.Text = "Einstellungen speichern";
		this.chkEinstellungSpeichern.UseVisualStyleBackColor = true;
		// 
		// tabControl1
		// 
		this.tabControl1.Controls.Add(this.tabPage1);
		this.tabControl1.Controls.Add(this.tabPage2);
		this.tabControl1.Location = new System.Drawing.Point(17, 22);
		this.tabControl1.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.tabControl1.Name = "tabControl1";
		this.tabControl1.SelectedIndex = 0;
		this.tabControl1.Size = new System.Drawing.Size(913, 585);
		this.tabControl1.TabIndex = 4;
		// 
		// tabPage1
		// 
		this.tabPage1.BackColor = System.Drawing.Color.Transparent;
		this.tabPage1.Controls.Add(this.chkOpenPDF);
		this.tabPage1.Controls.Add(this.chkOrdnerÖffnen);
		this.tabPage1.Controls.Add(this.btnOrdnerAuswahl);
		this.tabPage1.Controls.Add(this.txtSpeicherort);
		this.tabPage1.Controls.Add(this.txtPräfix);
		this.tabPage1.Controls.Add(this.txtDateiname);
		this.tabPage1.Controls.Add(this.chkDatumStempel);
		this.tabPage1.Controls.Add(this.chkZeitStempel);
		this.tabPage1.Controls.Add(this.cboAusgabeNach);
		this.tabPage1.Controls.Add(this.label6);
		this.tabPage1.Controls.Add(this.label4);
		this.tabPage1.Controls.Add(this.label3);
		this.tabPage1.Controls.Add(this.label2);
		this.tabPage1.Controls.Add(this.label1);
		this.tabPage1.Location = new System.Drawing.Point(4, 33);
		this.tabPage1.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.tabPage1.Name = "tabPage1";
		this.tabPage1.Padding = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.tabPage1.Size = new System.Drawing.Size(905, 548);
		this.tabPage1.TabIndex = 0;
		this.tabPage1.Text = "Ausgabe";
		// 
		// chkOpenPDF
		// 
		this.chkOpenPDF.AutoSize = true;
		this.chkOpenPDF.Location = new System.Drawing.Point(26, 388);
		this.chkOpenPDF.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.chkOpenPDF.Name = "chkOpenPDF";
		this.chkOpenPDF.Size = new System.Drawing.Size(276, 29);
		this.chkOpenPDF.TabIndex = 17;
		this.chkOpenPDF.Text = "Nach dem Export Anzeigen";
		this.chkOpenPDF.UseVisualStyleBackColor = true;
		// 
		// chkOrdnerÖffnen
		// 
		this.chkOrdnerÖffnen.AutoSize = true;
		this.chkOrdnerÖffnen.Location = new System.Drawing.Point(26, 430);
		this.chkOrdnerÖffnen.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.chkOrdnerÖffnen.Name = "chkOrdnerÖffnen";
		this.chkOrdnerÖffnen.Size = new System.Drawing.Size(313, 29);
		this.chkOrdnerÖffnen.TabIndex = 18;
		this.chkOrdnerÖffnen.Text = "Nach dem Export Ordner öffnen";
		this.chkOrdnerÖffnen.UseVisualStyleBackColor = true;
		// 
		// btnOrdnerAuswahl
		// 
		this.btnOrdnerAuswahl.Location = new System.Drawing.Point(814, 321);
		this.btnOrdnerAuswahl.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.btnOrdnerAuswahl.Name = "btnOrdnerAuswahl";
		this.btnOrdnerAuswahl.Size = new System.Drawing.Size(51, 37);
		this.btnOrdnerAuswahl.TabIndex = 16;
		this.btnOrdnerAuswahl.Text = "...";
		this.btnOrdnerAuswahl.UseVisualStyleBackColor = true;
		this.btnOrdnerAuswahl.Click += new System.EventHandler(this.btnOrdnerAuswahl_Click);
		// 
		// txtSpeicherort
		// 
		this.txtSpeicherort.ContextMenuStrip = this.contextMenuSpeicherort;
		this.txtSpeicherort.Location = new System.Drawing.Point(26, 323);
		this.txtSpeicherort.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.txtSpeicherort.Name = "txtSpeicherort";
		this.txtSpeicherort.ReadOnly = true;
		this.txtSpeicherort.Size = new System.Drawing.Size(774, 29);
		this.txtSpeicherort.TabIndex = 15;
		// 
		// contextMenuSpeicherort
		// 
		this.contextMenuSpeicherort.ImageScalingSize = new System.Drawing.Size(28, 28);
		this.contextMenuSpeicherort.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
			this.toolStripMenuItem1});
		this.contextMenuSpeicherort.Name = "contextMenuSpeicherort";
		this.contextMenuSpeicherort.ShowImageMargin = false;
		this.contextMenuSpeicherort.Size = new System.Drawing.Size(326, 40);
		this.contextMenuSpeicherort.Click += new System.EventHandler(this.contextMenuSpeicherort_Click);
		// 
		// toolStripMenuItem1
		// 
		this.toolStripMenuItem1.Name = "toolStripMenuItem1";
		this.toolStripMenuItem1.Size = new System.Drawing.Size(325, 36);
		this.toolStripMenuItem1.Text = "im Windows-Explorer öffnen";
		// 
		// txtPräfix
		// 
		this.txtPräfix.Location = new System.Drawing.Point(277, 98);
		this.txtPräfix.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.txtPräfix.Name = "txtPräfix";
		this.txtPräfix.Size = new System.Drawing.Size(585, 29);
		this.txtPräfix.TabIndex = 11;
		this.txtPräfix.TextChanged += new System.EventHandler(this.txtPräfix_TextChanged);
		// 
		// txtDateiname
		// 
		this.txtDateiname.Location = new System.Drawing.Point(26, 251);
		this.txtDateiname.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.txtDateiname.Name = "txtDateiname";
		this.txtDateiname.Size = new System.Drawing.Size(836, 29);
		this.txtDateiname.TabIndex = 14;
		// 
		// chkDatumStempel
		// 
		this.chkDatumStempel.AutoSize = true;
		this.chkDatumStempel.Location = new System.Drawing.Point(277, 151);
		this.chkDatumStempel.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.chkDatumStempel.Name = "chkDatumStempel";
		this.chkDatumStempel.Size = new System.Drawing.Size(174, 29);
		this.chkDatumStempel.TabIndex = 12;
		this.chkDatumStempel.Text = "Datum-Stempel";
		this.chkDatumStempel.UseVisualStyleBackColor = true;
		this.chkDatumStempel.CheckedChanged += new System.EventHandler(this.chkDatumStempel_CheckedChanged);
		// 
		// chkZeitStempel
		// 
		this.chkZeitStempel.AutoSize = true;
		this.chkZeitStempel.Location = new System.Drawing.Point(468, 151);
		this.chkZeitStempel.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.chkZeitStempel.Name = "chkZeitStempel";
		this.chkZeitStempel.Size = new System.Drawing.Size(149, 29);
		this.chkZeitStempel.TabIndex = 13;
		this.chkZeitStempel.Text = "Zeit-Stempel";
		this.chkZeitStempel.UseVisualStyleBackColor = true;
		this.chkZeitStempel.CheckedChanged += new System.EventHandler(this.chkZeitStempel_CheckedChanged);
		// 
		// cboAusgabeNach
		// 
		this.cboAusgabeNach.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.cboAusgabeNach.Items.AddRange(new object[] {
			"in den Projekt-Ordner",
			"Ausgabeverzeichnis aus Einstellungen: PDF-Export",
			"in den Ordner eine Ebene über dem Projekt-Ordner",
			"nach \"Dokumente\"",
			"auf den \"Desktop\"",
			"gleicher Pfad wie Projekt nur auf anderes Laufwerk",
			"Benutzerbezogener-Speicherort",
			"Projektbezogener-Speicherort"});
		this.cboAusgabeNach.Location = new System.Drawing.Point(176, 31);
		this.cboAusgabeNach.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.cboAusgabeNach.Name = "cboAusgabeNach";
		this.cboAusgabeNach.Size = new System.Drawing.Size(686, 32);
		this.cboAusgabeNach.TabIndex = 10;
		this.cboAusgabeNach.SelectedIndexChanged += new System.EventHandler(this.cboAusgabeNach_SelectedIndexChanged);
		// 
		// label6
		// 
		this.label6.AutoSize = true;
		this.label6.Location = new System.Drawing.Point(20, 98);
		this.label6.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
		this.label6.Name = "label6";
		this.label6.Size = new System.Drawing.Size(212, 25);
		this.label6.TabIndex = 7;
		this.label6.Text = "PDF-Dateiname Präfix:";
		// 
		// label4
		// 
		this.label4.AutoSize = true;
		this.label4.Location = new System.Drawing.Point(20, 153);
		this.label4.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
		this.label4.Name = "label4";
		this.label4.Size = new System.Drawing.Size(212, 25);
		this.label4.TabIndex = 7;
		this.label4.Text = "PDF-Dateiname Suffix:";
		// 
		// label3
		// 
		this.label3.AutoSize = true;
		this.label3.Location = new System.Drawing.Point(20, 222);
		this.label3.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
		this.label3.Name = "label3";
		this.label3.Size = new System.Drawing.Size(380, 25);
		this.label3.TabIndex = 7;
		this.label3.Text = "PDF-Dateiname (ohne Erweiterung (.pdf)):";
		// 
		// label2
		// 
		this.label2.AutoSize = true;
		this.label2.Location = new System.Drawing.Point(20, 294);
		this.label2.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
		this.label2.Name = "label2";
		this.label2.Size = new System.Drawing.Size(118, 25);
		this.label2.TabIndex = 8;
		this.label2.Text = "Speicherort:";
		// 
		// label1
		// 
		this.label1.AutoSize = true;
		this.label1.Location = new System.Drawing.Point(20, 37);
		this.label1.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(145, 25);
		this.label1.TabIndex = 9;
		this.label1.Text = "Ausgabe nach:";
		// 
		// tabPage2
		// 
		this.tabPage2.BackColor = System.Drawing.Color.Transparent;
		this.tabPage2.Controls.Add(this.label5);
		this.tabPage2.Controls.Add(this.comboBoxAusgabe);
		this.tabPage2.Controls.Add(this.chkAuswertungenAktualisieren);
		this.tabPage2.Controls.Add(this.chkStempelVorne);
		this.tabPage2.Controls.Add(this.chkPageFilter);
		this.tabPage2.Controls.Add(this.groupBox2);
		this.tabPage2.Controls.Add(this.btnSpeichern);
		this.tabPage2.Controls.Add(this.groupBox1);
		this.tabPage2.Location = new System.Drawing.Point(4, 33);
		this.tabPage2.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.tabPage2.Name = "tabPage2";
		this.tabPage2.Padding = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.tabPage2.Size = new System.Drawing.Size(905, 548);
		this.tabPage2.TabIndex = 1;
		this.tabPage2.Text = "Einstellungen";
		// 
		// label5
		// 
		this.label5.AutoSize = true;
		this.label5.Location = new System.Drawing.Point(37, 345);
		this.label5.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
		this.label5.Name = "label5";
		this.label5.Size = new System.Drawing.Size(111, 25);
		this.label5.TabIndex = 3;
		this.label5.Text = "Ausgabe in";
		// 
		// comboBoxAusgabe
		// 
		this.comboBoxAusgabe.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
		this.comboBoxAusgabe.Items.AddRange(new object[] {
			"Farbig",
			"Schwarz-Weiß",
			"Graustufen"});
		this.comboBoxAusgabe.Location = new System.Drawing.Point(158, 340);
		this.comboBoxAusgabe.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.comboBoxAusgabe.MaxDropDownItems = 4;
		this.comboBoxAusgabe.Name = "comboBoxAusgabe";
		this.comboBoxAusgabe.Size = new System.Drawing.Size(263, 32);
		this.comboBoxAusgabe.TabIndex = 4;
		// 
		// chkAuswertungenAktualisieren
		// 
		this.chkAuswertungenAktualisieren.AutoSize = true;
		this.chkAuswertungenAktualisieren.Location = new System.Drawing.Point(42, 474);
		this.chkAuswertungenAktualisieren.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.chkAuswertungenAktualisieren.Name = "chkAuswertungenAktualisieren";
		this.chkAuswertungenAktualisieren.Size = new System.Drawing.Size(466, 29);
		this.chkAuswertungenAktualisieren.TabIndex = 7;
		this.chkAuswertungenAktualisieren.Text = "Vor der Ausgabe alle Auswertungen aktualisieren";
		this.chkAuswertungenAktualisieren.UseVisualStyleBackColor = true;
		this.chkAuswertungenAktualisieren.CheckedChanged += new System.EventHandler(this.chkStempelVorne_CheckedChanged);
		// 
		// chkStempelVorne
		// 
		this.chkStempelVorne.AutoSize = true;
		this.chkStempelVorne.Location = new System.Drawing.Point(42, 432);
		this.chkStempelVorne.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.chkStempelVorne.Name = "chkStempelVorne";
		this.chkStempelVorne.Size = new System.Drawing.Size(472, 29);
		this.chkStempelVorne.TabIndex = 6;
		this.chkStempelVorne.Text = "Datum- / Zeit-Stempel vor den Dateinamen setzen";
		this.chkStempelVorne.UseVisualStyleBackColor = true;
		this.chkStempelVorne.CheckedChanged += new System.EventHandler(this.chkStempelVorne_CheckedChanged);
		// 
		// chkPageFilter
		// 
		this.chkPageFilter.AutoSize = true;
		this.chkPageFilter.Location = new System.Drawing.Point(42, 390);
		this.chkPageFilter.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.chkPageFilter.Name = "chkPageFilter";
		this.chkPageFilter.Size = new System.Drawing.Size(353, 29);
		this.chkPageFilter.TabIndex = 5;
		this.chkPageFilter.Text = "Seitenfilter (im PDF-Export Schema)";
		this.chkPageFilter.UseVisualStyleBackColor = true;
		this.chkPageFilter.CheckedChanged += new System.EventHandler(this.chkPageFilter_CheckedChanged);
		// 
		// groupBox2
		// 
		this.groupBox2.Controls.Add(this.btnProjektOrdnerAuswahl);
		this.groupBox2.Controls.Add(this.txtProjektGespeichertInOrdner);
		this.groupBox2.Controls.Add(this.chkIstInProjektOrdner);
		this.groupBox2.Controls.Add(this.chkOhneNachfrage);
		this.groupBox2.Location = new System.Drawing.Point(11, 177);
		this.groupBox2.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.groupBox2.Name = "groupBox2";
		this.groupBox2.Padding = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.groupBox2.Size = new System.Drawing.Size(876, 142);
		this.groupBox2.TabIndex = 2;
		this.groupBox2.TabStop = false;
		this.groupBox2.Text = "und zusätzliche Bedingungen erfüllt sind";
		// 
		// btnProjektOrdnerAuswahl
		// 
		this.btnProjektOrdnerAuswahl.Location = new System.Drawing.Point(810, 31);
		this.btnProjektOrdnerAuswahl.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.btnProjektOrdnerAuswahl.Name = "btnProjektOrdnerAuswahl";
		this.btnProjektOrdnerAuswahl.Size = new System.Drawing.Size(51, 37);
		this.btnProjektOrdnerAuswahl.TabIndex = 2;
		this.btnProjektOrdnerAuswahl.Text = "...";
		this.btnProjektOrdnerAuswahl.UseVisualStyleBackColor = true;
		this.btnProjektOrdnerAuswahl.Click += new System.EventHandler(this.btnProjektOrdnerAuswahl_Click);
		// 
		// txtProjektGespeichertInOrdner
		// 
		this.txtProjektGespeichertInOrdner.Location = new System.Drawing.Point(378, 31);
		this.txtProjektGespeichertInOrdner.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.txtProjektGespeichertInOrdner.Name = "txtProjektGespeichertInOrdner";
		this.txtProjektGespeichertInOrdner.Size = new System.Drawing.Size(418, 29);
		this.txtProjektGespeichertInOrdner.TabIndex = 1;
		// 
		// chkIstInProjektOrdner
		// 
		this.chkIstInProjektOrdner.AutoSize = true;
		this.chkIstInProjektOrdner.Location = new System.Drawing.Point(31, 35);
		this.chkIstInProjektOrdner.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.chkIstInProjektOrdner.Name = "chkIstInProjektOrdner";
		this.chkIstInProjektOrdner.Size = new System.Drawing.Size(328, 29);
		this.chkIstInProjektOrdner.TabIndex = 0;
		this.chkIstInProjektOrdner.Text = "wenn Projekt in diesem Ordner ist";
		this.chkIstInProjektOrdner.UseVisualStyleBackColor = true;
		this.chkIstInProjektOrdner.CheckedChanged += new System.EventHandler(this.chkIstInProjektOrdner_CheckedChanged);
		// 
		// chkOhneNachfrage
		// 
		this.chkOhneNachfrage.AutoSize = true;
		this.chkOhneNachfrage.Location = new System.Drawing.Point(31, 100);
		this.chkOhneNachfrage.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.chkOhneNachfrage.Name = "chkOhneNachfrage";
		this.chkOhneNachfrage.Size = new System.Drawing.Size(243, 29);
		this.chkOhneNachfrage.TabIndex = 3;
		this.chkOhneNachfrage.Text = "direkt (ohne Nachfrage)";
		this.chkOhneNachfrage.UseVisualStyleBackColor = true;
		// 
		// btnSpeichern
		// 
		this.btnSpeichern.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
		this.btnSpeichern.Location = new System.Drawing.Point(713, 482);
		this.btnSpeichern.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.btnSpeichern.Name = "btnSpeichern";
		this.btnSpeichern.Size = new System.Drawing.Size(174, 44);
		this.btnSpeichern.TabIndex = 8;
		this.btnSpeichern.Text = "Speichern";
		this.btnSpeichern.UseVisualStyleBackColor = true;
		this.btnSpeichern.Click += new System.EventHandler(this.btnSpeichern_Click);
		// 
		// groupBox1
		// 
		this.groupBox1.BackColor = System.Drawing.Color.Transparent;
		this.groupBox1.Controls.Add(this.chkByEplanEnd);
		this.groupBox1.Controls.Add(this.chkByProjectClose);
		this.groupBox1.Location = new System.Drawing.Point(11, 35);
		this.groupBox1.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.groupBox1.Name = "groupBox1";
		this.groupBox1.Padding = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.groupBox1.Size = new System.Drawing.Size(876, 131);
		this.groupBox1.TabIndex = 1;
		this.groupBox1.TabStop = false;
		this.groupBox1.Text = "Ausführen nur";
		// 
		// chkByEplanEnd
		// 
		this.chkByEplanEnd.AutoSize = true;
		this.chkByEplanEnd.Location = new System.Drawing.Point(31, 79);
		this.chkByEplanEnd.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.chkByEplanEnd.Name = "chkByEplanEnd";
		this.chkByEplanEnd.Size = new System.Drawing.Size(256, 29);
		this.chkByEplanEnd.TabIndex = 1;
		this.chkByEplanEnd.Text = "wenn Eplan beendet wird";
		this.chkByEplanEnd.UseVisualStyleBackColor = true;
		// 
		// chkByProjectClose
		// 
		this.chkByProjectClose.AutoSize = true;
		this.chkByProjectClose.Location = new System.Drawing.Point(31, 37);
		this.chkByProjectClose.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.chkByProjectClose.Name = "chkByProjectClose";
		this.chkByProjectClose.Size = new System.Drawing.Size(305, 29);
		this.chkByProjectClose.TabIndex = 0;
		this.chkByProjectClose.Text = "wenn Projekt geschlossen wird";
		this.chkByProjectClose.UseVisualStyleBackColor = true;
		// 
		// frmPDFAssistent
		// 
		this.AcceptButton = this.btnOK;
		this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
		this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.CancelButton = this.btnAbbrechen;
		this.ClientSize = new System.Drawing.Size(952, 705);
		this.Controls.Add(this.tabControl1);
		this.Controls.Add(this.chkEinstellungSpeichern);
		this.Controls.Add(this.btnOK);
		this.Controls.Add(this.btnAbbrechen);
		this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
		this.Margin = new System.Windows.Forms.Padding(6, 6, 6, 6);
		this.MaximizeBox = false;
		this.MinimizeBox = false;
		this.Name = "frmPDFAssistent";
		this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
		this.Text = "PDF-Export (Assistent)";
		this.Load += new System.EventHandler(this.frmPDFAssistent_Load);
		this.tabControl1.ResumeLayout(false);
		this.tabPage1.ResumeLayout(false);
		this.tabPage1.PerformLayout();
		this.contextMenuSpeicherort.ResumeLayout(false);
		this.tabPage2.ResumeLayout(false);
		this.tabPage2.PerformLayout();
		this.groupBox2.ResumeLayout(false);
		this.groupBox2.PerformLayout();
		this.groupBox1.ResumeLayout(false);
		this.groupBox1.PerformLayout();
		this.ResumeLayout(false);
		this.PerformLayout();

	}

	public frmPDFAssistent()
	{
		InitializeComponent();
	}

	#endregion

	string m_TabName = "Werkzeuge";
	string m_commandGroupName = "Erweiterungen";
	string m_commandName = "PDF Assistent";
	string m_SVGstringIcon = @"<svg version='1.1' id='Layer_1' xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink' x='0px' y='0px'
		viewBox='0 0 512 512' style='enable-background:new 0 0 512 512;' xml:space='preserve'>
		<path style='fill:#464646;' d='M437.456,512H21.212c-8.166,0-14.786-6.621-14.786-14.786V256.915c0-8.165,6.62-14.786,14.786-14.786
			s14.786,6.621,14.786,14.786v225.512h386.671v-32.939c0-8.165,6.62-14.786,14.786-14.786s14.786,6.621,14.786,14.786v47.725
			C452.242,505.379,445.622,512,437.456,512z'/>
		<polygon style='fill:#99CCFF;' points='21.212,177.092 21.212,172.3 176.068,14.786 176.068,177.092 	'/>
		<path style='fill:#464646;' d='M490.791,204.634h-38.549V14.786c0-8.165-6.62-14.786-14.786-14.786H176.068
			c-0.067,0-0.132,0.009-0.198,0.01c-0.359,0.004-0.717,0.022-1.075,0.053c-0.12,0.01-0.241,0.021-0.361,0.034
			c-0.41,0.046-0.816,0.105-1.22,0.185c-0.031,0.006-0.061,0.009-0.092,0.015c-0.432,0.089-0.858,0.2-1.28,0.325
			c-0.111,0.033-0.22,0.071-0.33,0.106c-0.322,0.105-0.642,0.22-0.958,0.347c-0.108,0.044-0.217,0.086-0.324,0.132
			c-0.807,0.346-1.585,0.766-2.326,1.257c-0.102,0.067-0.2,0.139-0.3,0.207c-0.274,0.191-0.541,0.392-0.803,0.603
			c-0.099,0.08-0.198,0.157-0.294,0.24c-0.339,0.287-0.67,0.586-0.985,0.906L10.668,161.935c-0.346,0.352-0.671,0.719-0.977,1.1
			c-0.182,0.226-0.342,0.463-0.509,0.696c-0.112,0.157-0.234,0.309-0.34,0.47c-0.194,0.294-0.364,0.599-0.534,0.903
			c-0.062,0.112-0.133,0.219-0.192,0.333c-0.166,0.315-0.308,0.639-0.449,0.963c-0.05,0.114-0.108,0.225-0.155,0.34
			c-0.126,0.312-0.231,0.628-0.336,0.946c-0.046,0.139-0.099,0.274-0.14,0.413c-0.087,0.294-0.152,0.591-0.22,0.889
			c-0.04,0.172-0.087,0.34-0.121,0.515c-0.053,0.274-0.084,0.55-0.121,0.825c-0.027,0.201-0.062,0.399-0.081,0.6
			c-0.025,0.268-0.03,0.535-0.04,0.801c-0.007,0.191-0.028,0.38-0.028,0.571v4.792c0,8.165,6.62,14.786,14.786,14.786h154.855
			c8.166,0,14.786-6.621,14.786-14.786V29.572h231.816v175.062H196.518c-8.166,0-14.786,6.621-14.786,14.786v163.705
			c0,8.165,6.62,14.786,14.786,14.786h294.272c8.166,0,14.786-6.621,14.786-14.786V219.421
			C505.577,211.256,498.957,204.634,490.791,204.634z M51.772,162.308l47.938-48.76l61.571-62.63v111.39L51.772,162.308
			L51.772,162.308z M476.005,368.339h-264.7V234.207h264.7V368.339z'/>
		<path style='fill:#464646;' d='M246.08,260.736c0-3.2,2.925-6.015,7.375-6.015h26.322c16.785,0,30.008,7.934,30.008,29.433v0.64
			c0,21.499-13.733,29.689-31.28,29.689h-12.589v27.641c0,4.096-4.959,6.142-9.919,6.142s-9.919-2.048-9.919-6.142L246.08,260.736
			L246.08,260.736z M265.916,272.124v27.002h12.589c7.121,0,11.444-4.096,11.444-12.797v-1.406c0-8.703-4.323-12.797-11.444-12.797
			h-12.589V272.124z'/>
		<path style='fill:#464646;' d='M349.586,254.721c17.548,0,31.282,8.19,31.282,30.202v33.145c0,22.011-13.733,30.201-31.282,30.201
			h-22.507c-5.214,0-8.647-2.815-8.647-6.014v-81.518c0-3.2,3.433-6.015,8.647-6.015h22.507V254.721z M338.269,272.124v58.739h11.317
			c7.121,0,11.444-4.096,11.444-12.796v-33.145c0-8.703-4.323-12.797-11.444-12.797h-11.317V272.124z'/>
		<path style='fill:#464646;' d='M393.458,260.863c0-4.096,4.323-6.142,8.647-6.142h44.125c4.196,0,5.977,4.479,5.977,8.574
			c0,4.735-2.162,8.83-5.977,8.83h-32.935v21.628h19.201c3.815,0,5.977,3.711,5.977,7.806c0,3.456-1.78,7.55-5.977,7.55h-19.201
			v33.016c0,4.096-4.959,6.142-9.919,6.142c-4.959,0-9.919-2.048-9.919-6.142V260.863z'/>
		</svg>";


	// //Assistent ohne Dialog direkt ausführen (Ohne Nachfrage ausführen)
[Start]
	public void PDFAssistent_ausführen(string Param1)
	{
		
		if (Param1 == "0")
			{
				PDFAssistent_Start_Ohne_Dialog();
			}
			else
			{
				PDFAssistent_Start();
			}
	}
	

	//Assistent Form starten
	
	public void PDFAssistent_Start()
	{
		ActionCallingContext acc = new ActionCallingContext();
		CommandLineInterpreter oCLI = new CommandLineInterpreter();
		string sProjects = string.Empty;
		//Aktuell markierte Projekte ermitteln
		acc.AddParameter("TYPE", "PROJECTS");
		oCLI.Execute("selectionset", acc);
		acc.GetParameter("PROJECTS", ref sProjects);

		if (sProjects == string.Empty)
		{
			MessageBox.Show("Es ist aktuell kein Projekt geöffnet, bitte erst ein Projekt öffnen.", "PDF-Assistent", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			Close();
			return;
		}

		//String zerlegen
		string[] arrayProjects = sProjects.Split(';');

		frmPDFAssistent frm = new frmPDFAssistent();

		//Schleife über alle markierten Projekte
		foreach (string sProject in arrayProjects)
		{
			//frmPDFAssistent frm = new frmPDFAssistent();
			frm.Tag = sProject; //Projekt in Tag speichern
			frm.txtDateiname.Text = System.IO.Path.GetFileNameWithoutExtension(sProject);
			frm.ShowDialog();
		}
		return;
	}


	//Assistent ohne Form ausführen
	//[DeclareAction("PDFAssistent_Start_Ohne_Dialog")]
	public void PDFAssistent_Start_Ohne_Dialog()
	{
		ActionCallingContext acc = new ActionCallingContext();
		CommandLineInterpreter oCLI = new CommandLineInterpreter();
		string sProjects = string.Empty;
		//Aktuell markierte Projekte ermitteln
		acc.AddParameter("TYPE", "PROJECTS");
		oCLI.Execute("selectionset", acc);
		acc.GetParameter("PROJECTS", ref sProjects);
		if (sProjects == string.Empty)
		{
			MessageBox.Show("Es ist aktuell kein Projekt geöffnet, bitte erst ein Projekt öffnen.", "PDF-Assistent", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			Close();
			return;
		}
		//String zerlegen
		string[] arrayProjects = sProjects.Split(';');
		//Schleife über alle markierten Projekte
		foreach (string sProject in arrayProjects)
		{
			Tag = sProject; //Projekt in Tag speichern
			txtDateiname.Text = System.IO.Path.GetFileNameWithoutExtension(sProject);
			cboAusgabeNach.SelectedIndex = 0;
			comboBoxAusgabe.SelectedIndex = 1;
			ReadSettings();
			SetPageFilter();
			PDFexport(sProject, txtSpeicherort.Text + txtDateiname.Text + @".pdf"); //frm.Tag als Träger für den Proejktnamen verwendet
			WriteSettings();
		}
		Close();
	}

	//Form wird geladen
	private void frmPDFAssistent_Load(object sender, EventArgs e)
	{
		//Einstellen
		cboAusgabeNach.SelectedIndex = 0;
		comboBoxAusgabe.SelectedIndex = 1;
		chkIstInProjektOrdner.CheckState = CheckState.Unchecked;
		txtProjektGespeichertInOrdner.Enabled = false;
		btnProjektOrdnerAuswahl.Enabled = false;
		ReadSettings();
		SetPageFilter();
	}

	//Button: Abbrechen
	private void btnAbbrechen_Click(object sender, System.EventArgs e)
	{
		this.Close();
	}

	//Button: OK
	private void btnOK_Click(object sender, System.EventArgs e)
	{
		if (txtDateiname.Text != string.Empty)
		{
			string sDateiName = txtSpeicherort.Text + txtDateiname.Text + @".pdf";
			if (File.Exists(sDateiName))
			{
				DialogResult result = MessageBox.Show("Die PDF-Datei existiert bereits, soll diese nun überschrieben werden?", "PDF Assistent", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
				if (result == DialogResult.Yes)
				{
					try
					{
						//Datei löschen
						File.Delete(sDateiName);
					}
					catch (IOException)
					{
						//Datei löschen hat nicht funktioniert, Datei wird vielleicht verwendet
						MessageBox.Show("Datei kann nicht erzeugt werden!\nDatei wird schon verwendet.", "PDF Assistent", MessageBoxButtons.OK, MessageBoxIcon.Error);
						return;
					}
				}
				else if (result == DialogResult.No)
				{
					return;
				}
			}

			//Auswertungen aktualisieren
			if (chkAuswertungenAktualisieren.Checked)
			{
				CommandLineInterpreter cmdLineItp = new CommandLineInterpreter();
				cmdLineItp.Execute("XFgEvaluateProjectAction");
			}

			//PDF Datei erzeugen
			PDFexport(Tag.ToString(), sDateiName); //frm.Tag als Träger für den Proejktnamen verwendet

			//Nach Eport PDF Anzeigen
			if (chkOpenPDF.Checked)
			{
				//PDF anzeigen mit dem Adobe Reader
				//MessageBox.Show(sDateiName, "Dateiname:" );
				Process.Start("AcroRd32.exe",sDateiName);
			}

			//Nach Export Ordner öffnen
			if (chkOrdnerÖffnen.Checked)
			{
				//Start Windows-Explorer mit Parameter
				System.Diagnostics.Process.Start("explorer", "/e," + txtSpeicherort.Text);
			}

		}
		//Einstellungen speichern speichern
		Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
		if (!oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.SaveSettings"))
		{
			oSettings.AddBoolSetting("USER.SCRIPTS.PDF_Assistent.SaveSettings",
				new bool[] { false },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetBoolSetting("USER.SCRIPTS.PDF_Assistent.SaveSettings", chkEinstellungSpeichern.Checked, 1); //1 = Visible = True
																												 //wenn Einstellungen speichern dann speichern
		if (chkEinstellungSpeichern.Checked)
		{
			WriteSettings();
		}
		this.Close();
	}

	//Ausgabe Nach hat sich geändert
	private void cboAusgabeNach_SelectedIndexChanged(object sender, EventArgs e)
	{
#if !DEBUG
        string sProjektOrdner = Tag.ToString(); //frm.Tag als Träger für den Proejktnamen verwendet
        sProjektOrdner = Path.ChangeExtension(sProjektOrdner, ".edb");
        string sDateiName = System.IO.Path.GetFileNameWithoutExtension(sProjektOrdner);
#else
		string sProjektOrdner = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
		string sDateiName = "TEST";
#endif

		string sAusgabeOrdner = sProjektOrdner;
		Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
		string sLastSchema = string.Empty;

		switch (cboAusgabeNach.SelectedIndex)
		{
			case 0: //in den Projekt-Ordner
				sAusgabeOrdner = sProjektOrdner;
				break;

			case 1: //Ausgabeverzeichnis aus Einstellungen: PDF-Export
#if !DEBUG
				if (oSettings.ExistSetting("USER.PDFExportGUI.SCHEMAS.LastUsed"))
				{
					sLastSchema = oSettings.GetStringSetting("USER.PDFExportGUI.SCHEMAS.LastUsed", 0);
				}
				if (oSettings.ExistSetting("USER.PDFExportGUI.SCHEMAS." + sLastSchema + ".Data.TargetPath"))
				{
					sAusgabeOrdner = oSettings.GetStringSetting("USER.PDFExportGUI.SCHEMAS." + sLastSchema + ".Data.TargetPath", 0);
					//falls mal der Dateiname aus dem Schema ausgelesen werden soll
					//sDateiName = oSettings.GetStringSetting("USER.PDFExportGUI.SCHEMAS." + sLastSchema + ".Data.FilenameFormat", 0);
				}
#endif
				break;

			case 2: //in den Ordner eine Ebene über dem Projekt-Ordner
				sAusgabeOrdner = sProjektOrdner.Substring(0, sProjektOrdner.LastIndexOf(@"\"));
				break;

			case 3: //nach "Dokumente"
				sAusgabeOrdner = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
				break;

			case 4: //auf den "Desktop"
				sAusgabeOrdner = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
				break;

			case 5: //gleicher Pfad wie Projekt nur auf anderes Laufwerk
				sAusgabeOrdner = sProjektOrdner.Substring(0, sProjektOrdner.LastIndexOf(@"\")); //in den Ordner eine Ebene über dem Projekt-Ordner
				sAusgabeOrdner = sAusgabeOrdner.Replace("C:", "D:");    //Hier anpassen, welche Laufwerksbuchstaben verwendet werden.
				break;

			case 6: //Ausgabeverzeichnis aus Einstellungen: Benutzerbezogener-Speicherort
#if !DEBUG
				if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.UserSpeicherOrt"))
				{
					sAusgabeOrdner = oSettings.GetStringSetting("USER.SCRIPTS.PDF_Assistent.UserSpeicherOrt", 0);
				}
#endif
				break;

			case 7: //Ausgabeverzeichnis aus Datei im Projektordner auslesen
				SaveProjectSpecific(sProjektOrdner, out sAusgabeOrdner);
				break;

			default:
				MessageBox.Show("Auswahl:default");
				break;
		}

		//Präfix hinzufügen
		if (txtPräfix.Text != string.Empty)
		{
			sDateiName = txtPräfix.Text + sDateiName;
		}

		//Datumsstempel hinzufügen
		if (chkDatumStempel.Checked && !chkZeitStempel.Checked)
		{
			if (chkStempelVorne.Checked)
			{
				sDateiName = DateTime.Now.ToString("yyyyMMdd") + "_" + sDateiName;
			}
			else
			{
				sDateiName += "_" + DateTime.Now.ToString("yyyyMMdd");
			}
		}

		//Zeitstempel hinzufügen
		if (chkZeitStempel.Checked && !chkDatumStempel.Checked)
		{
			if (chkStempelVorne.Checked)
			{
				sDateiName = DateTime.Now.ToString("HHmmss") + "-" + sDateiName;
			}
			else
			{
				sDateiName += "-" + DateTime.Now.ToString("HHmmss");
			}
		}

		//Datum + Zeitstempel hinzufügen
		if (chkZeitStempel.Checked && chkDatumStempel.Checked)
		{
			if (chkStempelVorne.Checked)
			{
				sDateiName = DateTime.Now.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("HHmmss") + "-" + sDateiName;
			}
			else
			{
				sDateiName += "_" + DateTime.Now.ToString("yyyyMMdd") + "-" + DateTime.Now.ToString("HHmmss");
			}
		}

		//Endet mit \ ?
		if (!sAusgabeOrdner.EndsWith(@"\"))
		{
			sAusgabeOrdner += @"\";
		}

		txtDateiname.Text = sDateiName;
		txtSpeicherort.Text = sAusgabeOrdner;
	}

	//Datumstempel zustand hat sich geändert
	private void chkDatumStempel_CheckedChanged(object sender, EventArgs e)
	{
		cboAusgabeNach_SelectedIndexChanged(sender, e);
	}

	//Zeitstempel zustand hat sich geändert
	private void chkZeitStempel_CheckedChanged(object sender, EventArgs e)
	{
		cboAusgabeNach_SelectedIndexChanged(sender, e);
	}

	//Stempel Vorne zustand hat sich geändert
	private void chkStempelVorne_CheckedChanged(object sender, EventArgs e)
	{
		cboAusgabeNach_SelectedIndexChanged(sender, e);
	}

	//PDF Export
	public void PDFexport(string sProjectName, string sZielDatei)
	{
		//Progressbar ein
		Eplan.EplApi.Base.Progress oProgress = new Eplan.EplApi.Base.Progress("SimpleProgress");
		oProgress.ShowImmediately();

		ActionCallingContext pdfContext = new ActionCallingContext();
		pdfContext.AddParameter("type", "PDFPROJECTSCHEME"); //PDFPROJECTSCHEME = Projekt im PDF-Format exportieren
		pdfContext.AddParameter("PROJECTNAME", sProjectName); //Projektname mit komplettem Pfad (optional)
															  //pdfContext.AddParameter("exportscheme", "NAME_SCHEMA"); //verwendetes Schema
		pdfContext.AddParameter("exportfile", sZielDatei); //Name export.Projekt, Vorgabewert: Projektname
		pdfContext.AddParameter("exportmodel", "0"); //0 = keine Modelle ausgeben
		switch (comboBoxAusgabe.SelectedIndex)
		{
			case 0: //0 = PDF wird farbig
				pdfContext.AddParameter("blackwhite", "0");
				break;
			case 1: //1 = PDF wird schwarz-weiss
				pdfContext.AddParameter("blackwhite", "1");
				break;
			case 2: //2 = PDF wird in Graustufen
				pdfContext.AddParameter("blackwhite", "2");
				break;
		}
		pdfContext.AddParameter("useprintmargins", "1"); //1 = Druckränder verwenden
		pdfContext.AddParameter("readonlyexport", "2"); //1 = PDF wird schreibgeschützt
		pdfContext.AddParameter("usesimplelink", "1"); //1 = einfache Sprungfunktion
		pdfContext.AddParameter("usezoomlevel", "1"); //Springen in Navigationsseiten
		pdfContext.AddParameter("fastwebview", "1"); //1 = schnelle Web-Anzeige
		pdfContext.AddParameter("zoomlevel", "1"); //wenn USEZOOMLEVEL auf 1 dann hier Zoomstufe in mm

		CommandLineInterpreter cmdLineItp = new CommandLineInterpreter();
		cmdLineItp.Execute("export", pdfContext);

		//'Progressbar aus
		oProgress.EndPart(true);

		return;
	}

	//Einstellungen speichern
	public void WriteSettings()
	{
		Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();

		//Einstellungen Seitenfilter im Schema setzen
		string sLastSchema = string.Empty;
		if (oSettings.ExistSetting("USER.PDFExportGUI.SCHEMAS.LastUsed"))
		{
			sLastSchema = oSettings.GetStringSetting("USER.PDFExportGUI.SCHEMAS.LastUsed", 0);
		}
		if (oSettings.ExistSetting("USER.PDFExportGUI.SCHEMAS." + sLastSchema + ".Data.ApplyPageFilter"))
		{
			oSettings.SetBoolSetting("USER.PDFExportGUI.SCHEMAS." + sLastSchema + ".Data.ApplyPageFilter", oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.PageFilterSchema", 1), 0);
		}

		//Einstellungen speichern
		if (!oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.SaveSettings"))
		{
			oSettings.AddBoolSetting("USER.SCRIPTS.PDF_Assistent.SaveSettings",
				new bool[] { false },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetBoolSetting("USER.SCRIPTS.PDF_Assistent.SaveSettings", chkEinstellungSpeichern.Checked, 1);

		//User-Speicherort
		if (!oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.UserSpeicherOrt"))
		{
			oSettings.AddStringSetting("USER.SCRIPTS.PDF_Assistent.UserSpeicherOrt",
			new string[] { },
			new string[] { },
			ISettings.CreationFlag.Insert);
		}
		oSettings.SetStringSetting("USER.SCRIPTS.PDF_Assistent.UserSpeicherOrt", txtSpeicherort.Text, 0);

		//PDF bei Projekt schließen
		if (!oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.ByProjectClose"))
		{
			oSettings.AddBoolSetting("USER.SCRIPTS.PDF_Assistent.ByProjectClose",
				new bool[] { false },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetBoolSetting("USER.SCRIPTS.PDF_Assistent.ByProjectClose", chkByProjectClose.Checked, 1);

		//PDF bei Eplan beenden
		if (!oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.ByEplanEnd"))
		{
			oSettings.AddBoolSetting("USER.SCRIPTS.PDF_Assistent.ByEplanEnd",
				new bool[] { false },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetBoolSetting("USER.SCRIPTS.PDF_Assistent.ByEplanEnd", chkByEplanEnd.Checked, 1);

		//Ausgabe Nach
		if (!oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.AusgabeNach"))
		{
			oSettings.AddNumericSetting("USER.SCRIPTS.PDF_Assistent.AusgabeNach",
				new int[] { 0 },
				new Range[] { new Range { FromValue = 0, ToValue = 32768 } },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetNumericSetting("USER.SCRIPTS.PDF_Assistent.AusgabeNach", cboAusgabeNach.SelectedIndex, 1);

		oSettings.SetBoolSetting("USER.SCRIPTS.PDF_Assistent.ByEplanEnd", chkByEplanEnd.Checked, 1);

		//Ausführen ohne Nachfrage
		if (!oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.OhneNachfrage"))
		{
			oSettings.AddBoolSetting("USER.SCRIPTS.PDF_Assistent.OhneNachfrage",
				new bool[] { false },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetBoolSetting("USER.SCRIPTS.PDF_Assistent.OhneNachfrage", chkOhneNachfrage.Checked, 1);

		//Präfix
		if (!oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.Präfix"))
		{
			oSettings.AddStringSetting("USER.SCRIPTS.PDF_Assistent.Präfix",
			new string[] { },
			new string[] { },
			ISettings.CreationFlag.Insert);
		}
		oSettings.SetStringSetting("USER.SCRIPTS.PDF_Assistent.Präfix", txtPräfix.Text, 0);

		//Datumsstempel
		if (!oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.DateStamp"))
		{
			oSettings.AddBoolSetting("USER.SCRIPTS.PDF_Assistent.DateStamp",
				new bool[] { false },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetBoolSetting("USER.SCRIPTS.PDF_Assistent.DateStamp", chkDatumStempel.Checked, 1);

		//Zeitstempel
		if (!oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.TimeStamp"))
		{
			oSettings.AddBoolSetting("USER.SCRIPTS.PDF_Assistent.TimeStamp",
				new bool[] { false },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetBoolSetting("USER.SCRIPTS.PDF_Assistent.TimeStamp", chkZeitStempel.Checked, 1);

		//Stempel Vorne
		if (!oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.StampBefore"))
		{
			oSettings.AddBoolSetting("USER.SCRIPTS.PDF_Assistent.StampBefore",
				new bool[] { false },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetBoolSetting("USER.SCRIPTS.PDF_Assistent.StampBefore", chkStempelVorne.Checked, 1);

		//Ausgabe in Farbe/Schwarz-Weiß/Graustufen/Schema
		if (!oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.AusgabeIn"))
		{
			oSettings.AddNumericSetting("USER.SCRIPTS.PDF_Assistent.AusgabeIn",
				new int[] { 0 },
				new Range[] { new Range { FromValue = 0, ToValue = 32768 } },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetNumericSetting("USER.SCRIPTS.PDF_Assistent.AusgabeIn", comboBoxAusgabe.SelectedIndex, 0);

		//Seitenfilter
		if (!oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.PageFilter"))
		{
			oSettings.AddBoolSetting("USER.SCRIPTS.PDF_Assistent.PageFilter",
				new bool[] { false },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetBoolSetting("USER.SCRIPTS.PDF_Assistent.PageFilter", chkPageFilter.Checked, 1);

		//Projekt in Ordner
		if (!oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.ProjectInOrdner"))
		{
			oSettings.AddBoolSetting("USER.SCRIPTS.PDF_Assistent.ProjectInOrdner",
				new bool[] { false },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetBoolSetting("USER.SCRIPTS.PDF_Assistent.ProjectInOrdner", chkIstInProjektOrdner.Checked, 1);

		if (!oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.ProjectInOrdnerName"))
		{
			oSettings.AddStringSetting("USER.SCRIPTS.PDF_Assistent.ProjectInOrdnerName",
			new string[] { },
			new string[] { },
			ISettings.CreationFlag.Insert);
		}
		oSettings.SetStringSetting("USER.SCRIPTS.PDF_Assistent.ProjectInOrdnerName", txtProjektGespeichertInOrdner.Text, 0);

		//Nach Export PDF anzeigen
		if (!oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.ViewExport"))
		{
			oSettings.AddBoolSetting("USER.SCRIPTS.PDF_Assistent.ViewExport",
				new bool[] { false },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetBoolSetting("USER.SCRIPTS.PDF_Assistent.ViewExport", chkOpenPDF.Checked, 0);

		//Nach Export Ordner öffnen
		if (!oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.OpenFolder"))
		{
			oSettings.AddBoolSetting("USER.SCRIPTS.PDF_Assistent.OpenFolder",
				new bool[] { false },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetBoolSetting("USER.SCRIPTS.PDF_Assistent.OpenFolder", chkOrdnerÖffnen.Checked, 0);

		//Auswertungen aktualisieren
		if (!oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.EvaluateProject"))
		{
			oSettings.AddBoolSetting("USER.SCRIPTS.PDF_Assistent.EvaluateProject",
				new bool[] { false },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetBoolSetting("USER.SCRIPTS.PDF_Assistent.EvaluateProject", chkAuswertungenAktualisieren.Checked, 0);

		//Wenn Projektbezogener-Speicherort ausgewählt ist
		if (cboAusgabeNach.SelectedIndex == 7)
		{
#if !DEBUG
        string sProjektOrdner = Tag.ToString(); //frm.Tag als Träger für den Proejktnamen verwendet
        sProjektOrdner = Path.ChangeExtension(sProjektOrdner, ".edb");
#else
			string sProjektOrdner = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
#endif
			string sPDFAssistentFile = Path.Combine(sProjektOrdner, "PDFAssistent.ini");
			if (File.Exists(sPDFAssistentFile))
			{
				File.Delete(sPDFAssistentFile);
				File.WriteAllText(sPDFAssistentFile, txtSpeicherort.Text, System.Text.Encoding.Default);
			}
			else
			{
				MessageBox.Show("Die Einstellung für den Speicherort ist noch nicht vorhanden.\nBitte den Speicherort nun auswählen", "PDF Assistent", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				OrdnerAuswahl();
				File.WriteAllText(sPDFAssistentFile, txtSpeicherort.Text, System.Text.Encoding.Default);
			}
		}
	}

	//Einstellungen einlesen
	public void ReadSettings()
	{
		Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();

		//Seitenfilter einstellung aus Schema auslesen und in Einstellung speichern
		string sLastSchema = string.Empty;
		if (oSettings.ExistSetting("USER.PDFExportGUI.SCHEMAS.LastUsed"))
		{
			sLastSchema = oSettings.GetStringSetting("USER.PDFExportGUI.SCHEMAS.LastUsed", 0);
		}
		if (oSettings.ExistSetting("USER.PDFExportGUI.SCHEMAS." + sLastSchema + ".Data.ApplyPageFilter"))
		{
			if (!oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.PageFilterSchema"))
			{
				oSettings.AddBoolSetting("USER.SCRIPTS.PDF_Assistent.PageFilterSchema",
					new bool[] { false },
					ISettings.CreationFlag.Insert);
			}
			oSettings.SetBoolSetting("USER.SCRIPTS.PDF_Assistent.PageFilterSchema", oSettings.GetBoolSetting("USER.PDFExportGUI.SCHEMAS." + sLastSchema + ".Data.ApplyPageFilter", 0), 1);
		}

		//ByProjectClose
		if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.ByProjectClose"))
		{
			chkByProjectClose.Checked = oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.ByProjectClose", 1);
		}

		//ByEplanEnd
		if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.ByEplanEnd"))
		{
			chkByEplanEnd.Checked = oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.ByEplanEnd", 1);
		}

		//Ausführen ohne Nachfrage
		if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.OhneNachfrage"))
		{
			chkOhneNachfrage.Checked = oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.OhneNachfrage", 1);
		}

		//Einstellungen speichern
		if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.SaveSettings"))
		{
			chkEinstellungSpeichern.Checked = oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.SaveSettings", 1);
		}

		//Präfix
		if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.Präfix"))
		{
			txtPräfix.Text = oSettings.GetStringSetting("USER.SCRIPTS.PDF_Assistent.Präfix", 0);
		}

		//Datumsstempel 
		if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.DateStamp"))
		{
			chkDatumStempel.Checked = oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.DateStamp", 1);
		}

		//Zeitstempel 
		if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.TimeStamp"))
		{
			chkZeitStempel.Checked = oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.TimeStamp", 1);
		}

		//Stempel Vorne
		if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.StampBefore"))
		{
			chkStempelVorne.Checked = oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.StampBefore", 1);
		}

		//Ausgabe farbig 
		if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.AusgabeIn"))
		{
			comboBoxAusgabe.SelectedIndex = oSettings.GetNumericSetting("USER.SCRIPTS.PDF_Assistent.AusgabeIn", 0);
		}

		//Seitenfilter 
		if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.PageFilter"))
		{
			chkPageFilter.Checked = oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.PageFilter", 1);
		}

		//Ausgabe Nach
		if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.AusgabeNach"))
		{
			cboAusgabeNach.SelectedIndex = oSettings.GetNumericSetting("USER.SCRIPTS.PDF_Assistent.AusgabeNach", 1);
		}

		//Projekt in Ordner
		if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.ProjectInOrdner"))
		{
			chkIstInProjektOrdner.Checked = oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.ProjectInOrdner", 1);
		}
		if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.ProjectInOrdnerName"))
		{
			txtProjektGespeichertInOrdner.Text = oSettings.GetStringSetting("USER.SCRIPTS.PDF_Assistent.ProjectInOrdnerName", 0);
		}

		//Anzeigen
		if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.ViewExport"))
		{
			chkOpenPDF.Checked = oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.ViewExport", 0);
		}

		//Ordner öffnen
		if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.OpenFolder"))
		{
			chkOrdnerÖffnen.Checked = oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.OpenFolder", 0);
		}

		//Auswertungen aktualisieren
		if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.EvaluateProject"))
		{
			chkAuswertungenAktualisieren.Checked = oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.EvaluateProject", 0);
		}

	}

	//Button: PDF Ordner auswählen
	private void btnOrdnerAuswahl_Click(object sender, EventArgs e)
	{
		OrdnerAuswahl();
	}

	//Button: Speichern
	private void btnSpeichern_Click(object sender, EventArgs e)
	{
		WriteSettings();
	}

	//XML-Reader
	private static string ReadXml(string filename, int ID)
	{
		string strLastVersion = "";
		XmlTextReader reader = new XmlTextReader(filename);
		while (reader.Read())
		{
			if (reader.HasAttributes)
			{
				while (reader.MoveToNextAttribute())
				{
					if (reader.Name == "id")
					{
						if (reader.Value == ID.ToString())
						{
							strLastVersion = reader.ReadString();
							reader.Close();
							return strLastVersion;
						}
					}
				}
			}
		}
		return strLastVersion;
	}

	//Button: Projekt Ordner auswählen
	private void btnProjektOrdnerAuswahl_Click(object sender, EventArgs e)
	{
		FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
		folderBrowserDialog1.SelectedPath = txtProjektGespeichertInOrdner.Text;
		folderBrowserDialog1.Description = "Wählen Sie den Ordner aus in dem das Projekt gespeichert sein muss:";
		DialogResult result = folderBrowserDialog1.ShowDialog();
		if (result == DialogResult.OK)
		{
			string sSpeicherort = folderBrowserDialog1.SelectedPath;
			if (!sSpeicherort.EndsWith(@"\"))
			{
				sSpeicherort += @"\";
			}
			txtProjektGespeichertInOrdner.Text = sSpeicherort;
		}
	}

	//Ist in Ordner hat sich geändert
	private void chkIstInProjektOrdner_CheckedChanged(object sender, EventArgs e)
	{
		if (chkIstInProjektOrdner.Checked)
		{
			txtProjektGespeichertInOrdner.Enabled = true;
			btnProjektOrdnerAuswahl.Enabled = true;
		}
		else
		{
			txtProjektGespeichertInOrdner.Enabled = false;
			btnProjektOrdnerAuswahl.Enabled = false;
		}
	}

	//Seitenfilter zustand hat sich geändert
	private void chkPageFilter_CheckedChanged(object sender, EventArgs e)
	{
		SetPageFilter();
	}

	//Seitenfilter setzen
	private void SetPageFilter()
	{
		//Letztes Schema für PDF-Export
		Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
		string sLastSchema = string.Empty;
		if (oSettings.ExistSetting("USER.PDFExportGUI.SCHEMAS.LastUsed"))
		{
			sLastSchema = oSettings.GetStringSetting("USER.PDFExportGUI.SCHEMAS.LastUsed", 0);
		}

		//Seitenfilter aktiv setzen
		if (oSettings.ExistSetting("USER.PDFExportGUI.SCHEMAS." + sLastSchema + ".Data.ApplyPageFilter"))
		{
			oSettings.SetBoolSetting("USER.PDFExportGUI.SCHEMAS." + sLastSchema + ".Data.ApplyPageFilter", chkPageFilter.Checked, 0); //1 = Visible = True
		}
	}

	//Kontextmenü Speicherort
	private void contextMenuSpeicherort_Click(object sender, EventArgs e)
	{
		//Start Windows-Explorer mit Parameter
		System.Diagnostics.Process.Start("explorer", "/e," + txtSpeicherort.Text);
	}

	//Präfix Text hat sich geändert
	private void txtPräfix_TextChanged(object sender, EventArgs e)
	{
		cboAusgabeNach_SelectedIndexChanged(sender, e);
	}

	//Ordner Auswahl
	private void OrdnerAuswahl()
	{
		FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
		folderBrowserDialog1.SelectedPath = txtSpeicherort.Text;
		folderBrowserDialog1.Description = "Wählen Sie den Zielordner aus in dem die PDF-Datei gespeichert werden soll:";
		DialogResult result = folderBrowserDialog1.ShowDialog();
		if (result == DialogResult.OK)
		{
			string sSpeicherort = folderBrowserDialog1.SelectedPath;
			if (!sSpeicherort.EndsWith(@"\"))
			{
				sSpeicherort += @"\";
			}
			txtSpeicherort.Text = sSpeicherort;
		}
	}

	//Projektbezogener-Speicherort
	private void SaveProjectSpecific(string sProjektOrdner, out string sAusgabeOrdner)
	{
		string sPDFAssistentFile = Path.Combine(sProjektOrdner, "PDFAssistent.ini");
		if (File.Exists(sPDFAssistentFile))
		{
			string value = File.ReadAllLines(sPDFAssistentFile, System.Text.Encoding.Default)[0];
			sAusgabeOrdner = value;
		}
		else
		{
			MessageBox.Show("Die Einstellung für den Speicherort ist noch nicht vorhanden.\nBitte den Speicherort nun auswählen", "PDF Assistent", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			OrdnerAuswahl();
			File.WriteAllText(sPDFAssistentFile, txtSpeicherort.Text, System.Text.Encoding.Default);
			sAusgabeOrdner = txtSpeicherort.Text;
		}
	}
}
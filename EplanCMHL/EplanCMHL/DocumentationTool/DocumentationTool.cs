// Documentation-Tool.cs
//
// Aufruf über einen neuen Menüpunkt "Dokumentations-Tool..."
// im Hauptmenü "Projekt"
//
// Copyright by Frank Schöneck, 2017-2018
//
// letzte Änderung:
// V2.0.0, 08.08.2017, Frank Schöneck,	Projektbeginn
// V2.1.0, 17.08.2017, Frank Schöneck,	Erweitert auf 20 Externe Dokumente bei Artikel (eingelagert)
// V2.2.0, 18.08.2017, Frank Schöneck,	Kontextmenüpunkte in Zielverzeichnis hinzugefügt
// V2.3.0, 25.08.2017, Frank Schöneck,	Beschreibung des Dokumentes, der Hersteller und die Artikelnummer wird jetzt angezeigt
//										Kontextmenü in Dokumentenliste hinzugefügt, darüber ist nun "Dokument öffnen" möglich
//										Kontextmenü Zielverzeichnis um Menüpunkt "Hersteller zur Verzeichnisstruktur hinzufügen" erweitert
//										Kontextmenü Zielverzeichnis um Menüpunkt "Artikelnummer zur Verzeichnisstruktur hinzufügen" erweitert
// V2.4.0, 06.09.2017, Frank Schöneck,	Die Oberfläche ist nun Größenveränderbar und die Position und Größe wird gespeichert
//										Button "Extras" hinzugefügt, zum Aufrufen des Kontextmenü Zielverzeichnis
//										Die Einstellungen des des Kontextmenü Zielverzeichnis werden nun gespeichert
//										Die Spalten können nun durch anklicken der Spaltenüberschrift sortiert werden
// V2.4.1, 19.12.2017, Frank Schöneck,	Fehlerbehandlung für Dateikopieren eingefügt.
//										Beim Zielverzeichnisnamen werden ungültige Zeichen automatisch entfernt
// V2.4.2, 24.09.2018, Frank Schöneck,	Darstellungsfehler im UI behoben.
// V2.5.0, 26.09.2018, Frank Schöneck,	Unterstützung von Pfadvariablen hinzugefügt.
// V2.6.0, 04.12.2018, Frank Schöneck,	Neues System mit Variablen zur Erzeugung der Ablagestruktur eingeführt.
// V3.0.0, 13.07.2022, Frank Schöneck,	Angepasst an die Ribbon-Technik der EPLAN Plattform 2022.
//
// für Eplan Electric P8, ab Plattform 2022 Update 2
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
using Eplan.EplApi.Base.SettingNode;
#endif

/*
On the other hand, some namespaces are not automatically added by EPLAN when
you load a script. Those have to be outside of the previous conditional compiler directive
*/

using System;
using System.IO;
using System.Xml;
using System.Collections;
using System.Windows.Forms;
using System.Linq;

public partial class frmDocumentationTool : System.Windows.Forms.Form
{
	private Button btnAbbrechen;
	private ListView listView;
	private ColumnHeader columnFileName;
	private Button btnKopieren;
	private TextBox txtZielverzeichnis;
	private Label label1;
	private Button btnOrdnerWählen;
	private Button btnOdnerÖffnen;
	private ContextMenuStrip contextMenuZielverzeichnis;
	private ToolStripMenuItem toolStripMenuItemProjektVerzeichnisstruktur;
	private ToolStripMenuItem toolStripMenuItemProjekteVerzeichnis;
	private ColumnHeader columnFileDescription;
	private ContextMenuStrip contextMenuListView;
	private ToolStripMenuItem toolStripMenuItemDocumentOpen;
	private ColumnHeader columnHersteller;
	private ColumnHeader columnArtikelnummer;
	private Button btnExtras;
	private ToolStripSeparator toolStripSeparator2;
	private ToolStripMenuItem toolStripMenuItemPfadvariabele;
	private ColumnHeader columnProductgroup;
	private TextBox txtAblageStruktur;
	private Label label2;
	private ContextMenuStrip contextMenuAblageStruktur;
	private ToolStripMenuItem toolStripMenuStrukturHersteller;
	private ToolStripMenuItem toolStripMenuStrukturProduktgruppe;
	private ToolStripMenuItem toolStripMenuStrukturArtikelnummer;
	private Button btnAblageStukturWählen;
	private ToolStripMenuItem toolStripMenuStrukturVerzeichnisebene;

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
			this.listView = new System.Windows.Forms.ListView();
			this.columnFileName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
			this.columnFileDescription = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
			this.columnHersteller = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
			this.columnArtikelnummer = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
			this.columnProductgroup = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
			this.contextMenuListView = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.toolStripMenuItemDocumentOpen = new System.Windows.Forms.ToolStripMenuItem();
			this.btnKopieren = new System.Windows.Forms.Button();
			this.txtZielverzeichnis = new System.Windows.Forms.TextBox();
			this.contextMenuZielverzeichnis = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.toolStripMenuItemProjekteVerzeichnis = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripMenuItemProjektVerzeichnisstruktur = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
			this.toolStripMenuItemPfadvariabele = new System.Windows.Forms.ToolStripMenuItem();
			this.label1 = new System.Windows.Forms.Label();
			this.btnOrdnerWählen = new System.Windows.Forms.Button();
			this.btnOdnerÖffnen = new System.Windows.Forms.Button();
			this.btnExtras = new System.Windows.Forms.Button();
			this.txtAblageStruktur = new System.Windows.Forms.TextBox();
			this.contextMenuAblageStruktur = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.toolStripMenuStrukturVerzeichnisebene = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripMenuStrukturHersteller = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripMenuStrukturProduktgruppe = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripMenuStrukturArtikelnummer = new System.Windows.Forms.ToolStripMenuItem();
			this.label2 = new System.Windows.Forms.Label();
			this.btnAblageStukturWählen = new System.Windows.Forms.Button();
			this.contextMenuListView.SuspendLayout();
			this.contextMenuZielverzeichnis.SuspendLayout();
			this.contextMenuAblageStruktur.SuspendLayout();
			this.SuspendLayout();
			// 
			// btnAbbrechen
			// 
			this.btnAbbrechen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAbbrechen.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnAbbrechen.Location = new System.Drawing.Point(594, 381);
			this.btnAbbrechen.Name = "btnAbbrechen";
			this.btnAbbrechen.Size = new System.Drawing.Size(120, 26);
			this.btnAbbrechen.TabIndex = 8;
			this.btnAbbrechen.Text = "Abbrechen";
			this.btnAbbrechen.UseVisualStyleBackColor = true;
			this.btnAbbrechen.Click += new System.EventHandler(this.btnAbbrechen_Click);
			// 
			// listView
			// 
			this.listView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.listView.AutoArrange = false;
			this.listView.CheckBoxes = true;
			this.listView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnFileName,
            this.columnFileDescription,
            this.columnHersteller,
            this.columnArtikelnummer,
            this.columnProductgroup});
			this.listView.ContextMenuStrip = this.contextMenuListView;
			this.listView.FullRowSelect = true;
			this.listView.HideSelection = false;
			this.listView.Location = new System.Drawing.Point(12, 12);
			this.listView.Name = "listView";
			this.listView.ShowGroups = false;
			this.listView.Size = new System.Drawing.Size(700, 254);
			this.listView.Sorting = System.Windows.Forms.SortOrder.Ascending;
			this.listView.TabIndex = 1;
			this.listView.UseCompatibleStateImageBehavior = false;
			this.listView.View = System.Windows.Forms.View.Details;
			this.listView.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.listView_ColumnClick);
			// 
			// columnFileName
			// 
			this.columnFileName.Text = "Dokument";
			this.columnFileName.Width = 540;
			// 
			// columnFileDescription
			// 
			this.columnFileDescription.Text = "Beschreibung";
			this.columnFileDescription.Width = 150;
			// 
			// columnHersteller
			// 
			this.columnHersteller.Text = "Hersteller";
			this.columnHersteller.Width = 150;
			// 
			// columnArtikelnummer
			// 
			this.columnArtikelnummer.Text = "Artikelnummer";
			this.columnArtikelnummer.Width = 150;
			// 
			// columnProductgroup
			// 
			this.columnProductgroup.Text = "Produktgruppe";
			this.columnProductgroup.Width = 150;
			// 
			// contextMenuListView
			// 
			this.contextMenuListView.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItemDocumentOpen});
			this.contextMenuListView.Name = "contextMenuListView";
			this.contextMenuListView.Size = new System.Drawing.Size(169, 26);
			this.contextMenuListView.Opening += new System.ComponentModel.CancelEventHandler(this.contextMenuListView_Opening);
			// 
			// toolStripMenuItemDocumentOpen
			// 
			this.toolStripMenuItemDocumentOpen.Name = "toolStripMenuItemDocumentOpen";
			this.toolStripMenuItemDocumentOpen.Size = new System.Drawing.Size(168, 22);
			this.toolStripMenuItemDocumentOpen.Text = "Dokument öffnen";
			this.toolStripMenuItemDocumentOpen.Click += new System.EventHandler(this.toolStripMenuItemDocumentOpen_Click);
			// 
			// btnKopieren
			// 
			this.btnKopieren.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnKopieren.Location = new System.Drawing.Point(457, 381);
			this.btnKopieren.Name = "btnKopieren";
			this.btnKopieren.Size = new System.Drawing.Size(120, 26);
			this.btnKopieren.TabIndex = 7;
			this.btnKopieren.Text = "&Kopieren";
			this.btnKopieren.UseVisualStyleBackColor = true;
			this.btnKopieren.Click += new System.EventHandler(this.btnKopieren_Click);
			// 
			// txtZielverzeichnis
			// 
			this.txtZielverzeichnis.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.txtZielverzeichnis.ContextMenuStrip = this.contextMenuZielverzeichnis;
			this.txtZielverzeichnis.Location = new System.Drawing.Point(12, 338);
			this.txtZielverzeichnis.Name = "txtZielverzeichnis";
			this.txtZielverzeichnis.Size = new System.Drawing.Size(669, 20);
			this.txtZielverzeichnis.TabIndex = 4;
			// 
			// contextMenuZielverzeichnis
			// 
			this.contextMenuZielverzeichnis.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItemProjekteVerzeichnis,
            this.toolStripMenuItemProjektVerzeichnisstruktur,
            this.toolStripSeparator2,
            this.toolStripMenuItemPfadvariabele});
			this.contextMenuZielverzeichnis.Name = "contextMenuZielverzeichnis";
			this.contextMenuZielverzeichnis.Size = new System.Drawing.Size(337, 76);
			// 
			// toolStripMenuItemProjekteVerzeichnis
			// 
			this.toolStripMenuItemProjekteVerzeichnis.Name = "toolStripMenuItemProjekteVerzeichnis";
			this.toolStripMenuItemProjekteVerzeichnis.Size = new System.Drawing.Size(336, 22);
			this.toolStripMenuItemProjekteVerzeichnis.Text = "Projekt-Verzeichnisstruktur hinzufügen";
			this.toolStripMenuItemProjekteVerzeichnis.Click += new System.EventHandler(this.toolStripMenuItemProjekteVerzeichnis_Click);
			// 
			// toolStripMenuItemProjektVerzeichnisstruktur
			// 
			this.toolStripMenuItemProjektVerzeichnisstruktur.Name = "toolStripMenuItemProjektVerzeichnisstruktur";
			this.toolStripMenuItemProjektVerzeichnisstruktur.Size = new System.Drawing.Size(336, 22);
			this.toolStripMenuItemProjektVerzeichnisstruktur.Text = "Komplette Projekt-Verzeichnisstruktur hinzufügen";
			this.toolStripMenuItemProjektVerzeichnisstruktur.Click += new System.EventHandler(this.toolStripMenuItemProjektVerzeichnisstruktur_Click);
			// 
			// toolStripSeparator2
			// 
			this.toolStripSeparator2.Name = "toolStripSeparator2";
			this.toolStripSeparator2.Size = new System.Drawing.Size(333, 6);
			// 
			// toolStripMenuItemPfadvariabele
			// 
			this.toolStripMenuItemPfadvariabele.Name = "toolStripMenuItemPfadvariabele";
			this.toolStripMenuItemPfadvariabele.Size = new System.Drawing.Size(336, 22);
			this.toolStripMenuItemPfadvariabele.Text = "Pfadvariable einfügen...";
			this.toolStripMenuItemPfadvariabele.Click += new System.EventHandler(this.toolStripMenuItemPfadvariabeleItem_Click);
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(9, 321);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(80, 13);
			this.label1.TabIndex = 4;
			this.label1.Text = "Zielverzeichnis:";
			// 
			// btnOrdnerWählen
			// 
			this.btnOrdnerWählen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnOrdnerWählen.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnOrdnerWählen.Location = new System.Drawing.Point(684, 335);
			this.btnOrdnerWählen.Name = "btnOrdnerWählen";
			this.btnOrdnerWählen.Size = new System.Drawing.Size(28, 24);
			this.btnOrdnerWählen.TabIndex = 5;
			this.btnOrdnerWählen.Text = "...";
			this.btnOrdnerWählen.UseVisualStyleBackColor = true;
			this.btnOrdnerWählen.Click += new System.EventHandler(this.btnOrdnerWählen_Click);
			// 
			// btnOdnerÖffnen
			// 
			this.btnOdnerÖffnen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnOdnerÖffnen.Location = new System.Drawing.Point(12, 381);
			this.btnOdnerÖffnen.Name = "btnOdnerÖffnen";
			this.btnOdnerÖffnen.Size = new System.Drawing.Size(120, 26);
			this.btnOdnerÖffnen.TabIndex = 9;
			this.btnOdnerÖffnen.Text = "Verzeichnis &öffnen";
			this.btnOdnerÖffnen.UseVisualStyleBackColor = true;
			this.btnOdnerÖffnen.Click += new System.EventHandler(this.btnOdnerÖffnen_Click);
			// 
			// btnExtras
			// 
			this.btnExtras.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExtras.Location = new System.Drawing.Point(318, 381);
			this.btnExtras.Name = "btnExtras";
			this.btnExtras.Size = new System.Drawing.Size(120, 26);
			this.btnExtras.TabIndex = 6;
			this.btnExtras.Text = "E&xtras";
			this.btnExtras.UseVisualStyleBackColor = true;
			this.btnExtras.Click += new System.EventHandler(this.btnExtras_Click);
			// 
			// txtAblageStruktur
			// 
			this.txtAblageStruktur.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.txtAblageStruktur.ContextMenuStrip = this.contextMenuAblageStruktur;
			this.txtAblageStruktur.Location = new System.Drawing.Point(12, 298);
			this.txtAblageStruktur.Name = "txtAblageStruktur";
			this.txtAblageStruktur.Size = new System.Drawing.Size(669, 20);
			this.txtAblageStruktur.TabIndex = 2;
			// 
			// contextMenuAblageStruktur
			// 
			this.contextMenuAblageStruktur.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuStrukturVerzeichnisebene,
            this.toolStripMenuStrukturHersteller,
            this.toolStripMenuStrukturProduktgruppe,
            this.toolStripMenuStrukturArtikelnummer});
			this.contextMenuAblageStruktur.Name = "contextMenuZielverzeichnis";
			this.contextMenuAblageStruktur.Size = new System.Drawing.Size(270, 92);
			// 
			// toolStripMenuStrukturVerzeichnisebene
			// 
			this.toolStripMenuStrukturVerzeichnisebene.Name = "toolStripMenuStrukturVerzeichnisebene";
			this.toolStripMenuStrukturVerzeichnisebene.Size = new System.Drawing.Size(269, 22);
			this.toolStripMenuStrukturVerzeichnisebene.Text = "Verzeichnisebene \'\\\'";
			this.toolStripMenuStrukturVerzeichnisebene.Click += new System.EventHandler(this.toolStripMenuStrukturVerzeichnisebene_Click);
			// 
			// toolStripMenuStrukturHersteller
			// 
			this.toolStripMenuStrukturHersteller.Name = "toolStripMenuStrukturHersteller";
			this.toolStripMenuStrukturHersteller.Size = new System.Drawing.Size(269, 22);
			this.toolStripMenuStrukturHersteller.Text = "Hersteller \'$(MANUFACTURER)\'";
			this.toolStripMenuStrukturHersteller.Click += new System.EventHandler(this.toolStripMenuStrukturHersteller_Click);
			// 
			// toolStripMenuStrukturProduktgruppe
			// 
			this.toolStripMenuStrukturProduktgruppe.Name = "toolStripMenuStrukturProduktgruppe";
			this.toolStripMenuStrukturProduktgruppe.Size = new System.Drawing.Size(269, 22);
			this.toolStripMenuStrukturProduktgruppe.Text = "Produktgruppe \'$(PRODUCTGROUP)\'";
			this.toolStripMenuStrukturProduktgruppe.Click += new System.EventHandler(this.toolStripMenuStrukturProduktgruppe_Click);
			// 
			// toolStripMenuStrukturArtikelnummer
			// 
			this.toolStripMenuStrukturArtikelnummer.Name = "toolStripMenuStrukturArtikelnummer";
			this.toolStripMenuStrukturArtikelnummer.Size = new System.Drawing.Size(269, 22);
			this.toolStripMenuStrukturArtikelnummer.Text = "Artikelnummer \'$(PARTNR)\'";
			this.toolStripMenuStrukturArtikelnummer.Click += new System.EventHandler(this.toolStripMenuStrukturArtikelnummer_Click);
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(9, 281);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(78, 13);
			this.label2.TabIndex = 4;
			this.label2.Text = "Ablagestruktur:";
			// 
			// btnAblageStukturWählen
			// 
			this.btnAblageStukturWählen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAblageStukturWählen.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnAblageStukturWählen.Location = new System.Drawing.Point(684, 294);
			this.btnAblageStukturWählen.Name = "btnAblageStukturWählen";
			this.btnAblageStukturWählen.Size = new System.Drawing.Size(28, 24);
			this.btnAblageStukturWählen.TabIndex = 3;
			this.btnAblageStukturWählen.Text = "...";
			this.btnAblageStukturWählen.UseVisualStyleBackColor = true;
			this.btnAblageStukturWählen.Click += new System.EventHandler(this.btnAblageStukturWählen_Click);
			// 
			// frmDocumentationTool
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.CancelButton = this.btnAbbrechen;
			this.ClientSize = new System.Drawing.Size(726, 419);
			this.Controls.Add(this.btnExtras);
			this.Controls.Add(this.btnOdnerÖffnen);
			this.Controls.Add(this.btnAblageStukturWählen);
			this.Controls.Add(this.btnOrdnerWählen);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtAblageStruktur);
			this.Controls.Add(this.txtZielverzeichnis);
			this.Controls.Add(this.btnKopieren);
			this.Controls.Add(this.listView);
			this.Controls.Add(this.btnAbbrechen);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmDocumentationTool";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Documentation-Tool";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmDocumentationTool_FormClosing);
			this.Load += new System.EventHandler(this.frmDocumentationTool_Load);
			this.contextMenuListView.ResumeLayout(false);
			this.contextMenuZielverzeichnis.ResumeLayout(false);
			this.contextMenuAblageStruktur.ResumeLayout(false);
			this.ResumeLayout(false);
			this.PerformLayout();

	}

	public frmDocumentationTool()
	{
		InitializeComponent();
	}

	#endregion

	private int sortColumn = -1;

	//RibbonBar Einträge dfinieren
	//string m_TabName = "Werkzeuge";
	//string m_commandGroupName = "Erweiterungen";
//	string m_commandName = "Dokumentations-Tool";

	// //Skript wird geladen
	// [DeclareRegister]
	// public void OnRegisterScript()
	// {
	// 	var newTab = new Eplan.EplApi.Gui.RibbonBar().Tabs.FirstOrDefault(item => item.Name == m_TabName);
	// 	if (newTab == null) //Tab noch nicht vorhanden, dann neu erzeugen
	// 	{
	// 		newTab = new Eplan.EplApi.Gui.RibbonBar().AddTab(m_TabName);
	// 	}
	// 	var commandGroup = newTab.CommandGroups.FirstOrDefault(item => item.Name == m_commandGroupName);
	// 	if (commandGroup == null) //CommandGroup noch nicht vorhanden, dann neu erzeugen
	// 	{
	// 		commandGroup = newTab.AddCommandGroup(m_commandGroupName);
	// 	}
	// 	Eplan.EplApi.Gui.RibbonIcon ribbonIcon = new Eplan.EplApi.Gui.RibbonIcon(Eplan.EplApi.Gui.CommandIcon.TaskList); //Icon festlegen
	// 	commandGroup.AddCommand(m_commandName, "Documentation_Tool_Start", m_commandName, "Externe Dokumente ermitteln und kopieren", ribbonIcon);
	// }

	// //Skript wird entladen
	// [DeclareUnregister]
	// public void OnUnRegisterScript()
	// {
	// 	//Einstellungen entfernen
	// 	SettingsDelete();

	// 	//Command entfernen
	// 	var vTab = new Eplan.EplApi.Gui.RibbonBar().Tabs.FirstOrDefault(item => item.Name == m_TabName);
	// 	if (vTab != null)
	// 	{
	// 		var commandGroup = vTab.CommandGroups.FirstOrDefault(item => item.Name == m_commandGroupName);
	// 		if (commandGroup != null)
	// 		{
	// 			var command = commandGroup.Commands.Values.FirstOrDefault(item => item.Text == m_commandName);
	// 			if (command != null)
	// 			{
	// 				command.Remove();
	// 			}
	// 			//Wenn CommandGroup leer ist diese auch entfernen
	// 			if (commandGroup.Commands.Count == 0)
	// 			{
	// 				commandGroup.Remove();
	// 			}
	// 		}
	// 		//Wenn Tab leer ist dieses auch entfernen
	// 		if (vTab.Commands.Count == 0)
	// 		{
	// 			vTab.Remove();
	// 		}
	// 	}
	// }

	[Start]
	//[DeclareAction("Documentation_Tool_Start")]
	public void Function()
	{
		frmDocumentationTool frm = new frmDocumentationTool();
		frm.ShowDialog();

		return;
	}

	//Form_Load
	private void frmDocumentationTool_Load(object sender, System.EventArgs e)
	{
		//Position und Größe aus Settings lesen
#if !DEBUG
		Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
		if (oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Top"))
		{
			this.Top = oSettings.GetNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Top", 0);
		}
		if (oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Left"))
		{
			this.Left = oSettings.GetNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Left", 0);
		}
		if (oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Height"))
		{
			this.Height = oSettings.GetNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Height", 0);
		}
		if (oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Width"))
		{
			this.Width = oSettings.GetNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Width", 0);
		}
		if (oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.txtAblageStruktur"))
		{
			this.txtAblageStruktur.Text = oSettings.GetStringSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.txtAblageStruktur", 0);
		}
#endif

		//Titelzeile anpassen
		string sProjekt = string.Empty;
#if DEBUG
		sProjekt = "DEBUG";
#else
		CommandLineInterpreter cmdLineItp = new CommandLineInterpreter();
		ActionCallingContext ProjektContext = new ActionCallingContext();
		ProjektContext.AddParameter("TYPE", "PROJECT");
		cmdLineItp.Execute("selectionset", ProjektContext);
		ProjektContext.GetParameter("PROJECT", ref sProjekt);
		sProjekt = Path.GetFileNameWithoutExtension(sProjekt); //Projektname Pfad und ohne .elk
		if (sProjekt == string.Empty)
		{
			Decider eDecision = new Decider();
			EnumDecisionReturn eAnswer = eDecision.Decide(EnumDecisionType.eOkDecision, "Es ist kein Projekt ausgewählt.", "Documentation-Tool", EnumDecisionReturn.eOK, EnumDecisionReturn.eOK);
			if (eAnswer == EnumDecisionReturn.eOK)
			{
				Close();
				return;
			}
		}
#endif
		Text = Text + " - " + sProjekt;

		//Button Extras Text festlegen
		//btnExtras.Text = "            Extras           ▾"; // ▾ ▼

		//Zielverzeichnis vorbelegen
#if DEBUG
		txtZielverzeichnis.Text = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\Test";
#else
		txtZielverzeichnis.Text = "$(DOC)";
		//letztes Zielverzeichnis aus Settings lesen
		if (oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.txtZielverzeichnis"))
		{
			this.txtZielverzeichnis.Text = oSettings.GetStringSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.txtZielverzeichnis", 0);
		}
#endif

		// Temporären Dateinamen festlegen
#if DEBUG
		string sTempFile = Path.Combine(Application.StartupPath, "tmp_Projekt_Export.epj");
#else
		string sTempFile = Path.Combine(PathMap.SubstitutePath(@"$(TMP)"), "tmpDocumentationTool.epj");
#endif

		//Projekt exportieren
#if !DEBUG
		PXFexport(sTempFile);
#endif

		//PXF Datei einlesen und in Listview schreiben
		PXFeinlesen(sTempFile);

		//PXF Datei wieder löschen
#if !DEBUG
		File.Delete(sTempFile);
#endif
	}

	//Form_Close 
	private void frmDocumentationTool_FormClosing(object sender, FormClosingEventArgs e)
	{
		//Einstellungen speichern
		SettingsWrite();
	}

	//Gesamtes Projekt als PXF ausgeben
	public void PXFexport(string sFilename)
	{
		//Progressbar ein
		Progress oProgress = new Progress("SimpleProgress");
		oProgress.SetTitle("Documentation-Tool");
		oProgress.SetActionText("Projektdaten exportieren");
		oProgress.ShowImmediately();
		Application.DoEvents(); // Screen aktualisieren

		ActionCallingContext pxfContext = new ActionCallingContext();
		pxfContext.AddParameter("TYPE", "PXFPROJECT");
		pxfContext.AddParameter("EXPORTFILE", sFilename);
		pxfContext.AddParameter("EXPORTMASTERDATA", "0"); //Stammdaten mit exportieren (Standard = 1(Ja)
		pxfContext.AddParameter("EXPORTCONNECTIONS", "0"); //Verbindungen mit exportieren (Standard = 0(Nein)

		CommandLineInterpreter cmdLineItp = new CommandLineInterpreter();
		cmdLineItp.Execute("export", pxfContext);

		//Progressbar aus
		oProgress.EndPart(true);

		return;
	}

	//PXF Datei einlesen
	private void PXFeinlesen(string sFileName)
	{
		//Progressbar ein
#if !DEBUG
		Progress oProgress = new Progress("SimpleProgress");
		oProgress.SetTitle("Documentation-Tool");
		//oProgress.SetActionText("Projektdaten durchsuchen");
		oProgress.BeginPart(50,"Projektdaten durchsuchen");
		oProgress.ShowImmediately();
#endif

		//MessageBox.Show("XML.Reader :" + sFileName);
		ListViewItem objListViewItem = new ListViewItem();

		//Wir benötigen einen XmlReader für das Auslesen der XML-Datei 
		XmlTextReader XMLReader = new XmlTextReader(sFileName);

		//Es folgt das Auslesen der XML-Datei 
		while (XMLReader.Read()) //Es sind noch Daten vorhanden 
		{
			//Alle Attribute (Name-Wert-Paare) abarbeiten 
			if (XMLReader.AttributeCount > 0)
			{
				string sArtikelnummer = string.Empty;
				string sHersteller = string.Empty;
				string sProduktgruppe = string.Empty;
				//Es sind noch weitere Attribute vorhanden 
				while (XMLReader.MoveToNextAttribute()) //nächstes
				{
					if (XMLReader.Name == "P22001") // Artikel (eingelagert) Artikelnummer)
					{
						sArtikelnummer = XMLReader.Value;
					}
					if (XMLReader.Name == "P22007") // Artikel (eingelagert) Hersteller)
					{
						sHersteller = XMLReader.Value;
					}
					if (XMLReader.Name == "P22041") // Artikel (eingelagert) Produktgruppe)
					{
						sProduktgruppe = TranslateProductgroup(XMLReader.Value);
					}
					if (
						XMLReader.Name == "A2082" || // Hyperlink Dokument
													 //XMLReader.Name == "P11058" || // Fremddokument
						XMLReader.Name == "P22149" || // Artikel (eingelagert) Externes Dokument 1
						XMLReader.Name == "P22150" || // Artikel (eingelagert) Externes Dokument 2
						XMLReader.Name == "P22151" || // Artikel (eingelagert) Externes Dokument 3
						XMLReader.Name == "P22235" || // Artikel (eingelagert) Externes Dokument 4
						XMLReader.Name == "P22236" || // Artikel (eingelagert) Externes Dokument 5
						XMLReader.Name == "P22237" || // Artikel (eingelagert) Externes Dokument 6
						XMLReader.Name == "P22238" || // Artikel (eingelagert) Externes Dokument 7
						XMLReader.Name == "P22239" || // Artikel (eingelagert) Externes Dokument 8
						XMLReader.Name == "P22240" || // Artikel (eingelagert) Externes Dokument 9
						XMLReader.Name == "P22241" || // Artikel (eingelagert) Externes Dokument 10
						XMLReader.Name == "P22242" || // Artikel (eingelagert) Externes Dokument 11
						XMLReader.Name == "P22243" || // Artikel (eingelagert) Externes Dokument 12
						XMLReader.Name == "P22244" || // Artikel (eingelagert) Externes Dokument 13
						XMLReader.Name == "P22245" || // Artikel (eingelagert) Externes Dokument 14
						XMLReader.Name == "P22246" || // Artikel (eingelagert) Externes Dokument 15
						XMLReader.Name == "P22247" || // Artikel (eingelagert) Externes Dokument 16
						XMLReader.Name == "P22248" || // Artikel (eingelagert) Externes Dokument 17
						XMLReader.Name == "P22249" || // Artikel (eingelagert) Externes Dokument 18
						XMLReader.Name == "P22250" || // Artikel (eingelagert) Externes Dokument 19
						XMLReader.Name == "P22251" // Artikel (eingelagert) Externes Dokument 20
						)
					{

						string[] sValue = XMLReader.Value.Split('\n');

						string sDateiname = string.Empty;
						string sDateiBeschreibung = string.Empty;
						sDateiname = sValue[0];

						//Überprüfen ob Beschreibung vorhanden ist
						if (sValue.Length == 2)
						{
							if (sValue[1] != null && sValue[1] != string.Empty)
							{
								sDateiBeschreibung = sValue[1];
#if !DEBUG
								MultiLangString mlstrDateiBeschreibung = new MultiLangString();
								//Nur die deutsche Übersetzung verwenden
								//String in MultiLanguages wandeln
								mlstrDateiBeschreibung.SetAsString(sDateiBeschreibung);
								//Daraus nur die Deutsche übersetzung
								sDateiBeschreibung = mlstrDateiBeschreibung.GetString(ISOCode.Language.L_de_DE);
								//Wenn es keine Deutsche gibt, dann die unbestimmte
								if (sDateiBeschreibung == "")
								{
									sDateiBeschreibung = mlstrDateiBeschreibung.GetString(ISOCode.Language.L___);
								}
#endif
							}
						}

#if !DEBUG
						sDateiname = PathMap.SubstitutePath(sDateiname);
#endif
						//keine Dokumente die mit HTTP anfangen bearbeiten
						if (!sDateiname.StartsWith("http"))
						{
							objListViewItem = new ListViewItem();
							objListViewItem.Name = sDateiname; // Name muß gesetzt werden damit ContainsKey funktioniert
							objListViewItem.Text = sDateiname;
							objListViewItem.SubItems.Add(sDateiBeschreibung); //Datei Beschreibung
							objListViewItem.SubItems.Add(sHersteller); //Hersteller
							objListViewItem.SubItems.Add(sArtikelnummer); //Artikelnummer
							objListViewItem.SubItems.Add(sProduktgruppe); //Produktgruppe
							objListViewItem.Checked = true;

							//Prüfen ob Datei vorhanden
							if (!File.Exists(sDateiname))
							{
								objListViewItem.Checked = false;
								objListViewItem.ForeColor = System.Drawing.Color.Gray;
							}

							//Eintrag in Listview, wenn nicht schon vorhanden
							if (!listView.Items.ContainsKey(sDateiname))
							{
								listView.Items.Add(objListViewItem);
							}
						}
					}
				}
			}
		}

		//XMLTextReader schließen
		XMLReader.Close();

		//Spaltenbreite automatisch an Inhaltsbreite anpassen
		//listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);

		//Progressbar aus
#if !DEBUG
		oProgress.EndPart(true);
#endif

		return;
	}

	//Dokument kopieren
	public void DocumentCopy(string sDokument, string sZiel)
	{
		//gibt es das Dokument auch?
		if (File.Exists(sDokument))
		{
			//Dateinamen ermitteln
			string sDateiname = Path.GetFileName(sDokument);

			try
			{
				//Dokument kopieren, mit überschreiben
				File.Copy(sDokument, Path.Combine(sZiel, sDateiname), true);
			}
			catch (Exception exc)
			{
				String strMessage = exc.Message;
				MessageBox.Show("Exception: " + strMessage + "\n\n" + sZiel + " -- " + sDateiname, "Documentation-Tool, DocumentCopy", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
		else
		{
			//Hinweis, Datei gibt es nicht
#if DEBUG
			MessageBox.Show("Das Dokument [" + sDokument + "] ist nicht auf dem Datenträger vorhanden!", "Documentation-Tool", MessageBoxButtons.OK, MessageBoxIcon.Information);
#else
			Decider eDecision = new Decider();
			eDecision.Decide(EnumDecisionType.eOkDecision, "Das Dokument [" + sDokument + "] ist nicht auf dem Datenträger vorhanden!", "Documentation-Tool", EnumDecisionReturn.eOK, EnumDecisionReturn.eOK);

#endif
		}
		return;
	}

	//Ordner auswählen
	public string OrdnerAuswählen(string InitialDirectory)
	{
		var folderBrowser = new FolderBrowserDialog();
		folderBrowser.Description = "Bitte wählen Sie einen Ordner aus";
		folderBrowser.SelectedPath = InitialDirectory;
		folderBrowser.ShowNewFolderButton = true;
		DialogResult result = folderBrowser.ShowDialog();
		if (result == DialogResult.OK)
		{
			return folderBrowser.SelectedPath;
		}
		else
		{
			return string.Empty;
		}
	}

	//Button: Abbrechen
	private void btnAbbrechen_Click(object sender, System.EventArgs e)
	{
		Close();
	}

	//Button: Zielverzeichniss auswählen
	private void btnOrdnerWählen_Click(object sender, EventArgs e)
	{
		string sZielverzeichnis = txtZielverzeichnis.Text;
		sZielverzeichnis = OrdnerAuswählen(PathMap.SubstitutePath(txtZielverzeichnis.Text));
		if (sZielverzeichnis != string.Empty)
		{
			txtZielverzeichnis.Text = sZielverzeichnis;
		}
	}

	//Button: Zielverzeichnis im Explorer öffnen
	private void btnOdnerÖffnen_Click(object sender, EventArgs e)
	{
		string sTargetPath = PathMap.SubstitutePath(txtZielverzeichnis.Text);

		//gibt es das Ziel auch?
		if (!Directory.Exists(sTargetPath))
		{
			//Hinweis, Ziel gibt es nicht
#if DEBUG
			MessageBox.Show("Das Zielverzeichnis [" + sTargetPath + "] ist nicht auf dem Datenträger vorhanden!", "Documentation-Tool", MessageBoxButtons.OK, MessageBoxIcon.Information);
#else
			new Decider().Decide(
				EnumDecisionType.eOkDecision, // type
				"Das Zielverzeichnis [" + sTargetPath + "] ist nicht auf dem Datenträger vorhanden!",
				"Documentation-Tool",
				EnumDecisionReturn.eOK, // selected Answer
				EnumDecisionReturn.eOK); // Answer if quite-mode on
#endif
		}
		//Es gibt dasZiel
		if (Directory.Exists(sTargetPath))
		{
			//Start Windows-Explorer mit Parameter
			System.Diagnostics.Process.Start("explorer", "/e," + sTargetPath);
		}
		return;
	}

	//Button: Dokumente kopieren
	private void btnKopieren_Click(object sender, EventArgs e)
	{
		string sTargetPath = PathMap.SubstitutePath(txtZielverzeichnis.Text);

		//gibt es das Ziel auch?
		if (!Directory.Exists(sTargetPath))
		{
			//Hinweis, Ziel gibt es nicht, anlegen?
#if DEBUG
			DialogResult dialogResult = MessageBox.Show("Das Zielverzeichnis [" + sTargetPath + "] ist nicht auf dem Datenträger vorhanden!\n\nSoll dieses nun angelegt werden?", "Documentation-Tool", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
			if (dialogResult == DialogResult.Yes)
			{
				Directory.CreateDirectory(sTargetPath);
			}
			else if (dialogResult == DialogResult.No)
			{
				return;
			}
#else
			Decider decider = new Decider();
			EnumDecisionReturn decision = decider.Decide(
				EnumDecisionType.eYesNoDecision, // type
				"Das Zielverzeichnis [" + sTargetPath + "] ist nicht auf dem Datenträger vorhanden!\n\nSoll dieses nun angelegt werden?",
				"Documentation-Tool",
				EnumDecisionReturn.eYES, // selected Answer
				EnumDecisionReturn.eYES); // Answer if quite-mode on

			if (decision == EnumDecisionReturn.eYES)
			{
				Directory.CreateDirectory(sTargetPath);
			}
			else if (decision == EnumDecisionReturn.eNO)
			{
				return;
			}
#endif
		}

		//Es gibt dasZiel
		if (Directory.Exists(sTargetPath))
		{
			ListView.CheckedListViewItemCollection checkedItems = listView.CheckedItems;
#if !DEBUG
		Progress oProgress = new Progress("SimpleProgress");
		oProgress.ShowImmediately();
		oProgress.SetAllowCancel(true);
		oProgress.SetTitle("Documentation-Tool");
		int nActionsPercent = 100 / checkedItems.Count;
#endif
			foreach (ListViewItem item in checkedItems)
			{
#if !DEBUG
			if (!oProgress.Canceled())
			{
				oProgress.BeginPart(nActionsPercent, "Kopiere: " + item.Text);
#endif
				string sTargetPathCopy = string.Empty;
				string sTemp = string.Empty;

				try
				{
					sTargetPathCopy = sTargetPath + txtAblageStruktur.Text;

					//Hersteller hinzufügen
					sTemp = item.SubItems[2].Text;
					sTemp = RemoveIlegaleCharackter(sTemp); //Verzeichnisnamen von ungültigen Zeichen bereinigen
					sTargetPathCopy = sTargetPathCopy.Replace("$(MANUFACTURER)", sTemp);

					//Produktgruppe hinzufügen
					sTemp = item.SubItems[4].Text;
					sTemp = RemoveIlegaleCharackter(sTemp); //Verzeichnisnamen von ungültigen Zeichen bereinigen
					sTargetPathCopy = sTargetPathCopy.Replace("$(PRODUCTGROUP)", sTemp);

					//Artikelnummer hinzufügen
					sTemp = item.SubItems[3].Text;
					sTemp = RemoveIlegaleCharackter(sTemp); //Verzeichnisnamen von ungültigen Zeichen bereinigen
					sTargetPathCopy = sTargetPathCopy.Replace("$(PARTNR)", sTemp);

					//Falls es das Verzeichnis nicht gibt erst anlegen dann kopieren
					if (!Directory.Exists(sTargetPathCopy))
					{
						Directory.CreateDirectory(sTargetPathCopy);
					}
					DocumentCopy(item.Text, sTargetPathCopy);
				}
				catch (Exception exc)
				{
					String strMessage = exc.Message;
					MessageBox.Show("Exception: " + strMessage + "\n\n" + sTargetPathCopy, "Documentation-Tool, btnKopieren", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}


#if !DEBUG
			oProgress.EndPart();
			}
#endif
			}
#if !DEBUG
		oProgress.EndPart(true);
#endif
			Close();
			return;
		}
	}

	//Button: AblageStukturWählen (Kontextmenü AblageStuktur anzeigen)
	private void btnAblageStukturWählen_Click(object sender, EventArgs e)
	{
		contextMenuAblageStruktur.Show(btnAblageStukturWählen, 0, 0 - (contextMenuAblageStruktur.Height - btnAblageStukturWählen.Height));
	}

	//Button: Extras (Kontextmenü Zielverzeichnis anzeigen)
	private void btnExtras_Click(object sender, EventArgs e)
	{
		contextMenuZielverzeichnis.Show(btnExtras, 0, 0 - (contextMenuZielverzeichnis.Height - btnExtras.Height));
	}

	//Listview: Spalten angeklickt
	private void listView_ColumnClick(object sender, ColumnClickEventArgs e)
	{
		// Determine whether the column is the same as the last column clicked.
		if (e.Column != sortColumn)
		{
			// Set the sort column to the new column.
			sortColumn = e.Column;
			// Set the sort order to ascending by default.
			listView.Sorting = SortOrder.Ascending;
		}
		else
		{
			// Determine what the last sort order was and change it.
			if (listView.Sorting == SortOrder.Ascending)
				listView.Sorting = SortOrder.Descending;
			else
				listView.Sorting = SortOrder.Ascending;
		}

		// Call the sort method to manually sort.
		listView.Sort();
		// Set the ListViewItemSorter property to a new ListViewItemComparer
		// object.
		this.listView.ListViewItemSorter = new ListViewItemComparer(e.Column, listView.Sorting);
	}

	//Kontextmenü 'Projekt-Verzeichnisstruktur hinzufügen'
	private void toolStripMenuItemProjekteVerzeichnis_Click(object sender, EventArgs e)
	{
		string sProjectePath = string.Empty;
		string sSelectedProjectPath = string.Empty;
#if DEBUG
		sProjectePath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
		sSelectedProjectPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\Test\Test.edb";
#else
		sProjectePath = PathMap.SubstitutePath(@"$(MD_PROJECTS)");
		sSelectedProjectPath = PathMap.SubstitutePath(@"$(P)");
#endif
		if (sProjectePath != string.Empty && sSelectedProjectPath != string.Empty)
		{

			string sTemp = string.Empty;

			sTemp = sSelectedProjectPath.Replace(sProjectePath, string.Empty); //Projekte-Verzeichnis aus Projekt-Verzeichnis entfernen
			sTemp = Path.ChangeExtension(sTemp, null); //Erweiterung entfernen

			//Überprüfen ob mit "\" beginnt, dann entfernen			
			if (Path.IsPathRooted(sTemp))
			{
				sTemp = sTemp.TrimStart(Path.DirectorySeparatorChar);
				sTemp = sTemp.TrimStart(Path.AltDirectorySeparatorChar);
			}

			if (!txtZielverzeichnis.Text.EndsWith(@"\"))
			{
				txtZielverzeichnis.Text += @"\";
			}

			//neues Zielverzeichnis eintragen
			txtZielverzeichnis.Text = Path.Combine(txtZielverzeichnis.Text, sTemp);
			txtZielverzeichnis.Select(); // to Set Focus
			txtZielverzeichnis.Select(txtZielverzeichnis.Text.Length, 0); //to set cursor at the end of textbox
		}
	}

	//Kontextmenü 'Komplette Projekt-Verzeichnisstruktur hinzufügen'
	private void toolStripMenuItemProjektVerzeichnisstruktur_Click(object sender, EventArgs e)
	{
		string sSelectedProjectPath = string.Empty;
#if DEBUG
		sSelectedProjectPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\Test\Test.edb";
#else
		sSelectedProjectPath = PathMap.SubstitutePath(@"$(P)");
#endif
		if (sSelectedProjectPath != string.Empty)
		{
			string sTemp = string.Empty;

			string sRootPath = Path.GetPathRoot(sSelectedProjectPath); //Root-Verzeichnis ermitteln
			sTemp = sSelectedProjectPath.Replace(sRootPath, string.Empty); //Root-Verzeichnis aus Projekt-Verzeichnis entfernen
			sTemp = Path.ChangeExtension(sTemp, null); //Erweiterung entfernen

			//Überprüfen ob mit "\" beginnt, dann entfernen			
			if (Path.IsPathRooted(sTemp))
			{
				sTemp = sTemp.TrimStart(Path.DirectorySeparatorChar);
				sTemp = sTemp.TrimStart(Path.AltDirectorySeparatorChar);
			}

			if (!txtZielverzeichnis.Text.EndsWith(@"\"))
			{
				txtZielverzeichnis.Text += @"\";
			}

			//neues Zielverzeichnis eintragen
			txtZielverzeichnis.Text = Path.Combine(txtZielverzeichnis.Text, sTemp);
			txtZielverzeichnis.Select(); // to Set Focus
			txtZielverzeichnis.Select(txtZielverzeichnis.Text.Length, 0); //to set cursor at the end of textbox
		}
	}

	//Kontextmenü 'Pfadvariable einfügen'
	private void toolStripMenuItemPfadvariabeleItem_Click(object sender, EventArgs e)
	{
#if !DEBUG
		string value = null;
		ActionCallingContext actionCallingContext = new ActionCallingContext();
		actionCallingContext.AddParameter("DialogName", "XSDSelectDBPathVariableDialog");
		new CommandLineInterpreter().Execute("GfDialogManagerDoModal", actionCallingContext);
		actionCallingContext.GetParameter("selectedpathvariable", ref value);
		txtZielverzeichnis.Text += value;
#endif
	}
	
	//Struktur: Verzeichnisebene hinzufügen
	private void toolStripMenuStrukturVerzeichnisebene_Click(object sender, EventArgs e)
	{
		txtAblageStruktur.Focus();
		txtAblageStruktur.Text = txtAblageStruktur.Text.Insert(txtAblageStruktur.SelectionStart, @"\");
		txtAblageStruktur.SelectionStart = txtAblageStruktur.Text.Length;
	}

	//Struktur: Hersteller hinzufügen
	private void toolStripMenuStrukturHersteller_Click(object sender, EventArgs e)
	{
		txtAblageStruktur.Focus();
		txtAblageStruktur.Text = txtAblageStruktur.Text.Insert(txtAblageStruktur.SelectionStart, "$(MANUFACTURER)");
		txtAblageStruktur.SelectionStart = txtAblageStruktur.Text.Length;
	}

	//Struktur: Produktgruppe hinzufügen
	private void toolStripMenuStrukturProduktgruppe_Click(object sender, EventArgs e)
	{
		txtAblageStruktur.Focus();
		txtAblageStruktur.Text = txtAblageStruktur.Text.Insert(txtAblageStruktur.SelectionStart, "$(PRODUCTGROUP)");
		txtAblageStruktur.SelectionStart = txtAblageStruktur.Text.Length;
	}

	//Struktur: Artikelnummer hinzufügen
	private void toolStripMenuStrukturArtikelnummer_Click(object sender, EventArgs e)
	{
		txtAblageStruktur.Focus();
		txtAblageStruktur.Text = txtAblageStruktur.Text.Insert(txtAblageStruktur.SelectionStart, "$(PARTNR)");
		txtAblageStruktur.SelectionStart = txtAblageStruktur.Text.Length;
	}

	//Kontextmenü listView öffnen
	private void contextMenuListView_Opening(object sender, System.ComponentModel.CancelEventArgs e)
	{
		if (listView.SelectedItems.Count != 0)
		{
			if (listView.SelectedItems[0].ForeColor == System.Drawing.Color.Gray)
			{
				toolStripMenuItemDocumentOpen.Enabled = false;
			}
			else
			{
				toolStripMenuItemDocumentOpen.Enabled = true;
			}
		}
		else
		{
			toolStripMenuItemDocumentOpen.Enabled = false;
		}
		return;
	}

	//Kontextmenü 'Dokument öffnen'
	private void toolStripMenuItemDocumentOpen_Click(object sender, EventArgs e)
	{
		if (listView.SelectedItems.Count == 1)
		{
			System.Diagnostics.Process.Start(listView.SelectedItems[0].Text);
			return;
		}
	}

	//Einstellungen speichen
	private void SettingsWrite()
	{
#if !DEBUG
		Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
		if (!oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Top"))
		{
			oSettings.AddNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Top",
			new int[] { },
			new Range[] { },
			ISettings.CreationFlag.Insert);
		}
		oSettings.SetNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Top", this.Top, 0);

		if (!oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Left"))
		{
			oSettings.AddNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Left",
				new int[] { },
				new Range[] { },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Left", this.Left, 0);

		if (!oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Height"))
		{
			oSettings.AddNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Height",
				new int[] { },
				new Range[] { },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Height", this.Height, 0);

		if (!oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Width"))
		{
			oSettings.AddNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Width",
				new int[] { },
				new Range[] { },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetNumericSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.Width", this.Width, 0);

		if (!oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.txtAblageStruktur"))
		{
			oSettings.AddStringSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.txtAblageStruktur",
				new string[] { },
				new string[] { },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetStringSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.txtAblageStruktur", txtAblageStruktur.Text, 0);

		if (!oSettings.ExistSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.txtZielverzeichnis"))
		{
			oSettings.AddStringSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.txtZielverzeichnis",
				new string[] { },
				new string[] { },
				ISettings.CreationFlag.Insert);
		}
		oSettings.SetStringSetting("USER.SCRIPTS.DOCUMENTATION_TOOL.txtZielverzeichnis", txtZielverzeichnis.Text, 0);

#endif
		return;
	}

	// //Einstellungen löschen
	// public void SettingsDelete()
	// {
	// 	DialogResult Result = MessageBox.Show("Sollen die Einstellungen wirklich aus EPLAN gelöscht werden?", "Documentation-Tool, Einstellungen löschen", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
	// 	if (Result == System.Windows.Forms.DialogResult.Yes)
	// 	{
	// 		//Anzahl der zu löschenden Settings anzeigen
	// 		Eplan.EplApi.Base.SettingNode oSettingNode = new Eplan.EplApi.Base.SettingNode("USER.SCRIPTS.DOCUMENTATION_TOOL");
	// 		MessageBox.Show("Es wurden " + oSettingNode.GetCountOfSettings().ToString() + " Einstellungen gelöscht.", "Documentation-Tool, Einstellungen löschen", MessageBoxButtons.OK, MessageBoxIcon.Information);
	// 		//Settings löschen
	// 		oSettingNode.ResetNode();
	// 	}
	// 	return;
	// }

	//ungültige Zeichen aus Dateinamen entfernen
	private static string RemoveIlegaleCharackter(string fileName)
	{
		string illegal = fileName;
		string invalid = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());

		foreach (char c in invalid)
		{
			illegal = illegal.Replace(c.ToString(), "");
		}

		return illegal;
	}

	//Produktgruppe Zahl in Klartext umwandeln
	private static string TranslateProductgroup(string Productgroup)
	{
		switch (Productgroup)
		{
			case "0": return "Undefiniert";
			case "1": return "Allgemeine";
			case "2": return "Relais, Schütze";
			case "3": return "Klemmen";
			case "4": return "Stecker";
			case "5": return "Umformer";
			case "6": return "Schutzeinrichtungen";
			case "7": return "Röhren, Halbleiter";
			case "8": return "Meldeeinrichtungen";
			case "9": return "Motoren";
			case "10": return "Messgeräte, Prüfeinrichtungen";
			case "11": return "Widerstände";
			case "12": return "Sensorik, Schalter und Taster";
			case "13": return "Transformatoren";
			case "14": return "Modulatoren";
			case "15": return "Elektrisch betätigte mechanische Einrichtungen";
			case "16": return "Elektrotechnik Sonderbauteile";
			case "17": return "Verschiedenes";
			case "18": return "Kondensatoren";
			case "19": return "Logikbauteile";
			case "20": return "Spannungsquelle und Generator";
			case "21": return "Induktivitäten";
			case "22": return "Verstärker, Regler";
			case "23": return "Starkstrom - Schaltgeräte";
			case "24": return "Abschlüsse, Filter";
			case "25": return "Übertragungswege";
			case "26": return "SPS";
			case "29": return "Kabel";
			case "30": return "Aggregate und Anlagen";
			case "32": return "Aktoren, allgemein";
			case "33": return "Fluidmotor";
			case "34": return "Filter";
			case "35": return "Fluid Control Terminal";
			case "36": return "Kupplungen";
			case "37": return "Verbindungen";
			case "38": return "Messanschlüsse";
			case "39": return "Messgeräte";
			case "40": return "Pumpen";
			case "41": return "Signalaufnehmer";
			case "42": return "Fluid Sonderbauteile";
			case "43": return "Speicher";
			case "44": return "Ventile";
			case "46": return "Wärmetauscher";
			case "47": return "Zubehör";
			case "48": return "Anschlussplatten";
			case "49": return "Gehäuse";
			case "50": return "Gehäusezubehör Außenanbau";
			case "51": return "Gehäusezubehör Innenanbau";
			case "52": return "Verschlusssysteme";
			case "53": return "Kabelkanäle";
			case "54": return "Sammelschienen";
			case "55": return "Schaltschrank";
			case "56": return "Abzugshaube";
			case "57": return "Elektrolysezelle";
			case "58": return "Lagerung";
			case "59": return "Schornstein";
			case "61": return "Antriebsmaschine";
			case "62": return "Absperrarmatur";
			case "63": return "Dreiwegearmatur";
			case "64": return "Rückschlagarmatur";
			case "65": return "Behälter";
			case "66": return "Behälteranschluss";
			case "67": return "Abscheider";
			case "68": return "Filter";
			case "69": return "Sieb";
			case "70": return "Förderer";
			case "71": return "Heber";
			case "72": return "Transporter";
			case "73": return "Behälter - Mantel";
			case "74": return "Behälter - Rohrschlange";
			case "75": return "Dampfkessel";
			case "76": return "Kühler";
			case "77": return "Ofen";
			case "78": return "Trockner";
			case "79": return "Verdampfer";
			case "80": return "Wärmetauscher";
			case "81": return "Kneter";
			case "82": return "Mischer";
			case "83": return "Rührer";
			case "84": return "Kompressor / Verdichter";
			case "85": return "Pumpen";
			case "86": return "Vakuumpumpe";
			case "87": return "Ventilator";
			case "89": return "Rohrleitungsteil";
			case "91": return "Zentrifuge";
			case "92": return "Messeinrichtung";
			case "93": return "Waage";
			case "94": return "Formgeber";
			case "95": return "Sichter";
			case "96": return "Sortierer";
			case "97": return "Zerkleinerer";
			case "98": return "Zuteiler";
			case "99": return "Feldverteiler";
			case "100": return "Verbindungen";
			case "101": return "Montageplatten";
			case "102": return "Heizung";
			case "103": return "Leuchte";
			case "104": return "Heizung";
			case "105": return "Zylinder";
			case "106": return "Bremse";
			case "107": return "Gleichrichter";
			case "108": return "Bremse";
			case "109": return "Kupplungen";
			case "110": return "Ventile";
			case "111": return "Druckübersetzer";
			case "112": return "Abscheider";
			case "113": return "Leitungsverteiler / -verbinder";
			case "114": return "Zubehör";
			case "115": return "Meldeeinrichtungen";
			case "116": return "Benutzerdefinierte Schiene";
			case "117": return "19 Zoll - Ausbautechnik";
			case "118": return "Strecke(Topologie)";
			case "119": return "Verlegepunkt(Topologie)";
			case "121": return "Verbindungen";
			case "125": return "Verlegezubehör";
			case "126": return "Montageanordnung";
			case "127": return "Leitungen / Pakete";
			case "128": return "Leitungen / Pakete";
			case "129": return "Leitungsverteiler / -verbinder";
			case "130": return "Leitungsverteiler / -verbinder";
			default: return "Unbekannte Produktgruppe";
		}
	}

}

// Implements the manual sorting of items by columns.
class ListViewItemComparer : IComparer
{
	private int col;
	private SortOrder order;
	public ListViewItemComparer()
	{
		col = 0;
		order = SortOrder.Ascending;
	}
	public ListViewItemComparer(int column, SortOrder order)
	{
		col = column;
		this.order = order;
	}
	public int Compare(object x, object y)
	{
		int returnVal = -1;
		returnVal = String.Compare(((ListViewItem)x).SubItems[col].Text,
								((ListViewItem)y).SubItems[col].Text);
		// Determine whether the sort order is descending.
		if (order == SortOrder.Descending)
			// Invert the value returned by String.Compare.
			returnVal *= -1;
		return returnVal;
	}
}
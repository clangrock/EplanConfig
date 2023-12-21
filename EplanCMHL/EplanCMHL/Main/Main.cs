/*
Lädt das Menüe um die Unterpunkte zu erzeugen für die Exportfunktionen
mit Prüfung ob Nutzer der BU Schkopau/Schkeuditz angemeldet sind
Christian Langrock
Version 1.5     21.12.2022

Original Projekt liegt auf GitHub: https://github.com/Actemium-Schkeuditz/EplanConfig

Änderungesverlauf
V0.2 update domainName nach Änderung durch IT
V1.0 verschoben auf gitHUB
V1.1 integration Documentation Tool
V1.2 update Aufruf PDF Assistent
V1.3 Ergänzung der Domain für BU Weil und Büro Achern
V1.4 Adaption für LKO Electro
V1.5 update auf Eplan 2024, update Dokumentation Tool auf V3.0,  aufräumen
*/

using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Scripting;
using Eplan.EplApi.Base;
using System;
using System.Security.Principal;
using System.Windows;

namespace EplanCMHL.Ausgabe
{
    class CreateMenue
    {
        // Menü zusammenbauen
        [DeclareMenu]
        public void MenuFunction()
        {
            try // Fehlerbehandlung
            {
                //Deklarationen
                //uint MenuID = new uint(); // Menü-ID vom neu erzeugten Menü   

                Eplan.EplApi.Gui.Menu oMenu = new Eplan.EplApi.Gui.Menu();
                // Abfragen der Menue-ID
                uint MenuID = oMenu.GetPersistentMenuId("Export / Beschriftung...");

                oMenu.AddMenuItem(
                        "Ausgabe Projektisten", // Name: Menüpunkt
                        "ExcecuteSummaryPartlist", // Name: Action
                        "Ausgabe verschiedener Listen", // Statustext
                        MenuID, // Menü-ID:
                        1, // 1 = Hinter Menüpunkt, 0 = Vor Menüpunkt
                        false, // Seperator davor anzeigen
                        false // Seperator dahinter anzeigen
                    );

                // Makro Navigator 
                uint presMenuId = 37024; //Menü-ID: Einfügen/Fenstermakro...
                oMenu.AddMenuItem("Makros Einfuegen mit Navigator",
                    "ShowMacroNavi",
                    "Navigator zum Einfügen von Makros",
                    presMenuId,
                    1,
                    false,
                    false
                );

                // PDF Assistant
                oMenu.AddMenuItem("PDF (Assistent)...",     // Name: Menüpunkt
                    "PDFAssistent_Start",                   // Name: Action
                    "PDF Assistent," +                      // Statustext
                    " aktuelles Projekt als PDF-Datei exportieren",
                    35287,
                    1,
                    false,
                    false
                );

                // Documentation Tool
                presMenuId = 35379; //Menü-ID: Einfügen/Fenstermakro...
                    oMenu.AddMenuItem("Dokumentations-Tool...",
                    "ShowDocumentationTool",                // Name: Action
                    "Externe Dokumente ermitteln und kopieren", 
                    presMenuId,
                    1,
                    false,
                    false
                );

                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

 //ab hier Eplan 2024
	//RibbonBar Einträge dfinieren
	string m_TabName = "Werkzeuge";
	string m_commandGroupName = "Erweiterungen";
	string m_commandNameDOK = "Dokumentations-Tool";
    string m_commandNameExport ="Export-Daten";
    string m_commandNamePDFExport ="PDF Export";

	//Skript wird geladen
	[DeclareRegister]
	public void OnRegisterScript()
	{
		var newTab = new Eplan.EplApi.Gui.RibbonBar().Tabs.FirstOrDefault(item => item.Name == m_TabName);
		if (newTab == null) //Tab noch nicht vorhanden, dann neu erzeugen
		{
			newTab = new Eplan.EplApi.Gui.RibbonBar().AddTab(m_TabName);
		}
		var commandGroup = newTab.CommandGroups.FirstOrDefault(item => item.Name == m_commandGroupName);
		if (commandGroup == null) //CommandGroup noch nicht vorhanden, dann neu erzeugen
		{
			commandGroup = newTab.AddCommandGroup(m_commandGroupName);
		}
		Eplan.EplApi.Gui.RibbonIcon ribbonIcon = new Eplan.EplApi.Gui.RibbonIcon(Eplan.EplApi.Gui.CommandIcon.TaskList); //Icon festlegen
		commandGroup.AddCommand(m_commandNameDOK, "ShowDocumentationTool", m_commandNameDOK, "Externe Dokumente ermitteln und kopieren", ribbonIcon);
        commandGroup.AddCommand(m_commandNameExport, "ExcecuteSummaryPartlist", m_commandNameExport, "Ausgabe verschiedener Listen", ribbonIcon);
        commandGroup.AddCommand(m_commandNamePDFExport, "PDFAssistent_Start", m_commandNamePDFExport, "aktuelles Projekt als PDF-Datei exportieren",ribbonIcon);
	}

	//Skript wird entladen
	[DeclareUnregister]
	public void OnUnRegisterScript()
	{
		//Einstellungen entfernen
		SettingsDelete();

		//Command entfernen
		var vTab = new Eplan.EplApi.Gui.RibbonBar().Tabs.FirstOrDefault(item => item.Name == m_TabName);
		if (vTab != null)
		{
			var commandGroup = vTab.CommandGroups.FirstOrDefault(item => item.Name == m_commandGroupName);
			if (commandGroup != null)
			{
				var command = commandGroup.Commands.Values.FirstOrDefault(item => item.Text == m_commandNameDOK);
				if (command != null)
				{
					command.Remove();
				}
                var commandExport = commandGroup.Commands.Values.FirstOrDefault(item => item.Text == m_commandNameExport);
				if (commandExport != null)
				{
					commandExport.Remove();
				}
				var commandPDFExport = commandGroup.Commands.Values.FirstOrDefault(item => item.Text == m_commandNamePDFExport);
				if (commandPDFExport != null)
				{
					commandPDFExport.Remove();
				}
                //Wenn CommandGroup leer ist diese auch entfernen
				if (commandGroup.Commands.Count == 0)
				{
					commandGroup.Remove();
				}
			}
			//Wenn Tab leer ist dieses auch entfernen
			if (vTab.Commands.Count == 0)
			{
				vTab.Remove();
			}
		}
	}
        	//Einstellungen löschen
        public void SettingsDelete()
        {
            // Documentation Tool
            DialogResult Result = MessageBox.Show("Sollen die Einstellungen wirklich aus EPLAN gelöscht werden?", "Documentation-Tool, Einstellungen löschen", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (Result == System.Windows.Forms.DialogResult.Yes)
            {
                //Anzahl der zu löschenden Settings anzeigen
                Eplan.EplApi.Base.SettingNode oSettingNode = new Eplan.EplApi.Base.SettingNode("USER.SCRIPTS.DOCUMENTATION_TOOL");
                MessageBox.Show("Es wurden " + oSettingNode.GetCountOfSettings().ToString() + " Einstellungen gelöscht.", "Documentation-Tool, Einstellungen löschen", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //Settings löschen
                oSettingNode.ResetNode();
            }
            // PDF Einstellunben löschen 
            DialogResult ResultPDF = MessageBox.Show("Sollen die Einstellungen wirklich aus EPLAN gelöscht werden?", "PDF Assistent_Einstellungen löschen", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (ResultPDF == System.Windows.Forms.DialogResult.Yes)
            {
                //Settings löschen
                Eplan.EplApi.Base.SettingNode oSettingNodePDF = new Eplan.EplApi.Base.SettingNode("USER.SCRIPTS.PDF_Assistent");

                MessageBox.Show("Es wurden " + oSettingNodePDF.GetCountOfSettings().ToString() + " Einstellungen gelöscht.", "PDF Assistent_Einstellungen löschen", MessageBoxButtons.OK, MessageBoxIcon.Information);
              //  oSettingNodePDF.DeleteNode();
              oSettingNodePDF.ResetNode();
            }
            return;
        }


        [DeclareAction("ExcecuteSummaryPartlist")]

        public void ExcecuteSummaryPartList()
        {
            string scriptName = @"$(MD_SCRIPTS)\Ausgabe\AusgabeListen.cs";
            scriptName = PathMap.SubstitutePath(scriptName);
            ExcecuteScripts(scriptName);
        }


        [DeclareAction("ShowMacroNavi")]

        public void ExcecuteMacroNavigator()
        {
            string scriptName = @"$(MD_SCRIPTS)\MakroNavigator\MacroNaviForm.cs";
            scriptName = PathMap.SubstitutePath(scriptName);
            ExcecuteScripts(scriptName);
        }

        [DeclareAction("PDFAssistent_Start")]

        public void ExcecutePDF_Assistent()
        {
            string scriptName = @"$(MD_SCRIPTS)\Ausgabe\PDF\PDF Assistent.cs";
            scriptName = PathMap.SubstitutePath(scriptName);
            ExcecuteScripts(scriptName, "1");
        }

        [DeclareEventHandler("onActionStart.String.XPrjActionProjectClose")]
        public void ExcecuteBackupScript()
        {
            string scriptName = @"$(MD_SCRIPTS)\Backup\BackupOnClosingProject_V03.cs";
            // Pfad auflösen
            scriptName = PathMap.SubstitutePath(scriptName);
            ExcecuteScripts(scriptName);
        }

        [DeclareAction("ShowDocumentationTool")]
        public void ExcecuteDocumentationScript()
        {
            string scriptName = @"$(MD_SCRIPTS)\DocumentationTool\DocumentationTool.cs";
            // Pfad auflösen
            scriptName = PathMap.SubstitutePath(scriptName);
            ExcecuteScripts(scriptName);
        }


        private void ExcecuteScripts(string sScriptName, string Param = "")
        {
            try
            {
                // Ausgabe der Einbauorte 
                // Parameter für die Action
                ActionCallingContext actionCallingContext = new ActionCallingContext();
                actionCallingContext.AddParameter("ScriptFile", sScriptName);
                if (!string.IsNullOrEmpty(Param))
                {
                    actionCallingContext.AddParameter("Param1", Param);
                }
                // ausführen
                new CommandLineInterpreter().Execute("ExecuteScript", actionCallingContext);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }

        }


        private void loadScripts(string sScriptName)
        {
            try
            {
                // Parameter für die Action
                ActionCallingContext actionCallingContext = new ActionCallingContext();
                actionCallingContext.AddParameter("ScriptFile", sScriptName);
                actionCallingContext.AddParameter("RegisterAgain", "1");        // Ermöglicht, das Scipt erneut zu registrieren
                actionCallingContext.AddParameter("ShowDecider", "0");          // Blendet eine Abfrage ein, ob das Scipt erneut registriert werden soll 

                // ausführen
                new CommandLineInterpreter().Execute("RegisterScript", actionCallingContext);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }
        }


        // ab hier alles für PDF Assistent

         //Event ProjectClose abfangen
	    [DeclareEventHandler("onActionStart.String.XPrjActionProjectClose")]
	    public void Project_Close_Event()
	         {
            //Einstellung einlesen
            Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
            if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.ByProjectClose"))
            {
                bool bChecked = oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.ByProjectClose", 1);
                if (bChecked) //Bei ProjectClose ausführen
                {
                    PDFAssistent_SollStarten();
                }
            }
            return;
        }

        [DeclareEventHandler("Eplan.EplApi.OnMainEnd")]
        public void ExcecutePDFonEplanEnd()
        {
                //Einstellung einlesen
                Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
                if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.ByEplanEnd"))
                {
                    bool bChecked = oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.ByEplanEnd", 1);
                    if (bChecked) //Bei EplanEnd ausführen
                    {
                        PDFAssistent_SollStarten();
                    }
                }
                return;
        }

	//Prüfen ob Assistent gestartet werden soll
	
	public void PDFAssistent_SollStarten()
	{
		Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();

		//Ist Projekt in Projekt-Ordner
		//Wenn angekreuzt dann muß Projekt im Ordner sein für Assistent, sonst kein Assistent
		//Wenn nicht angekreuzt dann Assistent
		if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.OhneNachfrage"))
		{
			bool bChecked = oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.ProjectInOrdner", 1);
			string sProjektInOrdner = oSettings.GetStringSetting("USER.SCRIPTS.PDF_Assistent.ProjectInOrdnerName", 0);
			if (bChecked)
			{
				string sProjektOrdner = PathMap.SubstitutePath("$(PROJECTPATH)");
				sProjektOrdner = sProjektOrdner.Substring(0, sProjektOrdner.LastIndexOf(@"\"));
				if (!sProjektOrdner.EndsWith(@"\"))
				{
					sProjektOrdner += @"\";
				}
				if (sProjektInOrdner == sProjektOrdner) //hier vielleicht noch erweitern auf alle Unterordner (InString?)
				{
					PDFAssistent_ausführen();
				}
			}
			else
			{
				PDFAssistent_ausführen();
			}
		}
	}

        public void PDFAssistent_ausführen()
        {
            Eplan.EplApi.Base.Settings oSettings = new Eplan.EplApi.Base.Settings();
            if (oSettings.ExistSetting("USER.SCRIPTS.PDF_Assistent.OhneNachfrage"))
            {
                bool bChecked = oSettings.GetBoolSetting("USER.SCRIPTS.PDF_Assistent.OhneNachfrage", 1);
                if (bChecked)
                {
                    string scriptName = @"$(MD_SCRIPTS)\Ausgabe\PDF\PDF Assistent.cs";
                    scriptName = PathMap.SubstitutePath(scriptName);
                    ExcecuteScripts(scriptName, "1");
                }
                else
                {
                    string scriptName = @"$(MD_SCRIPTS)\Ausgabe\PDF\PDF Assistent.cs";
                    scriptName = PathMap.SubstitutePath(scriptName);
                    ExcecuteScripts(scriptName, "0");
                }
            }
        }
    
   
    }
}
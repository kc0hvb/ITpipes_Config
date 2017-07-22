using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Speech.Synthesis;
using ITpipes_Config.SharedEnums;

namespace ITpipes_Config {
   
    public class ITpipesGeneralSettings : Freezable {

        #region Helper Properties

        private string globalOleConnString { get; set; }//
        public SpeechSynthesizer curSpeechSynthObject { get; set; }//

        #endregion

        #region Valid Itemssource Properties

        public string[] SpeechSynthVoices {
            get { return (string[])GetValue(SpeechSynthVoicesProperty); }
            set { SetValue(SpeechSynthVoicesProperty, value); }
        }
        public static readonly DependencyProperty SpeechSynthVoicesProperty =
            DependencyProperty.Register("SpeechSynthVoices", typeof(string[]), typeof(ITpipesGeneralSettings));
        
        public Array TtsSpeedArray {
            get { return (Array)GetValue(TtsSpeedArrayProperty); }
            set { SetValue(TtsSpeedArrayProperty, value); }
        }
        public static readonly DependencyProperty TtsSpeedArrayProperty =
            DependencyProperty.Register("TtsSpeedArray", typeof(Array), typeof(ITpipesGeneralSettings));

        public Array TextCasesArray {
            get { return (Array)GetValue(TextCasesArrayProperty); }
            set { SetValue(TextCasesArrayProperty, value); }
        }
        public static readonly DependencyProperty TextCasesArrayProperty =
            DependencyProperty.Register("TextCasesArray", typeof(Array), typeof(ITpipesGeneralSettings));

        public Array TextDisplaysArray {
            get { return (Array)GetValue(TextDisplaysArrayProperty); }
            set { SetValue(TextDisplaysArrayProperty, value); }
        }
        public static readonly DependencyProperty TextDisplaysArrayProperty =
            DependencyProperty.Register("TextDisplaysArray", typeof(Array), typeof(ITpipesGeneralSettings));
        
        #endregion

        #region Grid Display Option Properties

        public TextCases TextCase {
            get { return (TextCases)GetValue(TextCaseProperty); }
            set { SetValue(TextCaseProperty, value); }
        }
        public static readonly DependencyProperty TextCaseProperty =
            DependencyProperty.Register("TextCase", typeof(TextCases), typeof(ITpipesGeneralSettings));

        public TextDisplays TextDisplayType {
            get { return (TextDisplays)GetValue(TextDisplayTypeProperty); }
            set { SetValue(TextDisplayTypeProperty, value); }
        }
        public static readonly DependencyProperty TextDisplayTypeProperty =
            DependencyProperty.Register("TextDisplayType", typeof(TextDisplays), typeof(ITpipesGeneralSettings));

        public bool ShowGridNumbers {
            get { return (bool)GetValue(ShowGridNumbersProperty); }
            set { SetValue(ShowGridNumbersProperty, value); }
        }
        public static readonly DependencyProperty ShowGridNumbersProperty =
            DependencyProperty.Register("ShowGridNumbers", typeof(bool), typeof(ITpipesGeneralSettings));

        #endregion

        #region Miscellaneous Settings properties

        public bool AutoRefreshAFS {
            get { return (bool)GetValue(AutoRefreshAFSProperty); }
            set { SetValue(AutoRefreshAFSProperty, value); }
        }
        public static readonly DependencyProperty AutoRefreshAFSProperty =
            DependencyProperty.Register("AutoRefreshAFS", typeof(bool), typeof(ITpipesGeneralSettings));

        public bool BackupProjects {
            get { return (bool)GetValue(BackupProjectsProperty); }
            set { SetValue(BackupProjectsProperty, value); }
        }
        public static readonly DependencyProperty BackupProjectsProperty =
            DependencyProperty.Register("BackupProjects", typeof(bool), typeof(ITpipesGeneralSettings));

        public bool AutosaveReports {
            get { return (bool)GetValue(AutosaveReportsProperty); }
            set { SetValue(AutosaveReportsProperty, value); }
        }
        public static readonly DependencyProperty AutosaveReportsProperty =
            DependencyProperty.Register("AutosaveReports", typeof(bool), typeof(ITpipesGeneralSettings));

        public bool EnableRemarksPersistence {
            get { return (bool)GetValue(EnableRemarksPersistenceProperty); }
            set { SetValue(EnableRemarksPersistenceProperty, value); }
        }
        public static readonly DependencyProperty EnableRemarksPersistenceProperty =
            DependencyProperty.Register("EnableRemarksPersistence", typeof(bool), typeof(ITpipesGeneralSettings));

        public bool EnableSplashScreen {
            get { return (bool)GetValue(EnableSplashScreenProperty); }
            set { SetValue(EnableSplashScreenProperty, value); }
        }
        public static readonly DependencyProperty EnableSplashScreenProperty =
            DependencyProperty.Register("EnableSplashScreen", typeof(bool), typeof(ITpipesGeneralSettings));

        public string DefaultProjectPath {
            get { return (string)GetValue(DefaultProjectPathProperty); }
            set { SetValue(DefaultProjectPathProperty, value); }
        }
        public static readonly DependencyProperty DefaultProjectPathProperty =
            DependencyProperty.Register("DefaultProjectPath", typeof(string), typeof(ITpipesGeneralSettings));

        #endregion

        #region Text-To-Speech Settings Properties

        public bool TtsEnabled {
            get { return (bool)GetValue(TtsEnabledProperty); }
            set { SetValue(TtsEnabledProperty, value); }
        }
        public static readonly DependencyProperty TtsEnabledProperty =
            DependencyProperty.Register("TtsEnabled", typeof(bool), typeof(ITpipesGeneralSettings));

        public TtsSpeeds TtsSpeed {
            get { return (TtsSpeeds)GetValue(TtsSpeedProperty); }
            set { SetValue(TtsSpeedProperty, value); }
        }
        public static readonly DependencyProperty TtsSpeedProperty =
            DependencyProperty.Register("TtsSpeed", typeof(TtsSpeeds), typeof(ITpipesGeneralSettings));
        
        public string TtsVoice {
            get { return (string)GetValue(TtsVoiceProperty); }
            set { SetValue(TtsVoiceProperty, value); }
        }
        public static readonly DependencyProperty TtsVoiceProperty =
            DependencyProperty.Register("TtsVoice", typeof(string), typeof(ITpipesGeneralSettings));

        public string TtsTestText {
            get { return (string)GetValue(TtsTestTextProperty); }
            set { SetValue(TtsTestTextProperty, value); }
        }
        public static readonly DependencyProperty TtsTestTextProperty =
            DependencyProperty.Register("TtsTestText", typeof(string), typeof(ITpipesGeneralSettings));

        #endregion

        #region  Advanced Settings Properties

        public bool DisableVideoFeedCloning {
            get { return (bool)GetValue(DisableVideoFeedCloningProperty); }
            set { SetValue(DisableVideoFeedCloningProperty, value); }
        }
        public static readonly DependencyProperty DisableVideoFeedCloningProperty =
            DependencyProperty.Register("DisableVideoFeedCloning", typeof(bool), typeof(ITpipesGeneralSettings));

        public bool EnableVideoDebugging {
            get { return (bool)GetValue(EnableVideoDebuggingProperty); }
            set { SetValue(EnableVideoDebuggingProperty, value); }
        }
        public static readonly DependencyProperty EnableVideoDebuggingProperty =
            DependencyProperty.Register("EnableVideoDebugging", typeof(bool), typeof(ITpipesGeneralSettings));

        public bool DisableRequiredFieldValidation {
            get { return (bool)GetValue(DisableRequiredFieldValidationProperty); }
            set { SetValue(DisableRequiredFieldValidationProperty, value); }
        }
        public static readonly DependencyProperty DisableRequiredFieldValidationProperty =
            DependencyProperty.Register("DisableRequiredFieldValidation", typeof(bool), typeof(ITpipesGeneralSettings));

        public bool UseAdodc {
            get { return (bool)GetValue(UseAdodcProperty); }
            set { SetValue(UseAdodcProperty, value); }
        }
        public static readonly DependencyProperty UseAdodcProperty =
            DependencyProperty.Register("UseAdodc", typeof(bool), typeof(ITpipesGeneralSettings));
        
        public ValidVideoRenderers VideoRenderer {
            get { return (ValidVideoRenderers)GetValue(VideoRendererProperty); }
            set { SetValue(VideoRendererProperty, value); }
        }
        public static readonly DependencyProperty VideoRendererProperty =
            DependencyProperty.Register("VideoRenderer", typeof(ValidVideoRenderers), typeof(ITpipesGeneralSettings));

        #endregion

        #region Constructors

        public ITpipesGeneralSettings(string oleConnStringToSettingsDb = @"Provider = Microsoft.Jet.OLEDB.4.0; Data Source = C:\Program Files\InspectIT\setup.mdb") {

            if (string.IsNullOrEmpty(oleConnStringToSettingsDb)) {
                throw new Exception("Connection string passed to ITpipesGeneralSettings constructor was null or empty");
            }
            
            globalOleConnString = oleConnStringToSettingsDb;
            
            using (OleDbConnection curOleConn = new OleDbConnection(globalOleConnString))
            using (OleDbDataAdapter curOleDA = new OleDbDataAdapter("", curOleConn))
            using (DataTable tempSettingsDT = new DataTable()) {

                curOleConn.Open();
                curOleDA.SelectCommand.CommandText = "SELECT * From S";

                curOleDA.Fill(tempSettingsDT);

                if (tempSettingsDT.Rows.Count == 0) {

                    createNewSettingsRow(oleConnStringToSettingsDb);

                    curOleDA.Fill(tempSettingsDT);

                    }

                //populateCurrentSettings(tempSettingsDT.Rows[0]);
            }
        }

        private static void createNewSettingsRow(string oleConnString) {
            using (OleDbConnection curOleConn = new OleDbConnection(oleConnString))
            using (OleDbCommand newSettingsRowCommand = new OleDbCommand("", curOleConn)) {

                curOleConn.Open();

                newSettingsRowCommand.CommandText =
                    "INSERT INTO S (" +
                    "Video_Quality, " +
                    "Counter_Calibration, " +
                    "Speech_Speed, " +
                    "P_ID, " +
                    "Files_DefaultProjectPath, " +
                    "Misc_SplashScreen, " +
                    "Video_WriteData, " +
                    "Video_Format, " +
                    "Counter_Type, " +
                    "Misc_Backup_Files, " +
                    "Misc_DBUseADODC) " +
                    "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

                newSettingsRowCommand.Parameters.AddWithValue("Video_Quality", "CD");
                newSettingsRowCommand.Parameters.AddWithValue("Counter_Calibration", 1);
                newSettingsRowCommand.Parameters.AddWithValue("Speech_Speed", "Normal");
                newSettingsRowCommand.Parameters.AddWithValue("P_ID", 0);
                newSettingsRowCommand.Parameters.AddWithValue("Files_DefaultProjectPath", @"C:\IT Projects");
                newSettingsRowCommand.Parameters.AddWithValue("Misc_SplashScreen", false);
                newSettingsRowCommand.Parameters.AddWithValue("Video_WriteData", false);
                newSettingsRowCommand.Parameters.AddWithValue("Video_Format", "WMV");
                newSettingsRowCommand.Parameters.AddWithValue("Counter_Type", "USDigital");
                newSettingsRowCommand.Parameters.AddWithValue("Misc_Backup_Files", true);
                newSettingsRowCommand.Parameters.AddWithValue("Misc_DBUseADODC", true);

                newSettingsRowCommand.ExecuteNonQuery();
            }
        }

        protected override Freezable CreateInstanceCore() {
            return new ITpipesGeneralSettings();
        }


        #endregion

        #region Private Methods and Functions

        private void populateInstalledTtsVoices() {

            if (curSpeechSynthObject == null) {
                curSpeechSynthObject = new SpeechSynthesizer();
            }

            var voicesList = curSpeechSynthObject.GetInstalledVoices();
            SpeechSynthVoices = new string[voicesList.Count];

            for (int i = 0; i < voicesList.Count; i++) {
                SpeechSynthVoices[i] = voicesList[i].VoiceInfo.Description;
            }

        }

        //private void populateCurrentSettings(DataRow settingsRow) {

        //    #region Grid Display Settings

        //    this.TextCase = getTextCaseFromString(settingsRow["Misc_TextCase"].ToString());
        //    this.TextDisplayType = getTextDisplayFromString(settingsRow["Text_DisplayType"].ToString());
        //    this.ShowGridNumbers = (bool)settingsRow["Misc_GridNumber"];

        //    #endregion

        //    #region Miscellaneous Settings

        //    this.AutoRefreshAFS = (bool)settingsRow["Files_AutoRefresh_AFS"];
        //    this.BackupProjects = (bool)settingsRow["Misc_Backup_Files"];
        //    this.AutosaveReports = (bool)settingsRow["Reporting_AutoSave"];
        //    this.EnableRemarksPersistence = isRemarksPersistenceEnabled();
        //    this.EnableSplashScreen = (bool)settingsRow["Misc_SplashScreen"];
        //    this.DefaultProjectPath = settingsRow["Files_DefaultProjectPath"].ToString();

        //    #endregion

        //    #region Speech Settings
            
        //    populateInstalledTtsVoices();

        //    this.TtsEnabled = (bool)settingsRow["Speech_Enabled"];
        //    this.TtsSpeed = getTtsSpeedFromString(settingsRow["Speech_Speed"].ToString());
        //    this.TtsVoice = settingsRow["Speech_Voice"].ToString();

        //    #endregion

        //    #region Advanced Settings

        //    this.DisableVideoFeedCloning = isVideoFeedCloningDisabled();
        //    this.EnableVideoDebugging = isVideoDebuggingEnabled();
        //    this.DisableRequiredFieldValidation = settingsRow.Table.Columns.Contains("Misc_CompliancePrep") ? (bool)settingsRow["Misc_CompliancePrep"] : false;
        //    this.UseAdodc = (bool)settingsRow["Misc_DBUseADODC"];
        //    this.VideoRenderer = getEnabledVideoRenderer();

        //    #endregion

        //}
        #endregion

        #region Converters and value lookups
        

        #endregion
    }
}

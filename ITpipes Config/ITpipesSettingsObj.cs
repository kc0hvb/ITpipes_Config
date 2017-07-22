using System;
using System.ComponentModel;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.OleDb;
using System.IO;
using AForge.Video.DirectShow;
using ITpipes_Config.SharedEnums;
using ITPipes.OverlayDrivers.Legacy;

namespace ITpipes_Config {
    

    public class ITpipesSettingsObj : INotifyPropertyChanged, IDisposable {

        private string _globalSettingsConnString = string.Empty;

        #region Read-Only Resources
        private readonly Array _textCasesArray = Enum.GetValues(typeof(SharedEnums.TextCases));
        private readonly Array _textDisplaysArray = Enum.GetValues(typeof(SharedEnums.TextDisplays));
        private readonly Array _validVideoCaptureFormats = Enum.GetValues(typeof(SharedEnums.videoCaptureFormats));
        private readonly Array _ttsSpeedArray = Enum.GetValues(typeof(SharedEnums.TtsSpeeds));
        private readonly Array _validVideoRenderersArray = Enum.GetValues(typeof(SharedEnums.ValidVideoRenderers));
        
        public Array TextCasesArray {
            get {
                return _textCasesArray;
            }
        }
        public Array TextDisplaysArray {
            get {
                return _textDisplaysArray;
            }
        }
        public Array ValidVideoCaptureFormats {
            get {
                return _validVideoCaptureFormats;
            }
        }
        public Array TtsSpeedArray {
            get {
                return _ttsSpeedArray;
            }
        }
        public Array ValidVideoRenderersArray {
            get {
                return _validVideoRenderersArray;
            }
        }

        #endregion Read-Only Resources

        #region Video Capture Properties

        private FilterInfoCollection _aforgeVideoInputSources;
        private ObservableCollection<VideoCaptureProfile> _videoCaptureProfiles;
        private VideoCaptureProfile _curVidCapProfile;
        private VideoCaptureDevice _curAforgeCaptureDevice;
        private FilterInfo _selectedAforgeCaptureDevice;
        private VideoCapturePin _selectedItVideoCapturePin;
        
        public FilterInfoCollection AforgeVideoInputSources {
            get {
                return _aforgeVideoInputSources;
            }
            set {
                _aforgeVideoInputSources = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("AforgeVideoInputSources"));
            }
        }
        public ObservableCollection<VideoCaptureProfile> VideoCaptureProfiles {
            get {
                return _videoCaptureProfiles;
            }
            set {
                _videoCaptureProfiles = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("VideoCaptureProfiles"));
            }
        }
        public VideoCaptureProfile CurVidCapProfile {
            get {
                return _curVidCapProfile;
            }
            set {
                _curVidCapProfile = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("CurVidCapProfile"));
            }
        }
        public VideoCaptureDevice CurAforgeCaptureDevice {
            get {
                return _curAforgeCaptureDevice;
            }
            set {
                _curAforgeCaptureDevice = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("CurAforgeCaptureDevice"));
            }
        }
        public FilterInfo SelectedAforgeCaptureDevice {
            get {
                return _selectedAforgeCaptureDevice;
            }
            set {
                _selectedAforgeCaptureDevice = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("SelectedAforgeCaptureDevice"));
            }
        }
        public VideoCapturePin SelectedItVideoCapturePin {
            get {
                return _selectedItVideoCapturePin;
            }
            set {
                _selectedItVideoCapturePin = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("SelectedItVideoCapturePin"));
            }
        }
        
        #endregion

        #region Overlay Properties

        private OverlaySettings _overlaySettingsObj;
        private DriverLoader _driverLoaderObj;
        private OverlayControlClass _activeCounterProfile;
        
        public OverlaySettings OverlaySettingsObj {
            get {
                return _overlaySettingsObj;
            }
            set {
                _overlaySettingsObj = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("OverlaySettingsObj"));
            }
        }
        public DriverLoader DriverLoaderObj {
            get {
                return _driverLoaderObj;
            }
            set {
                _driverLoaderObj = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("DriverLoaderObj"));
            }
        }
        public OverlayControlClass ActiveCounterProfile {
            get {
                return _activeCounterProfile;
            }
            set {
                _activeCounterProfile = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("ActiveCounterProfile"));
            }
        }
        
        #endregion
        
        #region TTS Settings

        private System.Speech.Synthesis.SpeechSynthesizer _curSpeechSynthObject = new System.Speech.Synthesis.SpeechSynthesizer();
        private bool _ttsEnabled;
        private TtsSpeeds _ttsSpeed;
        private string _ttsVoice;
        private string _ttsTestText;
        private string[] _speechSynthVoices;

        public System.Speech.Synthesis.SpeechSynthesizer CurSpeechSynthObject {
            get {
                return _curSpeechSynthObject;
            }
        }
        public bool TtsEnabled {
            get {
                return _ttsEnabled;
            }
            set {
                _ttsEnabled = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("TtsEnabled"));
            }
        }
        public TtsSpeeds TtsSpeed {
            get {
                return _ttsSpeed;
            }
            set {
                _ttsSpeed = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("TtsSpeed"));
            }
        }
        public string TtsVoice {
            get {
                return _ttsVoice;
            }
            set {
                _ttsVoice = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("TtsVoice"));
            }
        }
        public string TtsTestText {
            get {
                return _ttsTestText;
            }
            set {
                _ttsTestText = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("TtsTestText"));
            }
        }
        public string[] SpeechSythVoices {
            get {
                return _speechSynthVoices;
            }
            set {
                _speechSynthVoices = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("SpeechSythVoices"));
            }
        }
        
        #endregion
        
        #region Grid Display Settings

        private SharedEnums.TextCases _textCase;
        private SharedEnums.TextDisplays _textDisplayType;
        private bool _showGridNumbers;

        public TextCases TextCase {
            get {
                return _textCase;
            }
            set {
                _textCase = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("TextCase"));
            }
        }
        public TextDisplays TextDisplayType {
            get {
                return _textDisplayType;
            }
            set {
                _textDisplayType = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("TextDisplayType"));
            }
        }
        public bool ShowGridNumbers {
            get {
                return _showGridNumbers;
            }
            set {
                _showGridNumbers = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("ShowGridNumbers"));
            }
        }
        
        #endregion

        #region Miscellaneous Settings

        private bool _autoRefreshAFS;
        private bool _backupProjects;
        private bool _autosaveReports;
        private bool _enableRemarksPersistence;
        private bool _enableSplashScreen;
        private string _defaultProjectPath;

        public bool AutoRefreshAFS {
            get {
                return _autoRefreshAFS;
            }
            set {
                _autoRefreshAFS = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("AutoRefreshAFS"));
            }
        }
        public bool BackupProjects {
            get {
                return _backupProjects;
            }
            set {
                _backupProjects = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("BackupProjects"));
            }
        }
        public bool AutosaveReports {
            get {
                return _autosaveReports;
            }
            set {
                _autosaveReports = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("AutosaveReports"));
            }
        }
        public bool EnableRemarksPersistence {
            get {
                return _enableRemarksPersistence;
            }
            set {
                _enableRemarksPersistence = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("EnableRemarksPersistence"));
            }
        }
        public bool EnableSplashScreen {
            get {
                return _enableSplashScreen;
            }
            set {
                _enableSplashScreen = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("EnableSplashScreen"));
            }
        }
        public string DefaultProjectPath {
            get {
                return _defaultProjectPath;
            }
            set {
                _defaultProjectPath = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("DefaultProjectPath"));
            }
        }
        
        #endregion Advanced Settings

        #region Advanced, File-Controlled Settings

        private bool _disableVideoFeedCloning;
        private bool _enableVideoDebugging;
        private bool _disableRequiredFieldValidation;
        private bool _useAdodc;
        private bool _disableStatusBarUpdates;
        private ValidVideoRenderers _videoRenderer;
        
        public bool DisableVideoFeedCloning {
            get {
                return _disableVideoFeedCloning;
            }
            set {
                _disableVideoFeedCloning = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("DisableVideoFeedCloning"));
            }
        }
        public bool EnableVideoDebugging {
            get {
                return _enableVideoDebugging;
            }
            set {
                _enableVideoDebugging = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("EnableVideoDebugging"));
            }
        }
        public bool DisableRequiredFieldValidation {
            get {
                return _disableRequiredFieldValidation;
            }
            set {
                _disableRequiredFieldValidation = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("DisableRequiredFieldValidation"));
            }
        }
        public bool UseAdodc {
            get {
                return _useAdodc;
            }
            set {
                _useAdodc = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("UseAdodc"));
            }
        }
        public bool DisableStatusBarUpdates {
            get {
                return _disableStatusBarUpdates;
            }
            set {
                _disableStatusBarUpdates = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("DisableStatusBarUpdates"));
            }
        }
        public ValidVideoRenderers VideoRenderer {
            get {
                return _videoRenderer;
            }
            set {
                _videoRenderer = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("VideoRenderer"));
            }
        }

        #endregion Advanced, File-Controlled Settings
        
        #region Events

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion Events

        #region Constructors

        public ITpipesSettingsObj(string inputOleDbConnString = @"Provider = Microsoft.Jet.OLEDB.4.0; Data Source = C:\Program Files\InspectIT\setup.mdb") {

            if (string.IsNullOrWhiteSpace(inputOleDbConnString)) {
                throw new Exception("Connection String passed to settings constructor cannot be null or empty.");
            }

            this._globalSettingsConnString = inputOleDbConnString;

            using (OleDbConnection curOleConnection = new OleDbConnection(inputOleDbConnString))
            using (OleDbCommand curOleCommand = curOleConnection.CreateCommand())
            using (DataTable tempSettingsTable = new DataTable()) {

                curOleConnection.Open();

                curOleCommand.CommandText = "SELECT * FROM S";

                tempSettingsTable.Load(curOleCommand.ExecuteReader());

                if (tempSettingsTable.Rows.Count == 0) {

                    //closing and reopening connection to make sure that the new row is created before selecting it...Had exceptions being thrown if I didn't close and reopen the connection. Yeah. I thought it was bizarre too.
                    curOleConnection.Close();

                    createNewSettingsRow(inputOleDbConnString);

                    curOleConnection.Open();
                    tempSettingsTable.Load(curOleCommand.ExecuteReader());
                }

                populateCurrentSettings(tempSettingsTable.Rows[0]);
            }
        }

        #endregion Constructors

        #region Private Functions

        private static void createNewSettingsRow(string oleConnString) {

            try {

                using (OleDbConnection curOleConn = new OleDbConnection(oleConnString))
                using (OleDbCommand newSettingsRowCommand = curOleConn.CreateCommand()) {

                    curOleConn.Open();

                    newSettingsRowCommand.CommandText =
                        "INSERT INTO `S` (" +
                        "[Video_Quality], " +
                        "[Counter_Calibration], " +
                        "[Speech_Speed], " +
                        "[P_ID], " +
                        "[Files_DefaultProjectPath], " +
                        "[Misc_SplashScreen], " +
                        "[Video_WriteData], " +
                        "[Video_Format], " +
                        "[Counter_Type], " +
                        "[Misc_Backup_Files], " +
                        "[Misc_DBUseADODC]) " +
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

                    newSettingsRowCommand.Parameters.AddWithValue("Video_Quality", "CD");
                    newSettingsRowCommand.Parameters.AddWithValue("Counter_Calibration", 1.0);
                    newSettingsRowCommand.Parameters.AddWithValue("Speech_Speed", "Normal");
                    newSettingsRowCommand.Parameters.AddWithValue("P_ID", 0);
                    newSettingsRowCommand.Parameters.AddWithValue("Files_DefaultProjectPath", @"C:\IT Projects");
                    newSettingsRowCommand.Parameters.AddWithValue("Misc_SplashScreen", 0);
                    newSettingsRowCommand.Parameters.AddWithValue("Video_WriteData", 0);
                    newSettingsRowCommand.Parameters.AddWithValue("Video_Format", "WMV");
                    newSettingsRowCommand.Parameters.AddWithValue("Counter_Type", "USDigital");
                    newSettingsRowCommand.Parameters.AddWithValue("Misc_Backup_Files", -1);
                    newSettingsRowCommand.Parameters.AddWithValue("Misc_DBUseADODC", -1);

                    newSettingsRowCommand.ExecuteNonQuery();
                }
            }
            catch {
                throw;
            }
        }

        private void populateCurrentSettings(DataRow settingsRow) {
            try {



                #region Video Capture Settings

                this.AforgeVideoInputSources = new FilterInfoCollection(AForge.Video.DirectShow.FilterCategory.VideoInputDevice);
                this.VideoCaptureProfiles = getVideoCaptureProfiles(_globalSettingsConnString);
                this.setCurVidCapProfile(settingsRow);
                this.SelectedItVideoCapturePin = getCurrentVideoCapturePin(settingsRow);
                setCurrentVideoCaptureDevice(settingsRow);
                this.CurAforgeCaptureDevice = getCurrentVideoCaptureDevice(settingsRow);

                #endregion

                #region Overlay Settings

                this.DriverLoaderObj = new DriverLoader();
                this.OverlaySettingsObj = OverlaySettings.getFullyPopulatedOverlaySettingsObject(this._globalSettingsConnString);
                setActiveCounterProfile(settingsRow);

                #endregion

                #region TTS Settings

                populateInstalledTtsVoices();

                this.TtsEnabled = (bool)settingsRow["Speech_Enabled"];

                this.TtsSpeed = getTtsSpeedFromString(settingsRow["Speech_Speed"].ToString());

                this.TtsVoice = getTtsNameFromDescription(settingsRow["Speech_Voice"] == DBNull.Value ? "" : (string)settingsRow["Speech_Voice"]);

                this.TtsTestText = "This is a test for ITpipes Text-To-Speech Functionality";

                #endregion

                #region Grid Display Settings

                this.TextCase = getTextCaseEnumFromString(settingsRow["Misc_TextCase"] == DBNull.Value ? "" : (string)settingsRow["Misc_TextCase"]);
                this.TextDisplayType = getTextDisplayEnumFromString(settingsRow["Text_DisplayType"] == DBNull.Value ? "" : (string)settingsRow["Text_DisplayType"]);
                this.ShowGridNumbers = (bool)settingsRow["Misc_GridNumber"];

                #endregion

                #region Miscellaneous Settings

                this.AutoRefreshAFS = (bool)settingsRow["Files_AutoRefresh_AFS"];
                this.BackupProjects = (bool)settingsRow["Misc_Backup_Files"];
                this.AutosaveReports = (bool)settingsRow["Reporting_AutoSave"];
                this.EnableRemarksPersistence = isRemarksPersistenceEnabled();
                this.EnableSplashScreen = (bool)settingsRow["Misc_SplashScreen"];
                this.DefaultProjectPath = settingsRow["Files_DefaultProjectPath"] == DBNull.Value ? @"C:\IT Projects" : (string)settingsRow["Files_DefaultProjectPath"];

                #endregion

                #region Advanced Settings

                this.DisableVideoFeedCloning = isVideoFeedCloningDisabled();
                this.EnableVideoDebugging = isVideoDebuggingEnabled();
                this.DisableRequiredFieldValidation = settingsRow.Table.Columns.Contains("Misc_CompliancePrep") ? (bool)settingsRow["Misc_CompliancePrep"] : false;
                this.UseAdodc = (bool)settingsRow["Misc_DBUseADODC"];
                this.DisableStatusBarUpdates = isStatusBarUpdatesDisabled();
                this.VideoRenderer = getEnabledVideoRenderer();

                #endregion
            }
            catch {
                throw;
            }
            }

        private void setActiveCounterProfile(DataRow curSettingsRow) {

            if (this.OverlaySettingsObj.IsOverlayEnabled == false ||
                    ((curSettingsRow.Table.Columns.Contains("CounterV2_Type") && curSettingsRow["CounterV2_Type"] == DBNull.Value) &&
                    curSettingsRow["Counter_Type"] == DBNull.Value)) {
                ActiveCounterProfile = null;
                return;
            }

            if ((bool)curSettingsRow["CounterV2_Enabled"]) {
                foreach (OverlayControlClass curControl in OverlaySettingsObj.ControlsAvailable) {
                    if (curControl.OverlayName == (string)curSettingsRow["CounterV2_Type"]) {
                        this.ActiveCounterProfile = curControl;
                        return;
                    }
                }
            }
            else {
                foreach (OverlayControlClass curControl in OverlaySettingsObj.ControlsAvailable) {
                    if (curControl.OverlayName == (string)curSettingsRow["Counter_Type"]) {
                        this.ActiveCounterProfile = curControl;
                        return;
                    }
                }
            }
        }

        private bool isStatusBarUpdatesDisabled() {
            return File.Exists(@"C:\Program Files\InspectIT\turnstatusbar.off");
        }

        private void setCurVidCapProfile(DataRow curSettingsRow) {

            if (this.VideoCaptureProfiles != null &&
                this.VideoCaptureProfiles.Count != 0) {


                if (curSettingsRow["Video_Quality"] == DBNull.Value ||
                    curSettingsRow["Video_Format"] == DBNull.Value) {
                    CurVidCapProfile = null;
                    return;
                }

                videoCaptureResolutions activeResolution = getVidCapResolutionFromString((string)curSettingsRow["Video_Quality"]);
                videoCaptureFormats curVidCapFormat = getVideoFormatEnumFromString((string)curSettingsRow["Video_Format"]);

                foreach (VideoCaptureProfile curProfile in this.VideoCaptureProfiles) {

                    if (curProfile.VideoCaptureFormat == curVidCapFormat &&
                        curProfile.VideoCaptureResolution == activeResolution) {

                        this.CurVidCapProfile = curProfile;
                        return;
                    }
                }
            }
        }

        private VideoCapturePin getCurrentVideoCapturePin(DataRow curSettingsRow) {

            if (curSettingsRow["Video_CapturePin"] == DBNull.Value || curSettingsRow["Video_CapturePin"].ToString() == string.Empty) {
                return SharedEnums.VideoCapturePin.Composite;
            }

            return VideoCapturePin.Svideo;
        }

        private void setCurrentVideoCaptureDevice(DataRow curSettingsRow) {

            try {
                if (curSettingsRow["Video_CaptureDevice"] == DBNull.Value || curSettingsRow["Video_CaptureDevice"].ToString() == string.Empty) {
                    SelectedAforgeCaptureDevice = null;
                    return;
                }

                string curVidCapDevice = (string)curSettingsRow["Video_CaptureDevice"];

                for (int i = 0; i < this.AforgeVideoInputSources.Count; i++) {
                    if (this.AforgeVideoInputSources[i].Name == curVidCapDevice) {
                        SelectedAforgeCaptureDevice = this.AforgeVideoInputSources[i];
                    }
                }
            }
            catch {
                throw;
            }
        }

        private VideoCaptureDevice getCurrentVideoCaptureDevice(DataRow curSettingsRow) {
            try {


                if (curSettingsRow["Video_CaptureDevice"] == DBNull.Value || curSettingsRow["Video_CaptureDevice"].ToString() == string.Empty) {
                    return null;
                }

                string curVidCapDevice = (string)curSettingsRow["Video_CaptureDevice"];

                for (int i = 0; i < this.AforgeVideoInputSources.Count; i++) {
                    if (this.AforgeVideoInputSources[i].Name == curVidCapDevice) {
                        return new VideoCaptureDevice(this.AforgeVideoInputSources[i].MonikerString);
                    }
                }
            }
            catch {
                throw;
            }

            return null;
        }

        private void populateInstalledTtsVoices() {

            var voicesList = CurSpeechSynthObject.GetInstalledVoices();
            SpeechSythVoices = new string[voicesList.Count];

            for (int i = 0; i < voicesList.Count; i++) {
                SpeechSythVoices[i] = voicesList[i].VoiceInfo.Name;
            }

        }

        private TextCases getTextCaseEnumFromString(string input) {

            if (string.IsNullOrWhiteSpace(input)) {

                return TextCases.No_Restriction;
            }

            switch (input.ToLower()) {
                case ("n/a"):
                    {
                        return TextCases.No_Restriction;
                    }
                case ("upper"):
                    {
                        return TextCases.UPPER;
                    }
                case ("lower"):
                    {
                        return TextCases.lower;
                    }
                case ("proper"):
                    {
                        return TextCases.Proper;
                    }
                default:
                    {
                        return TextCases.No_Restriction;
                    }
            }

        }

        private string getTextCaseStringFromEnum(TextCases input) {
            if (input == TextCases.lower) {
                return "Lower";
            }
            else if (input == TextCases.Proper) {
                return "Proper";
            }
            else if (input == TextCases.UPPER) {
                return "Upper";
            }

            return "N/A";
        }

        private TextDisplays getTextDisplayEnumFromString(string input) {

            if (string.IsNullOrWhiteSpace(input)) {

                return TextDisplays.No_Restriction;
            }

            switch (input.ToLower()) {                
                case ("code"):
                    {
                        return TextDisplays.Code;
                    }
                case ("description"):
                    {
                        return TextDisplays.Description;
                    }
                case ("code and description"):
                    {
                        return TextDisplays.Code_And_Description;
                    }
                case ("alias"):
                    {
                        return TextDisplays.Alias;
                    }
                default:
                    {
                        return TextDisplays.No_Restriction;
                    }
            }
        }

        private string getTextDisplayStringFromEnum(TextDisplays input) {
            if (input == TextDisplays.Alias) {
                return "Alias";
            }
            else if (input == TextDisplays.Code) {
                return "Code";
            }
            else if (input == TextDisplays.Code_And_Description) {
                return "Code AND Description";
            }
            else if (input == TextDisplays.Description || input == TextDisplays.No_Restriction) {
                return "Description";
            }

            return "Description";
        }

        private TtsSpeeds getTtsSpeedFromString(string input) {

            if (string.IsNullOrWhiteSpace(input)) {

                return TtsSpeeds.Normal;
            }

            switch (input) {
                case (""):
                    {
                        return TtsSpeeds.Normal;
                    }
                case ("Fast"):
                    {
                        return TtsSpeeds.Fast;
                    }
                case ("Normal"):
                    {
                        return TtsSpeeds.Normal;
                    }
                case ("Slow"):
                    {
                        return TtsSpeeds.Slow;
                    }
                default:
                    {
                        return TtsSpeeds.Normal;
                    }
            }
        }

        private bool? getBooleanFromString(string input) {

            if (string.IsNullOrWhiteSpace(input)) {
                throw new Exception("Argument passed to getBooleanFromString function was null or empty");
            }

            switch (input.ToLower()) {

                case ("y"):
                case ("yes"):
                case ("t"):
                case ("true"):
                    {
                        return true;
                    }
                case ("n"):
                case ("no"):
                case ("f"):
                case ("false"):
                    {
                        return false;
                    }
                default:
                    {
                        return null;
                    }
            }
        }

        private bool isRemarksPersistenceEnabled() {
            return System.IO.File.Exists(@"C:\Program Files\InspectIT\turnremarks.on");
        }

        private bool isVideoFeedCloningDisabled() {
            return System.IO.File.Exists(@"C:\Program Files\InspectIT\turnoverlay.off");
        }

        private bool isVideoDebuggingEnabled() {
            return System.IO.File.Exists(@"C:\Program Files\InspectIT\turndebugvideo.on");
        }

        private ValidVideoRenderers getEnabledVideoRenderer() {

            if (File.Exists(@"C:\Program Files\InspectIT\videogdi.on")) {
                return ValidVideoRenderers.GDI;
            }
            else if (File.Exists(@"C:\Program Files\InspectIT\videovmr7.on")) {
                return ValidVideoRenderers.VMR7;
            }
            else if (File.Exists(@"C:\Program Files\InspectIT\videovmr9.on")) {
                return ValidVideoRenderers.VMR9;
            }
            else if (File.Exists(@"C:\Program Files\InspectIT\videooverlay.on")) {
                return ValidVideoRenderers.OverlaySurface;
            }
            else {
                return ValidVideoRenderers.Default;
            }
        }

        private ObservableCollection<VideoCaptureProfile> getVideoCaptureProfiles(string oleConnString) {
            try {

                ObservableCollection<VideoCaptureProfile> returnOC = new ObservableCollection<VideoCaptureProfile>();

                using (OleDbConnection curOleConn = new OleDbConnection(oleConnString))
                using (OleDbCommand curOleCommand = curOleConn.CreateCommand())
                using (DataTable vidProfileTable = new DataTable()) {

                    curOleConn.Open();

                    curOleCommand.CommandText = "SELECT ProfileName, Bitrate, VideoType, Quality, VBR FROM R_Video_Profiles";
                    vidProfileTable.Load(curOleCommand.ExecuteReader());

                    curOleConn.Close();

                    foreach (videoCaptureFormats curVideoFormat in this.ValidVideoCaptureFormats) {
                        DataRow[] matchingCdProfiles = vidProfileTable.Select(string.Format("ProfileName = '{0} (CD)'", getVideoFormatStringFromEnum(curVideoFormat)));

                        if (matchingCdProfiles.Length == 0) {
                            continue;
                        }
                        else if (matchingCdProfiles.Length > 1) {
                            throw new Exception(string.Format("Multiple CD profiles defined for video profile:\n{0}", getVideoFormatStringFromEnum(curVideoFormat)));
                        }

                        DataRow curRow = matchingCdProfiles[0];

                        string _videoProfileName =
                            curRow["ProfileName"].ToString()
                            .Replace("(CD)", "")
                            .Replace("(DVD)", "")
                            .Trim();

                        returnOC.Add(new VideoCaptureProfile(_videoProfileName, (string)curRow["VideoType"], (bool)curRow["VBR"], (int)curRow["Bitrate"], (int)curRow["Quality"]));

                        DataRow[] matchingDVDProfiles = vidProfileTable.Select(string.Format("ProfileName = '{0} (DVD)'", getVideoFormatStringFromEnum(curVideoFormat)));

                        if (matchingDVDProfiles.Length == 0) {
                            continue;
                        }
                        else if (matchingDVDProfiles.Length > 1) {
                            throw new Exception(string.Format("Multiple DVD profiles defined for video profile:\n{0}", getVideoFormatStringFromEnum(curVideoFormat)));
                        }

                        curRow = matchingDVDProfiles[0];

                        _videoProfileName =
                            curRow["ProfileName"].ToString()
                            .Replace("(CD)", "")
                            .Replace("(DVD)", "")
                            .Trim();

                        returnOC.Add(new VideoCaptureProfile(_videoProfileName, (string)curRow["VideoType"], (bool)curRow["VBR"], (int)curRow["Bitrate"], (int)curRow["Quality"]));
                    }

                }

                return returnOC;
            }
            catch {
                throw;
            }
        }

        private videoCaptureResolutions getVidCapResolutionFromString(string input) {
            if (input == null) {
                return videoCaptureResolutions.CD;
            }

            if (input.ToLower() == "cd") {
                return videoCaptureResolutions.CD;
            }

            return videoCaptureResolutions.DVD;
        }

        private string getVidCapResolutionStringFromEnum(videoCaptureResolutions input) {
            if (input == videoCaptureResolutions.CD) {
                return "CD";
            }
            else if (input == videoCaptureResolutions.DVD) {
                return "DVD";
            }
            else {
                throw new Exception("Received invalid enum value in getVidCapResolutionsStringFromEnum: " + input.ToString());
            }
        }

        private videoCaptureFormats getVideoFormatEnumFromString(string input) {
            if (string.IsNullOrEmpty(input)) {
                return videoCaptureFormats.WMV7;
            }

            switch (input.ToLower()) {
                case ("wmv"):
                    {
                        return videoCaptureFormats.WMV7;
                    }
                case ("wmv9"):
                    {
                        return videoCaptureFormats.WMV9;
                    }
                case ("mpeg 1"):
                    {
                        return videoCaptureFormats.MPEG1;
                    }
                case ("mpeg 2"):
                    {
                        return videoCaptureFormats.MPEG2;
                    }
                case ("mpeg 4"):
                    {
                        return videoCaptureFormats.MPEG4;
                    }
                default:
                    {
                        return videoCaptureFormats.WMV7;
                    }
            }
        }

        private string getVideoFormatStringFromEnum(videoCaptureFormats inputFormat) {

            switch (inputFormat) {

                case (videoCaptureFormats.WMV7):
                    {
                        return "WMV";
                    }
                case (videoCaptureFormats.WMV9):
                    {
                        return "WMV9";
                    }
                case (videoCaptureFormats.MPEG1):
                    {
                        return "MPEG 1";
                    }
                case (videoCaptureFormats.MPEG2):
                    {
                        return "MPEG 2";
                    }
                case (videoCaptureFormats.MPEG4):
                    {
                        return "MPEG 4";
                    }
                default:
                    {
                        throw new Exception("You broke the videoCaptureFormats enum. You passed a value other than a possible Enum value. Wow.");
                    }

            }

        }

        private string getTtsSpeedStringFromEnum(TtsSpeeds input) {
            if (input == TtsSpeeds.Fast) {
                return "Fast";
            }
            else if (input == TtsSpeeds.Normal) {
                return "Normal";
            }
            else if (input == TtsSpeeds.Slow) {
                return "Slow";
            }

            return "Normal";
        }

        private void saveFileControlledSettings() {

            try {


                #region Enable Remarks Persistence

                if (this.EnableRemarksPersistence == true) {
                    if (File.Exists(@"C:\Program Files\InspectIT\turnremarks.on") == false) {
                        File.Create(@"C:\Program Files\InspectIT\turnremarks.on").Close();
                    }
                }
                else {
                    if (File.Exists(@"C:\Program Files\InspectIT\turnremarks.on")) {
                        File.Delete(@"C:\Program Files\InspectIT\turnremarks.on");
                    }
                }

                #endregion Enable Remarks Persistence

                #region Disable Video Feed Cloning

                if (this.DisableVideoFeedCloning == true) {
                    if (File.Exists(@"C:\Program Files\InspectIT\turnoverlay.off") == false) {
                        File.Create(@"C:\Program Files\InspectIT\turnoverlay.off").Close();
                    }
                }
                else {
                    if (File.Exists(@"C:\Program Files\InspectIT\turnoverlay.off")) {
                        File.Delete(@"C:\Program Files\InspectIT\turnoverlay.off");
                    }
                }

                #endregion Disable Video Feed Cloning

                #region Enable Video Capture Debugging

                if (this.EnableVideoDebugging == true) {
                    if (File.Exists(@"C:\Program Files\InspectIT\turndebugvideo.on") == false) {
                        File.Create(@"C:\Program Files\InspectIT\turndebugvideo.on").Close();
                    }
                }
                else {
                    if (File.Exists(@"C:\Program Files\InspectIT\turndebugvideo.on")) {
                        File.Delete(@"C:\Program Files\InspectIT\turndebugvideo.on");
                    }
                }

                #endregion Enable Video Capture Debugging

                #region Status Bar Updates Disabled

                if (this.DisableStatusBarUpdates == true) {
                    if (File.Exists(@"C:\Program Files\InspectIT\turnstatusbar.off") == false) {
                        File.Create(@"C:\Program Files\InspectIT\turnstatusbar.off").Close();
                    }
                }
                else {
                    if (File.Exists(@"C:\Program Files\InspectIT\turnstatusbar.off")) {
                        File.Delete(@"C:\Program Files\InspectIT\turnstatusbar.off");
                    }
                }

                #endregion Status Bar Updates Disabled

                #region Video Renderer Settings

                if (VideoRenderer == ValidVideoRenderers.Default) {

                    //if (File.Exists(@"C:\Program Files\InspectIT\videogdi.on") == false) {
                    //    File.Create(@"C:\Program Files\InspectIT\videogdi.on").Close();
                    //}
                    if (File.Exists(@"C:\Program Files\InspectIT\videogdi.on")) { File.Delete(@"C:\Program Files\InspectIT\videogdi.on"); }
                    if (File.Exists(@"C:\Program Files\InspectIT\videovmr7.on")) { File.Delete(@"C:\Program Files\InspectIT\videovmr7.on"); }
                    if (File.Exists(@"C:\Program Files\InspectIT\videovmr9.on")) { File.Delete(@"C:\Program Files\InspectIT\videovmr9.on"); }
                    if (File.Exists(@"C:\Program Files\InspectIT\videooverlay.on")) { File.Delete(@"C:\Program Files\InspectIT\videooverlay.on"); }
                }
                else if (VideoRenderer == ValidVideoRenderers.GDI) {
                    if (File.Exists(@"C:\Program Files\InspectIT\videogdi.on") == false) {
                        File.Create(@"C:\Program Files\InspectIT\videogdi.on").Close();
                    }
                    if (File.Exists(@"C:\Program Files\InspectIT\videovmr7.on")) { File.Delete(@"C:\Program Files\InspectIT\videovmr7.on"); }
                    if (File.Exists(@"C:\Program Files\InspectIT\videovmr9.on")) { File.Delete(@"C:\Program Files\InspectIT\videovmr9.on"); }
                    if (File.Exists(@"C:\Program Files\InspectIT\videooverlay.on")) { File.Delete(@"C:\Program Files\InspectIT\videooverlay.on"); }
                }
                else if (VideoRenderer == ValidVideoRenderers.OverlaySurface) {
                    if (File.Exists(@"C:\Program Files\InspectIT\videooverlay.on") == false) {
                        File.Create(@"C:\Program Files\InspectIT\videooverlay.on").Close();
                    }
                    if (File.Exists(@"C:\Program Files\InspectIT\videogdi.on")) { File.Delete(@"C:\Program Files\InspectIT\videogdi.on"); }
                    if (File.Exists(@"C:\Program Files\InspectIT\videovmr7.on")) { File.Delete(@"C:\Program Files\InspectIT\videovmr7.on"); }
                    if (File.Exists(@"C:\Program Files\InspectIT\videovmr9.on")) { File.Delete(@"C:\Program Files\InspectIT\videovmr9.on"); }
                }
                else if (VideoRenderer == ValidVideoRenderers.VMR7) {
                    if (File.Exists(@"C:\Program Files\InspectIT\videovmr7.on") == false) {
                        File.Create(@"C:\Program Files\InspectIT\videovmr7.on").Close();
                    }
                    if (File.Exists(@"C:\Program Files\InspectIT\videogdi.on")) { File.Delete(@"C:\Program Files\InspectIT\videogdi.on"); }
                    if (File.Exists(@"C:\Program Files\InspectIT\videovmr9.on")) { File.Delete(@"C:\Program Files\InspectIT\videovmr9.on"); }
                    if (File.Exists(@"C:\Program Files\InspectIT\videooverlay.on")) { File.Delete(@"C:\Program Files\InspectIT\videooverlay.on"); }
                }
                else if (VideoRenderer == ValidVideoRenderers.VMR9) {
                    if (File.Exists(@"C:\Program Files\InspectIT\videovmr9.on") == false) {
                        File.Create(@"C:\Program Files\InspectIT\videovmr9.on").Close();
                    }
                    if (File.Exists(@"C:\Program Files\InspectIT\videogdi.on")) { File.Delete(@"C:\Program Files\InspectIT\videogdi.on"); }
                    if (File.Exists(@"C:\Program Files\InspectIT\videovmr7.on")) { File.Delete(@"C:\Program Files\InspectIT\videovmr7.on"); }
                    if (File.Exists(@"C:\Program Files\InspectIT\videooverlay.on")) { File.Delete(@"C:\Program Files\InspectIT\videooverlay.on"); }
                }


                #endregion
            }
            catch {

            }
        }

        private void saveCounterSetting(OleDbConnection curConn, OleDbCommand curCommand) {

            if (curConn.State != ConnectionState.Open) {
                curConn.Open();
            }

            //setting all overlay settings to the default "everything off" settings

            curCommand.CommandText =
                "UPDATE S SET " +
                "Counter_Enabled = 0, " +
                "Counter_Type = 'USDigital', " +
                "IBAK_Counter_Enabled = 0, " +
                "IBAK_Counter_Port = NULL, " +
                "Aries_Counter_Enabled = 0, " +
                "Aries_Counter2_Enabled = 0, " +
                "Aries_Overlay_Port = NULL, " +
                "Aries_Overlay_Enabled = 0, " +
                "RST_Counter_Port = NULL, " +
                "CounterV2_Enabled = 0";

            //The command will throw an exception if the ITpipes v2x fields don't exist
            try {
                curCommand.ExecuteNonQuery();
            }
            catch {

                curCommand.CommandText =
                "UPDATE S SET " +
                "Counter_Enabled = 0, " +
                "Counter_Type = 'USDigital', " +
                "IBAK_Counter_Enabled = 0, " +
                "IBAK_Counter_Port = NULL, " +
                "Aries_Counter_Enabled = 0, " +
                "Aries_Counter2_Enabled = 0, " +
                "Aries_Overlay_Port = NULL, " +
                "Aries_Overlay_Enabled = 0";

                try {
                    curCommand.ExecuteNonQuery();
                }
                catch (Exception ex) {
                    throw new Exception(string.Format("Unable to save changes to setup.mdb.\nException:{0}", ex.Message), ex);
                }
            }

            if (OverlaySettingsObj.IsOverlayEnabled == false ||
                ActiveCounterProfile == null) {

                return; //the command above disables all overlay settings, so our work here is done.
            }

            //If we get to this point the overlay is enabled and the port is not allowed to be null
            if (OverlaySettingsObj.OverlayPort == null) {
                OverlaySettingsObj.OverlayPort = 1;
            }

            if (ActiveCounterProfile.UsesV2DriverLoader == true &&
                ActiveCounterProfile.ForceUseV1Control == false) {

                writeSetupMdbField(curConn, curCommand, "CounterV2_Port", OverlaySettingsObj.OverlayPort);
                writeSetupMdbField(curConn, curCommand, "CounterV2_Type", ActiveCounterProfile.OverlayName);
                writeSetupMdbField(curConn, curCommand, "CounterV2_Enabled", true);
            }

            else {

                string v1OverlayName = "";

                if (ActiveCounterProfile.ForceUseV1Control == true &&
                    ActiveCounterProfile.UsesV2DriverLoader == true) {

                    v1OverlayName = ActiveCounterProfile.AlternateOverlayControlName;
                }
                else {
                    v1OverlayName = ActiveCounterProfile.OverlayName;
                }


                if (v1OverlayName.Contains("Aries")) {
                    writeV1AriesSettings(curConn, curCommand, v1OverlayName);
                }
                else if (v1OverlayName == "IBAK") {
                    writeV1IbakSettings(curConn, curCommand, v1OverlayName);
                }
                else if (v1OverlayName == "IPEK_DE03SW") {
                    writeV1IpekSettings(curConn, curCommand, v1OverlayName);
                }
                else if (v1OverlayName == "BOB4") {
                    writeV1Bob4Settings(curConn, curCommand, v1OverlayName);
                }
                else {
                    writeV1UsDigitalSettings(curConn, curCommand, v1OverlayName);
                }
            }
        }

        private void writeV1UsDigitalSettings(OleDbConnection curConn, OleDbCommand curCommand, string overlayType) {
            writeSetupMdbField(curConn, curCommand, "Counter_Enabled", true);
            writeSetupMdbField(curConn, curCommand, "Counter_Port", OverlaySettingsObj.OverlayPort == null ? (object)DBNull.Value : OverlaySettingsObj.OverlayPort);
            writeSetupMdbField(curConn, curCommand, "Counter_Type", overlayType);
        }

        private void writeV1Bob4Settings(OleDbConnection curConn, OleDbCommand curCommand, string overlayType) {

            //V1 Bob4 control reads from the RST Counter Port. Don't ask me why--I just work here.
            writeSetupMdbField(curConn, curCommand, "Counter_Enabled", true);
            writeSetupMdbField(curConn, curCommand, "Counter_Port", OverlaySettingsObj.OverlayPort == null ? (object)DBNull.Value : OverlaySettingsObj.OverlayPort);
            writeSetupMdbField(curConn, curCommand, "Counter_Type", overlayType);
            writeSetupMdbField(curConn, curCommand, "RST_Counter_Port", OverlaySettingsObj.OverlayPort == null ? (object)DBNull.Value : OverlaySettingsObj.OverlayPort);

        }

        private void writeV1IpekSettings(OleDbConnection curConn, OleDbCommand curCommand, string overlayType) {

            //Not at all sure why, but the v1 Ipek control requires that the Aries_Overlay_Port be set. Probably just because Bob.
            writeSetupMdbField(curConn, curCommand, "Counter_Enabled", true);
            writeSetupMdbField(curConn, curCommand, "Counter_Port", OverlaySettingsObj.OverlayPort == null ? (object)DBNull.Value : OverlaySettingsObj.OverlayPort);
            writeSetupMdbField(curConn, curCommand, "Counter_Type", overlayType);
            writeSetupMdbField(curConn, curCommand, "Aries_Overlay_Port", OverlaySettingsObj.OverlayPort == null ? (object)DBNull.Value : OverlaySettingsObj.OverlayPort);
        }

        private void writeV1IbakSettings(OleDbConnection curConn, OleDbCommand curCommand, string overlayType) {
            writeSetupMdbField(curConn, curCommand, "Counter_Enabled", true);
            writeSetupMdbField(curConn, curCommand, "Counter_Port", OverlaySettingsObj.OverlayPort == null ? (object)DBNull.Value : OverlaySettingsObj.OverlayPort);
            writeSetupMdbField(curConn, curCommand, "Counter_Type", overlayType);
            writeSetupMdbField(curConn, curCommand, "IBAK_Counter_Port", OverlaySettingsObj.OverlayPort == null ? (object)DBNull.Value : OverlaySettingsObj.OverlayPort);
            writeSetupMdbField(curConn, curCommand, "IBAK_Counter_Enabled", true);
        }

        private void writeV1AriesSettings(OleDbConnection curConn, OleDbCommand curCommand, string overlayType) {
            writeSetupMdbField(curConn, curCommand, "Counter_Enabled", true);
            writeSetupMdbField(curConn, curCommand, "Counter_Port", OverlaySettingsObj.OverlayPort == null ? (object)DBNull.Value : OverlaySettingsObj.OverlayPort);
            writeSetupMdbField(curConn, curCommand, "Counter_Type", overlayType);
            writeSetupMdbField(curConn, curCommand, "Aries_Overlay_Port", OverlaySettingsObj.OverlayPort == null ? (object)DBNull.Value : OverlaySettingsObj.OverlayPort);
            writeSetupMdbField(curConn, curCommand, "Aries_Counter_Enabled", true);
        }

        private void saveVideoRecordingProfiles(OleDbConnection curConn, OleDbCommand curCommand) {

            if (curConn.State != ConnectionState.Open) {
                curConn.Open();
            }

            for (int i = 0; i < VideoCaptureProfiles.Count; i++) {

                VideoCaptureProfile curVidCapProfile = VideoCaptureProfiles[i];

                curCommand.Parameters.Clear();

                curCommand.CommandText = "UPDATE R_Video_Profiles SET Bitrate = ?, Quality = ?, VBR = ? WHERE ProfileName = ?";

                float newCBRBitrate = -1;
                int newCBRBitrateInt = -1;


                if (float.TryParse(curVidCapProfile.VideoCaptureConstantBitrate, out newCBRBitrate) == false) {
                    newCBRBitrateInt = curVidCapProfile.VideoCaptureResolution == videoCaptureResolutions.CD ? 1000000 : 7250000;
                }
                else {
                    newCBRBitrateInt = (int)(newCBRBitrate * 1000000);
                }

                curCommand.Parameters.AddWithValue("Bitrate", newCBRBitrateInt);
                curCommand.Parameters.AddWithValue("Quality", curVidCapProfile.QualityPercent);
                curCommand.Parameters.AddWithValue("VBR", curVidCapProfile.UseVariableBitrate);
                curCommand.Parameters.AddWithValue("ProfileName", curVidCapProfile.DatabaseSpecifiedName);

                curCommand.ExecuteNonQuery();

            }


        }

        private void writeSetupMdbField(OleDbConnection curConn, OleDbCommand curCommand, string fieldName, object value) {

            if (value == null || value.ToString() == "") {
                value = DBNull.Value;
            }


            if (curConn.State != ConnectionState.Open) {
                curConn.Open();
            }

            curCommand.Parameters.Clear();

            curCommand.CommandText = string.Format("UPDATE S SET {0} = ?", fieldName);
            curCommand.Parameters.AddWithValue("fieldName", value);

            curCommand.ExecuteNonQuery();
        }

        private string getTtsNameFromDescription(string input) {

            if (string.IsNullOrWhiteSpace(input)) {
                return string.Empty;
            }

            if (this.CurSpeechSynthObject != null) {

                var validVoices = this.CurSpeechSynthObject.GetInstalledVoices();

                foreach (var curVoice in validVoices) {

                    if (curVoice.VoiceInfo.Description == input) {
                        return curVoice.VoiceInfo.Name;
                    }
                }
            }

            return string.Empty;
        }

        private string getTtsDescriptionFromTtsName(string input) {

            var tempSpeechSynthObject = new System.Speech.Synthesis.SpeechSynthesizer();

            var validVoices = tempSpeechSynthObject.GetInstalledVoices();

            foreach (var curVoice in validVoices) {

                if (curVoice.VoiceInfo.Name == input) {
                    return curVoice.VoiceInfo.Description;
                }
            }

            return string.Empty;
        }

        private void OnPropertyChanged(object sender, PropertyChangedEventArgs e) {

            if (PropertyChanged != null) {
                PropertyChanged(sender, e);
            }
        }

        #endregion Private Functions
        
        #region Public Functions

        public void SaveSettings() {

            using (OleDbConnection curOleConn = new OleDbConnection(_globalSettingsConnString))
            using (OleDbCommand curOleCommand = curOleConn.CreateCommand()) {

                writeSetupMdbField(curOleConn, curOleCommand, "Video_Quality",          getVidCapResolutionStringFromEnum(CurVidCapProfile.VideoCaptureResolution));
                writeSetupMdbField(curOleConn, curOleCommand, "Speech_Enabled",         TtsEnabled);
                writeSetupMdbField(curOleConn, curOleCommand, "Speech_Voice",           getTtsDescriptionFromTtsName(TtsVoice));
                writeSetupMdbField(curOleConn, curOleCommand, "Counter_Calibration",    OverlaySettingsObj.MainlineCalibration);
                writeSetupMdbField(curOleConn, curOleCommand, "Speech_Speed",           getTtsSpeedStringFromEnum(TtsSpeed));
                writeSetupMdbField(curOleConn, curOleCommand, "Misc_TextCase",          getTextCaseStringFromEnum(TextCase));
                writeSetupMdbField(curOleConn, curOleCommand, "Reporting_AutoSave",     AutosaveReports);
                //writeSetupMdbField("Counter_PreSet") //not currently implemented.


                string captureDeviceName = "";

                if (AforgeVideoInputSources != null && CurAforgeCaptureDevice != null) {
                    captureDeviceName = SelectedAforgeCaptureDevice.Name;
                }

                writeSetupMdbField(curOleConn, curOleCommand, "Video_CaptureDevice",            captureDeviceName == "MV Demo Source" ? (object)DBNull.Value : captureDeviceName);
                writeSetupMdbField(curOleConn, curOleCommand, "Video_CapturePin",               SelectedItVideoCapturePin == VideoCapturePin.Svideo ? "svideo" : "");
                writeSetupMdbField(curOleConn, curOleCommand, "Files_DefaultProjectPath",       DefaultProjectPath);
                writeSetupMdbField(curOleConn, curOleCommand, "Text_DisplayType",               getTextDisplayStringFromEnum(TextDisplayType));
                writeSetupMdbField(curOleConn, curOleCommand, "Misc_GridNumber",                ShowGridNumbers);
                writeSetupMdbField(curOleConn, curOleCommand, "Misc_DBUseADODC",                UseAdodc);
                writeSetupMdbField(curOleConn, curOleCommand, "Misc_SplashScreen",              EnableSplashScreen);
                writeSetupMdbField(curOleConn, curOleCommand, "Video_WriteData",                false); //writedata can cause video merge/write issues. There is no use case for it being enabled in 18xx+ versions.
                writeSetupMdbField(curOleConn, curOleCommand, "Video_Format",                   getVideoFormatStringFromEnum(CurVidCapProfile.VideoCaptureFormat));
                writeSetupMdbField(curOleConn, curOleCommand, "Counter_Port",                   OverlaySettingsObj.OverlayPort);
                writeSetupMdbField(curOleConn, curOleCommand, "Counter_Lateral_Port",           OverlaySettingsObj.LateralPort);
                writeSetupMdbField(curOleConn, curOleCommand, "Counter_Lateral_Calibration",    OverlaySettingsObj.LateralCalibration);
                writeSetupMdbField(curOleConn, curOleCommand, "RST_Inclinometer_Port",          OverlaySettingsObj.InclinometerPort == null ? (object)DBNull.Value : OverlaySettingsObj.InclinometerPort);
                writeSetupMdbField(curOleConn, curOleCommand, "Misc_Backup_Files",              BackupProjects);
                writeSetupMdbField(curOleConn, curOleCommand, "Files_AutoRefresh_AFS",          AutoRefreshAFS);
                writeSetupMdbField(curOleConn, curOleCommand, "Misc_CompliancePrep",            DisableRequiredFieldValidation);

                saveCounterSetting(curOleConn, curOleCommand);
                saveFileControlledSettings();

                saveVideoRecordingProfiles(curOleConn, curOleCommand);
            }
        }

        public void Dispose() {

            //have to explicitly signal for close and null the capture device, otherwise the process stays open indefinitely.
            if (this.CurAforgeCaptureDevice != null) {

                this.CurAforgeCaptureDevice.SignalToStop();
                int waitedMs = 0;
                while (this.CurAforgeCaptureDevice.IsRunning) {
                    System.Threading.Thread.Sleep(5);
                    waitedMs += 5;

                    if (waitedMs > 250) {
                        break;
                    }
                }
                this.CurAforgeCaptureDevice = null;
            }

            if (this.AforgeVideoInputSources != null) {
                this.AforgeVideoInputSources = null;
            }
            if (this.CurAforgeCaptureDevice != null) {
                this.CurAforgeCaptureDevice = null;
            }
            if (this.CurSpeechSynthObject != null) {
                this.CurSpeechSynthObject.Dispose();
                this._curSpeechSynthObject = null;
            }

            if (this.DriverLoaderObj != null) {
                this.DriverLoaderObj.ClosePort();
                this.DriverLoaderObj.UnloadDriver();
                this.DriverLoaderObj = null;
            }

            if (this.OverlaySettingsObj != null) {
                this.OverlaySettingsObj.Dispose();
                this.OverlaySettingsObj = null;
            }

            if (this.VideoCaptureProfiles != null) {
                this.VideoCaptureProfiles = null;
            }
        }

        #endregion Public Functions
        
    }
}

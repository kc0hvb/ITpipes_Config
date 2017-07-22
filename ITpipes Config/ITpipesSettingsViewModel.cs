using AForge.Video.DirectShow;
using System;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Speech.Synthesis;

namespace ITpipes_Config {

    public class ITpipesSettingsViewModel : INotifyPropertyChanged, IDisposable {

        #region Read-Only Resources

        private static readonly string TTS_RECORDING_WARNING_TEXT =
            "To record Text-to-Speech audio onto your inspection videos you MUST have an audio recording device recording Windows audio output.\n\n" +
            "Your Windows audio configuration has been opened to the recording device configuration tab. To ensure that Text-to-speech records " +
            "properly you will need to ensure that your Windows default recording device is set to \"Stereo Mix\" or a third-party virtual audio device.\n\n" +
            "If you do not see \"Stereo Mix\" as a recording device option, right-click below the list of devices and show disabled devices. If you still do not " +
            "see the \"Stereo Mix\" recording device you will need to install a third-party virtual audio input device.\n\n" +
            "If your Windows default recording device (will have a checkmark on the icon) is set to anything with \"Microphone\" in the name you are NOT configured for Text-to-Speech recording\n\n" +
            "When your audio recording device is set up correctly you will see the green audio level bar next to the recording device move in time to the voice reading the test text.\n\n" +
            "If assistance is required to configure Text-to-Speech recording, please contact your ITpipes dealer or ITpipes support.";
        
        #endregion Read-Only Resources

        #region Private variables

        private Random _randGen = new Random(DateTime.Now.Millisecond);
        private MemoryStream _newFrameMS = new MemoryStream();
        private SpeechSynthesizer _testTtsSynthesizer = new SpeechSynthesizer();
        private ITpipesSettingsObj _settings;
        private bool _textToSpeechTestInstructionsHaveBeenOpened = false;
        private FilterInfo _selectedAforgeFilterInfo;
        private BitmapImage _aforgeFrameImage;
        private string[] _aforgeVideoPinsWithFriendlyNames;
        private string _selectedAforgeVideoPinFriendlyName;
        private AudioRecordingInstructionsWindow _audioRecordingInstructionsWindow;
        private FilterInfo _settingsSelectedFilterInfo;
        private VideoInput _selectedVideoInput;
        private SharedEnums.VideoCapturePin _settingsVideoCapturePin;

        #endregion

        #region Public Properties
        
        public ICommand OpenAForgeDevicePropertiesPage {
            get {
                return new GenericCommandCanAlwaysExecute(openDevicePropertyPage);
            }
        }
        public ICommand ScanSerialPortsForSelectedOverlay {
            get {
                return new GenericCommandCanAlwaysExecute(scanSerialPortsForValidDevice);
            }
        }
        public ICommand SetDefaultProjectPathCommand {
            get {
                return new GenericCommandCanAlwaysExecute(GetNewDefaultProjectPathFromFolderFindDlg);
            }
        }
        public ICommand ReadTtsTextCommand {
            get {
                return new GenericCommandCanAlwaysExecute(ReadTextWithTts);
            }
        }
        public ICommand OpenTTSRecordingInstructions {
            get {
                return new GenericCommandCanAlwaysExecute(openTTSRecordingInstructions);
            }
        }

        public string[] AForgeVideoPinsWithFriendlyNames {
            get {
                return _aforgeVideoPinsWithFriendlyNames;
            }

            set {
                _aforgeVideoPinsWithFriendlyNames = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("AForgeVideoPinsWithFriendlyNames"));
            }
        }
        public string SelectedAforgeVideoPinFriendlyName {
            get {
                return _selectedAforgeVideoPinFriendlyName;
            }
            set {
                _selectedAforgeVideoPinFriendlyName = value;
                
                setAforgeVideoCapabilies();

                OnPropertyChanged(this, new PropertyChangedEventArgs("SelectedAforgeVideoPinFriendlyName"));
            }
        }
        public ITpipesSettingsObj Settings {
            get {
                return _settings;
            }
            set {
                _settings = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("Settings"));
            }
        }
        public bool TextToSpeechTestInstructionsHaveBeenOpened {
            get {
                return _textToSpeechTestInstructionsHaveBeenOpened;
            }
            set {
                _textToSpeechTestInstructionsHaveBeenOpened = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("TextToSpeechTestInstructionsHaveBeenOpened"));
            }
        }
        public FilterInfo SelectedAforgeFilterInfo {
            get {
                return _selectedAforgeFilterInfo;
            }
            set {

                _selectedAforgeFilterInfo = value;

                ConnectToSelectedCaptureDevice();

                OnPropertyChanged(this, new PropertyChangedEventArgs("SelectedAforgeFilterInfo"));
            }
        }
        public BitmapImage AforgeFrameImage {
            get {
                return _aforgeFrameImage;
            }
            set {
                _aforgeFrameImage = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("AforgeFrameImage"));
            }
        }
        public FilterInfo SettingsSelectedFilterInfo {
            get {
                return _settingsSelectedFilterInfo;
            }
            set {
                
                if (Settings != null) {
                    Settings.SelectedAforgeCaptureDevice = value;
                    _settingsSelectedFilterInfo = Settings.SelectedAforgeCaptureDevice;
                }

                ConnectToSelectedCaptureDevice();

                OnPropertyChanged(this, new PropertyChangedEventArgs("SettingsSelectedFilterInfo"));
            }
        }
        public VideoInput SelectedVideoInput {
            get {
                return _selectedVideoInput;
            }
            set {
                _selectedVideoInput = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("SelectedVideoInput"));
            }
        }
        public SharedEnums.VideoCapturePin SettingsVideoCapturePin {
            get {
                return _settingsVideoCapturePin;
            }
            set {
                _settingsVideoCapturePin = value;
                if (Settings != null) {
                    Settings.SelectedItVideoCapturePin = value;
                }

                ConnectToSelectedCaptureDevice();

                OnPropertyChanged(this, new PropertyChangedEventArgs("SettingsVideoCapturePin"));
            }
        }

        #endregion Public Properties
        
        #region Events

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion Events

        #region Constructors

        public ITpipesSettingsViewModel() {

            Settings = new ITpipesSettingsObj();

            SettingsSelectedFilterInfo = Settings.SelectedAforgeCaptureDevice;
            getSelectedVideoInput();
        }

        #endregion Constructors

        #region Private Functions

        private void openTTSRecordingInstructions() {
            if (_audioRecordingInstructionsWindow == null || _audioRecordingInstructionsWindow.IsActive == false) {

                //wrapping the control panel call in a try/catch just in case for some reason the installation of Windows doesn't have/allow access to this function.

                try {
                    System.Diagnostics.Process.Start(@"C:\Windows\System32\rundll32.exe", "Shell32.dll,Control_RunDLL Mmsys.cpl,,1 Sound and Audio Device Properties");
                }
                catch { }

                _audioRecordingInstructionsWindow = new AudioRecordingInstructionsWindow();
                _audioRecordingInstructionsWindow.Show();
            }
        }

        private void ReadTextWithTts()
        {
            if (string.IsNullOrWhiteSpace(Settings.TtsTestText) == false &&
                string.IsNullOrWhiteSpace(Settings.TtsVoice) == false)
            {
                //I want to make sure that users know how to test for TTS being set up correctly, or at least have good direction:

                if (TextToSpeechTestInstructionsHaveBeenOpened == false)
                {
                    try
                    {
                        System.Diagnostics.Process.Start(@"C:\Windows\System32\rundll32.exe", "Shell32.dll,Control_RunDLL Mmsys.cpl,,1 Sound and Audio Device Properties");
                    }
                    catch
                    {

                    }

                    //MessageBox.Show(TTS_RECORDING_WARNING_TEXT, "Important!", MessageBoxButton.OK, MessageBoxImage.Warning);
                    TextToSpeechTestInstructionsHaveBeenOpened = true;
                }

                if (_testTtsSynthesizer.State == SynthesizerState.Speaking)
                {
                    _testTtsSynthesizer.SpeakAsyncCancelAll();
                }

                _testTtsSynthesizer.SelectVoice(Settings.TtsVoice);
                _testTtsSynthesizer.SpeakAsync(Settings.TtsTestText);
            }
        }
        
        private void scanSerialPortsForValidDevice() {

            int? portAtStartOfScan = this.Settings.OverlaySettingsObj.OverlayPort;
            this.Settings.OverlaySettingsObj.OverlayPort = null;
            Settings.OverlaySettingsObj.ScanSerialPortsForV2Overlay(Settings.ActiveCounterProfile);

            if (this.Settings.OverlaySettingsObj.OverlayPort == null) {
                this.Settings.OverlaySettingsObj.OverlayPort = portAtStartOfScan;
                MessageBox.Show("No device of the selected type responded to requests for current distance. Please verify the device type is correct and the device is powered on.", "No Response");
            }
        }

        private void GetNewDefaultProjectPathFromFolderFindDlg() {

            var selectFolder = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.DialogResult result = selectFolder.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK && System.IO.Directory.Exists(selectFolder.SelectedPath)) {
                this.Settings.DefaultProjectPath = selectFolder.SelectedPath;
            }
        }

        private void openDevicePropertyPage() {
            if (this.Settings != null && this.Settings.CurAforgeCaptureDevice != null) {
                this.Settings.CurAforgeCaptureDevice.DisplayPropertyPage(IntPtr.Zero);
            }
        }
        
        private void ConnectToSelectedCaptureDevice() {

            if (Settings.SelectedAforgeCaptureDevice == null) {
                return;
            }

            if (Settings.CurAforgeCaptureDevice != null) {
                Settings.CurAforgeCaptureDevice.SignalToStop();
                Settings.CurAforgeCaptureDevice.NewFrame -= newAforgeFrameHandler;
            }

            Settings.CurAforgeCaptureDevice = new AForge.Video.DirectShow.VideoCaptureDevice(Settings.SelectedAforgeCaptureDevice.MonikerString);

            //if the memorystream isn't reinitialized it will interlace the graphics of the previous cam with frames of the new cam.
            _newFrameMS = new System.IO.MemoryStream();
            AforgeFrameImage = new BitmapImage();

            Settings.CurAforgeCaptureDevice = new VideoCaptureDevice(Settings.SelectedAforgeCaptureDevice.MonikerString);


            string[] videoPinsWithFriendlyNames = new string[Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs.Length];


            for (int i = 0; i < videoPinsWithFriendlyNames.Length; i++) {
                if (Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs[i].Type == AForge.Video.DirectShow.PhysicalConnectorType.VideoSVideo) {
                    videoPinsWithFriendlyNames[i] = i.ToString() + " - SVideo";
                }
                else if (Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs[i].Type == AForge.Video.DirectShow.PhysicalConnectorType.VideoComposite) {
                    videoPinsWithFriendlyNames[i] = i.ToString() + " - Composite";
                }
                else if (Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs[i].Type == AForge.Video.DirectShow.PhysicalConnectorType.VideoTuner) {
                    videoPinsWithFriendlyNames[i] = i.ToString() + " - Tuner";
                }
                else {
                    videoPinsWithFriendlyNames[i] = i.ToString() + " - Unrecognized";
                }
            }

            AForgeVideoPinsWithFriendlyNames = videoPinsWithFriendlyNames;

            //Need to handle pin if the svideo flag is set in ITpipes:
            //ITpipes will take the first capture pin of type VideoCapturePin.Svideo, so that's what I'll do.

            if (Settings.SelectedItVideoCapturePin == SharedEnums.VideoCapturePin.Svideo) {

                for (int i = 0; i < Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs.Length; i++) {

                    if (Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs[i].Type == AForge.Video.DirectShow.PhysicalConnectorType.VideoSVideo) {

                        SelectedVideoInput = Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs[i];
                        SelectedAforgeVideoPinFriendlyName = videoPinsWithFriendlyNames[i];
                        break;
                    }

                    if (i == Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs.Length - 1) {
                        //on devices with no svideo pin, the selected pin needs to be set to Composite so the settings saving won't keep 'svideo' in the video pin.

                        Settings.SelectedItVideoCapturePin = SharedEnums.VideoCapturePin.Composite;
                    }
                }
            }
            else {
                for (int i = 0; i < Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs.Length; i++) {

                    if (Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs[i].Type == AForge.Video.DirectShow.PhysicalConnectorType.VideoComposite) {

                        SelectedVideoInput = Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs[i];
                        SelectedAforgeVideoPinFriendlyName = videoPinsWithFriendlyNames[i];
                        break;
                    }
                }
            }

            if (Settings.CurAforgeCaptureDevice.IsRunning == false) {

                Settings.CurAforgeCaptureDevice.NewFrame += newAforgeFrameHandler;

                Settings.CurAforgeCaptureDevice.Start();
            }

        }

        private void newAforgeFrameHandler(object sender, AForge.Video.NewFrameEventArgs e) {

            try {

                System.Drawing.Image newRawImage = (System.Drawing.Bitmap)e.Frame.Clone();

                newRawImage.Save(_newFrameMS, System.Drawing.Imaging.ImageFormat.Bmp);
                _newFrameMS.Seek(0, System.IO.SeekOrigin.Begin);

                System.Windows.Media.Imaging.BitmapImage newFrameImage = new System.Windows.Media.Imaging.BitmapImage();

                newFrameImage.BeginInit();
                newFrameImage.StreamSource = _newFrameMS;
                newFrameImage.EndInit();
                newFrameImage.Freeze();

                this.AforgeFrameImage = newFrameImage;

            }
            catch { }
        }
        
        private void getSelectedVideoInput() {
            if (this.Settings.CurAforgeCaptureDevice != null && this.Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs.Length != 0) {
                if (this.Settings.SelectedItVideoCapturePin == SharedEnums.VideoCapturePin.Svideo) {
                    for (int i = 0; i < this.Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs.Length; i++) {

                        if (this.Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs[i].Type == AForge.Video.DirectShow.PhysicalConnectorType.VideoSVideo) {

                            this.SelectedVideoInput = this.Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs[i];
                        }
                    }
                }
                else {
                    for (int i = 0; i < this.Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs.Length; i++) {
                        if (this.Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs[i].Type == AForge.Video.DirectShow.PhysicalConnectorType.VideoComposite) {
                            this.SelectedVideoInput = this.Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs[i];
                        }
                    }
                }
            }
        }
        
        private void setAforgeVideoCapabilies() {
            if (this.AForgeVideoPinsWithFriendlyNames == null || 
                this.SettingsSelectedFilterInfo == null || 
                this.Settings == null || 
                this.Settings.CurAforgeCaptureDevice == null ||
                this.Settings.CurAforgeCaptureDevice.IsRunning == false) {

                return;
            }

            Settings.CurAforgeCaptureDevice.Stop();

            Settings.CurAforgeCaptureDevice.NewFrame -= newAforgeFrameHandler;

            for (int i = 0; i < this.AForgeVideoPinsWithFriendlyNames.Length; i++) {
                if (SelectedAforgeVideoPinFriendlyName == AForgeVideoPinsWithFriendlyNames[i]) {
                    
                    if (Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs[i].Type == PhysicalConnectorType.VideoComposite) {
                        this.Settings.SelectedItVideoCapturePin = SharedEnums.VideoCapturePin.Composite;
                    }
                    else if (Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs[i].Type == PhysicalConnectorType.VideoSVideo) {
                        this.Settings.SelectedItVideoCapturePin = SharedEnums.VideoCapturePin.Svideo;
                    }

                    Settings.CurAforgeCaptureDevice.CrossbarVideoInput = Settings.CurAforgeCaptureDevice.AvailableCrossbarVideoInputs[i];
                    break;
                }
            }
            
            this._newFrameMS = new MemoryStream();
            this.AforgeFrameImage = new BitmapImage();

            Settings.CurAforgeCaptureDevice.NewFrame += newAforgeFrameHandler;

            Settings.CurAforgeCaptureDevice.Start();

        }

        private void OnPropertyChanged(object sender, PropertyChangedEventArgs e) {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) {
                handler(sender, new PropertyChangedEventArgs(e.PropertyName));
            }
        }
        
        #endregion Private Functions

        #region Public Functions

        public void Dispose() {
            if (this._audioRecordingInstructionsWindow != null) {
                this._audioRecordingInstructionsWindow.Close();
                this._audioRecordingInstructionsWindow = null;
            }
            
            if (this.Settings != null) {
                if (this.Settings.CurAforgeCaptureDevice != null) {

                    this.Settings.CurAforgeCaptureDevice.NewFrame -= this.newAforgeFrameHandler;
                }

                    this.Settings.Dispose();
            }
        }

        #endregion Public Functions

    }

    public class GenericCommandCanAlwaysExecute : ICommand {

        private Action _action;

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter) {
            return true;
        }

        public void Execute(object parameter) {
            _action();
        }

        public GenericCommandCanAlwaysExecute(Action action) {
            _action = action;
        }
    }

    public class BooleanInverseConverter : IValueConverter {

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
            return !((bool)value);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
            return !((bool)value);
        }
    }
    
    public class WholeNumberPercentConverter : IValueConverter {

        public static char[] NUMERIC_CHARS = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {

            return value.ToString();

        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {

            if (value == null) {
                return 0;
            }

            int validationInt = -1;

            if (int.TryParse((string)value, out validationInt) == true) {

                if (validationInt > 100) {
                    return 100;
                }
                else if (validationInt < 0) {
                    return 0;
                }
            }

            else {

                string valueAsString = value.ToString();
                string valueAsDigits = "";

                int returnInt = -1;

                for (int i = 0; i < valueAsString.Length; i++) {
                    if (NUMERIC_CHARS.Contains(valueAsString[i])) {
                        valueAsDigits += valueAsString[i];
                    }
                }

                if (valueAsDigits == "") {
                    return 0;
                }

                else if (int.TryParse(valueAsDigits, out returnInt)) {

                    if (returnInt > 100) {
                        return 100;
                    }
                    else if (returnInt < 0) {
                        return 0;
                    }

                    return returnInt;
                }
                else {
                    return 0;
                }

            }

            return 0;
        }
    }

    public class StringToNullableIntConverter : IValueConverter {

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
            return value == null ? string.Empty : value.ToString();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
            

            if (string.IsNullOrWhiteSpace((string)value)) {
                return null;
            }

            int returnInt = 0;

            if (int.TryParse((string)value, out returnInt)) {
                return returnInt;
            }

            return null;
        }
    }
}

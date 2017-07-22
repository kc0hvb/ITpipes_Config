using System;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using ITPipes.OverlayDrivers.Legacy;

namespace ITpipes_Config {

    public class OverlaySettings : INotifyPropertyChanged, IDisposable {
        
        #region Private Variables

        private ObservableCollection<OverlayControlClass> _controlsAvailable;
        private bool _isOverlayEnabled;
        private ObservableCollection<string> _overlayNames = new ObservableCollection<string>();
        private int? _overlayPort; 
        private int? _lateralPort;
        private int? _inclinometerPort;
        private float? _mainlineCalibration;
        private float? _lateralCalibration;
        private DriverLoader _v2DriverLoader;

        #endregion Private Variables
        
        #region Public Properties

        public ObservableCollection<OverlayControlClass> ControlsAvailable {
            get {
                return _controlsAvailable;
            }
            set {
                _controlsAvailable = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("ControlsAvailable"));
            }
        }
        public bool IsOverlayEnabled {
            get {
                return _isOverlayEnabled;
            }
            set {
                _isOverlayEnabled = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("IsOverlayEnabled"));
            }
        }
        public ObservableCollection<string> OverlayNames {
            get {
                return _overlayNames;
            }
            set {
                _overlayNames = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("OverlayNames"));
            }
        }
        public int? OverlayPort {
            get {
                return _overlayPort;
            }
            set {
                _overlayPort = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("OverlayPort"));
            }
        }
        public int? LateralPort {
            get {
                return _lateralPort;
            }
            set {
                _lateralPort = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("LateralPort"));
            }
        }
        public int? InclinometerPort {
            get {
                return _inclinometerPort;
            }
            set {
                _inclinometerPort = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("InclinometerPort"));
            }
        }
        public float? MainlineCalibration {
            get {
                return _mainlineCalibration;
            }
            set {
                _mainlineCalibration = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("MainlineCalibration"));
            }
        }
        public float? LateralCalibration {
            get {
                return _lateralCalibration;
            }
            set {
                _lateralCalibration = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("LateralCalibration"));
            }
        }
        public DriverLoader V2DriverLoader {
            get {
                return _v2DriverLoader;
            }
            set {
                _v2DriverLoader = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("V2DriverLoader"));
            }
        }

        #endregion Public Properties
        
        #region Events

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion

        #region Constructors

        public OverlaySettings(DataRow curSettingsRow, string itpipesDriversDirectory = @"C:\Program Files\InspectIT\Drivers") {

            this.ControlsAvailable = new ObservableCollection<OverlayControlClass>();

            string[] availableV2OverlayDrivers = null;

            if (File.Exists(Path.Combine(itpipesDriversDirectory, "overlay.dll"))) {
                this.V2DriverLoader = new DriverLoader();
                this.V2DriverLoader.DriverDirectory = itpipesDriversDirectory;
                availableV2OverlayDrivers = this.V2DriverLoader.ScanDriverDirectory();

                if (availableV2OverlayDrivers.Contains("PCI4E") == false) {
                    this.addOverlayToOverlayList("US Digital", false);
                }
                if (availableV2OverlayDrivers.Contains("QSB") == false) {
                    this.addOverlayToOverlayList("QSB", false);
                }
                if (availableV2OverlayDrivers.Contains("USB4") == false) {
                    this.addOverlayToOverlayList("USB4", false);
                }

                foreach (string curOverlay in availableV2OverlayDrivers) {
                    this.addOverlayToOverlayList(curOverlay);
                }
            }
            else {
                this.addOverlayToOverlayList("US Digital", false);
                this.addOverlayToOverlayList("QSB", false);
                this.addOverlayToOverlayList("USB4", false);
                this.addOverlayToOverlayList("Aries", false);
                this.addOverlayToOverlayList("Aries5000", false);
                this.addOverlayToOverlayList("IBAK", false);
                this.addOverlayToOverlayList("IPEK_DE03SW", false);
            }


            if ((bool)curSettingsRow["Counter_Enabled"] == true ||
                (curSettingsRow.Table.Columns.Contains("CounterV2_Enabled") && (bool)curSettingsRow["CounterV2_Enabled"] == true)) {
                this.IsOverlayEnabled = true;
            }

            //Ports:
            if (curSettingsRow.Table.Columns.Contains("CounterV2_Enabled") &&
                (bool)curSettingsRow["CounterV2_Enabled"] == true &&
                curSettingsRow.Table.Columns.Contains("CounterV2_Port") &&
                curSettingsRow["CounterV2_Port"] != DBNull.Value) {

                this.OverlayPort = int.Parse((string)curSettingsRow["CounterV2_Port"]); //no idea why this is a string in the database.
            }
            else {
                this.OverlayPort = (curSettingsRow["Counter_Port"] == DBNull.Value) ? (int?)null : int.Parse((string)curSettingsRow["Counter_Port"]); //no idea why this is a string in setup.mdb either. *shrug*
            }

            this.LateralPort = (curSettingsRow["Counter_Lateral_Port"] == DBNull.Value) ? (int?)null : (int?)curSettingsRow["Counter_Lateral_Port"];

            if (curSettingsRow.Table.Columns.Contains("RST_Inclinometer_Port")) {
                this.InclinometerPort = (curSettingsRow["RST_Inclinometer_Port"] == DBNull.Value) ? (int?)null : (int)curSettingsRow["RST_Inclinometer_Port"];
            }

            //Calibrations:
            this.MainlineCalibration = (curSettingsRow["Counter_Calibration"] == DBNull.Value) ? (float?)null : (float)curSettingsRow["Counter_Calibration"];
            this.LateralCalibration = (curSettingsRow["Counter_Lateral_Calibration"] == DBNull.Value) ? (float?)null : (float)curSettingsRow["Counter_Lateral_Calibration"];
        }

        #endregion Constructors
        
        #region Private Functions

        private void addOverlayToOverlayList(string name, bool useV2DriverLoader = true) {
            OverlayControlClass newOverlay = new OverlayControlClass(name, useV2DriverLoader);
            this.ControlsAvailable.Add(newOverlay);
            this.OverlayNames.Add(newOverlay.FriendlyOverlayName);
        }

        private void OnPropertyChanged(object sender, PropertyChangedEventArgs e) {
            if (PropertyChanged != null) {
                PropertyChanged(sender, e);
            }
        }

        #endregion
        
        #region Public Functions

        public static OverlaySettings getFullyPopulatedOverlaySettingsObject(string curOleConnString, string itpipesDriversDirectory = @"C:\Program Files\InspectIT\Drivers") {

            if (string.IsNullOrEmpty(curOleConnString)) {
                throw new Exception("Connection string used to get populated overlay settings object cannot be null or empty");
            }

            OverlaySettings returnOverlaySettings = null;

            using (OleDbConnection curOleConn = new OleDbConnection(curOleConnString))
            using (OleDbCommand curCommand = curOleConn.CreateCommand())
            using (DataTable curSettingsTable = new DataTable()) {

                curOleConn.Open();
                curCommand.CommandText = "SELECT TOP 1 * FROM S"; //selecting * instead of specific fields to prevent crashes when loading from 18xx versions, which lack v2 settings.
                curSettingsTable.Load(curCommand.ExecuteReader());

                if (curSettingsTable.Rows.Count == 0) {
                    throw new Exception("ITpipes settings database (Setup.mdb) does not contain a row from which to populate overlay settings.");
                }

                returnOverlaySettings = new OverlaySettings(curSettingsTable.Rows[0], itpipesDriversDirectory);
            }

            return returnOverlaySettings;
        }

        public void ScanSerialPortsForV2Overlay(OverlayControlClass v2OverlayControl) {

            if (v2OverlayControl == null ||
                v2OverlayControl.UsesV2DriverLoader == false ||
                File.Exists(@"C:\Program Files\InspectIT\Drivers\overlay.dll") == false) {

                return;
            }

            string[] validPorts = System.IO.Ports.SerialPort.GetPortNames();

            if (validPorts.Length == 0) {
                return;
            }

            if (this.V2DriverLoader == null) {
                this.V2DriverLoader = new DriverLoader();
                this.V2DriverLoader.DriverDirectory = @"C:\Program Files\InspectIT\Drivers";
                this.V2DriverLoader.ScanDriverDirectory();
            }

            try {
                this.V2DriverLoader.LoadDriver(v2OverlayControl.OverlayName);
            }
            catch (Exception ex) {
                MessageBox.Show("Error encountered loading driver: '" + v2OverlayControl.OverlayName + "'\n\nSpecific Error: " + ex.Message);
            }

            if (this.V2DriverLoader.DriverIsLoaded == true) {

                for (int i = 0; i < validPorts.Length; i++) {
                    string curPort = validPorts[i];

                    //Handle if it's an LPT (or otherwise non-COM) port:
                    if (curPort.Contains("COM") == false) {
                        continue;
                    }

                    this.V2DriverLoader.PortId = int.Parse(curPort.Replace("COM", ""));

                    try {
                        if (this.V2DriverLoader.PortIsOpen) { this.V2DriverLoader.ClosePort(); }
                        this.V2DriverLoader.OpenPort();
                        float tempTestFloat = this.V2DriverLoader.GetCounter(1);
                        V2DriverLoader.ClosePort();

                        if (tempTestFloat.Equals(float.NaN) == false) {
                            this.OverlayPort = V2DriverLoader.PortId;
                            this.LateralPort = V2DriverLoader.PortId + 1;

                            return;
                        }
                    }
                    catch (Exception ex) {
                        //This violates single-responsibility, but I'm allowing this to show a messagebox because the port may be locked by another process, or it may have a driver-level issue
                        //I've seen ports which claim they don't exist when I scan them, even though Windows says they exist. I'd rather these ports don't prevent scanning the other ports.
                        MessageBox.Show("Error encountered scanning port " + V2DriverLoader.PortId.ToString() + "\n\nSpecific Error: " + ex.Message);
                    }
                    finally {
                        this.V2DriverLoader.ClosePort();
                        this.V2DriverLoader.UnloadDriver();
                        this.V2DriverLoader.PortId = -1;
                    }
                }
            }
        }

        public void Dispose() {
            if (this.V2DriverLoader != null) {
                if (this.V2DriverLoader.PortIsOpen) {
                    this.V2DriverLoader.ClosePort();
                }

                if (this.V2DriverLoader.DriverIsLoaded) {
                    this.V2DriverLoader.UnloadDriver();
                }

                this.V2DriverLoader = null;
            }

            if (this.ControlsAvailable != null) {
                this.ControlsAvailable = null;
            }
        }

        #endregion Public Functions

    }

    public class OverlayControlClass : INotifyPropertyChanged {

        #region Read-Only Resources

        public static readonly Dictionary<string, string> UglyOverlayNameToFriendlyNameDict = new Dictionary<string, string>()
        {
            //V1:
            {"USDigital", "US Digital PCI4E" },
            {"QSB", "US Digital QSB" },
            {"Aries", "Aries VL3000" },
            {"Aries5000", "Aries VL5000" },
            {"IPEK_DE03SW", "IPEK CVO" },
            {"IBAK", "IBAK EDE7" },
            {"BOB4", "XBOB-4E" }, //We're pretending that the v1 Bob4 control doesn't exist. It doesn't exist in 18xx versions, and in versions in which it exists, it's pretty janky.

            //V1 and V2 Controls have the same control with the same name:
            {"USB4", "US Digital USB4" },

            //V2-only:
            {"WKI", "Rausch WKI" },
            {"BOB-4", "XBOB-4E" },
            {"IBAK EDE7", "IBAK (EDE7)" },
            {"IPEK CVO", "IPEK (CVO, 3SW, 7C, Pendant)" },
            {"VL3000", "Aries VL3000" },
            {"VL5000", "Aries VL5000" }
        };

        public static readonly Dictionary<string, string> V2ToV1Dict = new Dictionary<string, string>()
        {
            {"BOB-4", "BOB4" }, //We're pretending that the v1 Bob4 control doesn't exist. It doesn't exist in 18xx versions, and in versions in which it exists, it's pretty janky.
            {"IBAK EDE7", "IBAK" },
            {"IPEK CVO", "IPEK_DE03SW" },
            {"VL3000", "Aries" },
            {"VL5000", "Aries5000" },
            {"USB4", "USB4" }
        };

        #endregion Read-Only Resources

        #region Private Variables

        private string _overlayName;
        private string _alternateOverlayControlName;
        private string _friendlyOverlayName;
        private bool _hasV1Control;
        private bool _forceUseV1Control;
        private bool _usesV2DriverLoader;
        private bool _driverSupportsPortScanning = false;

        #endregion Private Variables

        #region Public Properties

        public string OverlayName {
            get {
                return _overlayName;
            }
            set {
                _overlayName = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("OverlayName"));
            }
        }
        public string AlternateOverlayControlName {
            get {
                return _alternateOverlayControlName;
            }
            set {
                _alternateOverlayControlName = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("AlternateOverlayControlName"));
            }
        }
        public string FriendlyOverlayName {
            get {
                return _friendlyOverlayName;
            }
            set {
                _friendlyOverlayName = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("FriendlyOverlayName"));
            }
        }
        public bool HasV1Control {
            get {
                return _hasV1Control;
            }
            set {
                _hasV1Control = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("HasV1Control"));
            }
        }
        public bool ForceUseV1Control {
            get {
                return _forceUseV1Control;
            }
            set {
                _forceUseV1Control = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("ForceUseV1Control"));
            }
        }
        public bool UsesV2DriverLoader {
            get {
                return _usesV2DriverLoader;
            }
            set {
                _usesV2DriverLoader = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("UsesV2DriverLoader"));
            }
        }
        public bool DriverSupportsPortScanning {
            get {
                return _driverSupportsPortScanning;
            }
            set {
                _driverSupportsPortScanning = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("DriverSupportsPortScanning"));
            }
        }

        #endregion Public Properties

        #region Events

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion Events

        #region Constructors

        public OverlayControlClass(string name, bool usesV2DriverLoader = true) {

            this.OverlayName = name;
            this.UsesV2DriverLoader = usesV2DriverLoader;

            if (this.UsesV2DriverLoader && 
                this.OverlayName != "USB4" &&
                this.OverlayName != "PCI4E" &&
                this.OverlayName != "QSB") {

                this.DriverSupportsPortScanning = true;
            }
            else {
                this.DriverSupportsPortScanning = false;
            }

            string altOverlayControlname = "";
            string tempFriendlyName = "";

            V2ToV1Dict.TryGetValue(name, out altOverlayControlname);
            if (string.IsNullOrWhiteSpace(altOverlayControlname) == false &&
                usesV2DriverLoader == true) {
                this.AlternateOverlayControlName = altOverlayControlname;
                this.HasV1Control = true;
            }

            UglyOverlayNameToFriendlyNameDict.TryGetValue(name, out tempFriendlyName);
            if (tempFriendlyName != null) {
                this.FriendlyOverlayName = tempFriendlyName;
            }
            else {
                this.FriendlyOverlayName = name;
            }
        }

        #endregion Constructors

        #region Private Functions

        private void OnPropertyChanged(object sender, PropertyChangedEventArgs e) {
            if (PropertyChanged != null) {
                PropertyChanged(sender, e);
            }
        }
        
        #endregion Private Functions

        #region Public Functions



        #endregion Public Functions
        
    }
}
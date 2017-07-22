#define DEBUG

using System;
using System.ComponentModel;
using System.Windows.Media.Imaging;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using ITpipes_Config.SharedEnums;

namespace ITpipes_Config.AddressBook {
    

    class AddressBookModel : INotifyPropertyChanged, IDisposable {
        
        #region Private Variables
        
        private BackgroundWorker _templateBW = null;
        private ObservableCollection<AddressBookContact> _addressBookContactList = new ObservableCollection<AddressBookContact>();
        private ObservableCollection<TemplateContact> _templateContactList = new ObservableCollection<TemplateContact>();

        private List<string> _corruptTemplates = new List<string>();
        private object _corruptTplLock = new object();

        #endregion Private Variables

        public static string Path_To_Default_ITPipes_Install_Directory = @"C:\Program Files\InspectIT";

        #region Public Properties

        public ObservableCollection<AddressBookContact> AddressBookContactList {
            get {
                return _addressBookContactList;
            }
            set {
                _addressBookContactList = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("AddressBookContactList"));
            }
        }
        public ObservableCollection<TemplateContact> TemplateContacts {
            get {
                return _templateContactList;
            }
            set {
                _templateContactList = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("TemplateContacts"));
            }
        }

        #endregion Public Properties

        #region Events

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion Events

        #region Constructors

        public AddressBookModel() {
            this.populateAddressBookContactList();
        }

        #endregion Constructors

        #region Private Functions

        private void populateAddressBookContactList(string pathToAddressDatabase = @"C:\Program Files\InspectIT\address.mdb") {

            if (File.Exists(pathToAddressDatabase) == false) {

                string pathToBlankAddressDb = Path.Combine(Path.GetDirectoryName(pathToAddressDatabase), @"blankdb\address.mdb");

                if (File.Exists(pathToBlankAddressDb)) {
                    File.Copy(pathToBlankAddressDb, pathToAddressDatabase);
                }
                else {

                    throw new FileNotFoundException(string.Format("The path to the address database does not exist: {0}", pathToAddressDatabase));
                }
            }

            string abConnString = @"Provider = Microsoft.Jet.OLEDB.4.0; Data Source = C:\Program Files\InspectIT\address.mdb";

            using (OleDbConnection curOleConn = new OleDbConnection(abConnString))
            using (OleDbCommand curOleCommand = curOleConn.CreateCommand())
            using (DataTable addressBookDT = new DataTable()) {

                try {
                    curOleConn.Open();
                }
                catch {
                    throw new Exception(string.Format("Could not open OleDb connection using Jet 4.0 to database located at: {0}", pathToAddressDatabase));
                }

                curOleCommand.CommandText = "SELECT * FROM `Address`";


                addressBookDT.Load(curOleCommand.ExecuteReader());


                foreach (DataRow curRow in addressBookDT.Rows) {
                    this.AddressBookContactList.Add(new AddressBookContact(curRow));
                }
            }


            this._templateBW = new BackgroundWorker();

            _templateBW.DoWork += TemplateBW_DoWork;
            _templateBW.RunWorkerCompleted += TemplateBW_RunWorkerCompleted;
            _templateBW.RunWorkerAsync();

        }

        private void TemplateBW_DoWork(object sender, DoWorkEventArgs e) {

            ObservableCollection<TemplateContact> returnTemplateContacts = new ObservableCollection<TemplateContact>();

            string[] availableTemplates = Directory.GetFiles(@"C:\Program Files\InspectIT\Templates", "*.tpl", SearchOption.TopDirectoryOnly);

            foreach (string curAvaialableTemplate in availableTemplates) {
                TemplateContact newTC = null;
                try {
                    newTC = new TemplateContact(curAvaialableTemplate);
                    returnTemplateContacts.Add(newTC);
                }
                catch (Exception ex) {

                    //lock (_corruptTplLock) {

                        this._corruptTemplates.Add(curAvaialableTemplate);
                    //}

#if DEBUG

                    if (System.Diagnostics.Debugger.IsAttached) {
                        System.Diagnostics.Debugger.Break();
                    }
#endif
                }
            }

            e.Result = returnTemplateContacts;
        }

        private void TemplateBW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {

            if (e.Cancelled == false && e.Error == null) {
                this.TemplateContacts = (ObservableCollection<TemplateContact>)e.Result;
            }

            else {
                throw new Exception(string.Format("Unable to populate template address book records: {0}", e.Error.Message));
            }
        }

        private string getAllPkValuesFromContactsForUseWithSqlWhereInClause(ObservableCollection<AddressBookContact> input) {

            string returnString = "(";

            foreach (AddressBookContact curContact in input) {
                if (curContact.SourceRecordIntPK > 0 && string.IsNullOrWhiteSpace(curContact.Name) == false) {
                    returnString += curContact.SourceRecordIntPK.ToString() + ", ";
                }
            }

            returnString += ")";

            return returnString;

        }

        private void OnPropertyChanged(object sender, PropertyChangedEventArgs e) {
            if (PropertyChanged != null) {
                PropertyChanged(sender, e);
            }
        }

        #endregion Private Functions

        #region Public Functions

        public void SaveChangesToFile(string addressBookPath = @"C:\Program Files\InspectIT\Address.mdb") {

            if (File.Exists(addressBookPath) == false) {
                throw new FileNotFoundException();
            }

            using (OleDbConnection curOleConn = new OleDbConnection(string.Format("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = {0}", addressBookPath)))
            using (OleDbCommand curOleCommand = curOleConn.CreateCommand()) {

                curOleConn.Open();

                //Need to clear out all contacts which have been removed from the contact list in the template:

                string pkValuesInContactList = getAllPkValuesFromContactsForUseWithSqlWhereInClause(this.AddressBookContactList);

                if (pkValuesInContactList != "()") {

                    curOleCommand.CommandText = "DELETE FROM `Address` WHERE [Address_ID] NOT IN " + pkValuesInContactList;
                    curOleCommand.ExecuteNonQuery();
                }
                else {
                    curOleCommand.CommandText = "DELETE FROM `Address`";
                    curOleCommand.ExecuteNonQuery();
                }

                foreach (AddressBookContact curContact in this.AddressBookContactList) {

                    if (curContact == null || string.IsNullOrWhiteSpace(curContact.Name) || curContact.ContactHasBeenModified == false) {
                        continue;
                    }

                    AddressBookContact.CopyLogoToITpipesDirectory(curContact, "AddressBook");

                    if (curContact.SourceRecordIntPK > 0) {
                        AddressBookContact.LoadCommandToUpdateExistingContactRecord(curOleCommand, curContact);
                    }

                    else {
                        AddressBookContact.LoadCommandToInsertNewContactRecord(curOleCommand, curContact);
                    }

                    curOleCommand.ExecuteNonQuery();

                    if (curContact.SourceRecordIntPK > 0 == false) {
                        curOleCommand.Parameters.Clear();
                        curOleCommand.CommandText = "SELECT TOP 1 [Address_ID] FROM Address ORDER BY [Address_ID] DESC";
                        curContact.SourceRecordIntPK = (int)curOleCommand.ExecuteScalar();
                    }
                }
            }

        }
        
        public void Dispose() {

            if (this.AddressBookContactList != null && AddressBookContactList.Count > 0) {

                //using a for loop because you can't remove items inside an enumeration of the collection
                for (int i = AddressBookContactList.Count - 1; i >= 0; i--) {
                    AddressBookContact curContact = AddressBookContactList[i];
                    curContact.Dispose();
                    AddressBookContactList.RemoveAt(i);
                }
            }

            if (this.TemplateContacts != null && TemplateContacts.Count > 0) {

                for (int i = TemplateContacts.Count - 1; i >= 0; i--) {

                    TemplateContact curContact = TemplateContacts[i];
                    curContact.Dispose();
                    TemplateContacts.RemoveAt(i);
                }
            }

            if (_templateBW != null) {

                _templateBW.Dispose();
            }

            if (this.TemplateContacts != null) {
                foreach (TemplateContact curContact in this.TemplateContacts) {

                    curContact.Dispose();
                }

                this.TemplateContacts = null;
            }
        }

        #endregion Public Functions

    }

    class TemplateContact : INotifyPropertyChanged, IDisposable {


        #region Private Variables

        private string _pathToTemporarilyExtractedProjectMdb;
        private string _privateOleConnString;
        private string _pathToTemplateFile;
        private string _templateName;
        private int _sourceRecordNumericPK;
        private Ionic.Zip.ZipFile _templateZipObject;
        private ObservableCollection<AddressBookContact> _templateContacts;

        #endregion Private Variables

        #region Public Properties
        
        public string PathToTemplateFile {
            get {
                return _pathToTemplateFile;
            }
            set {
                _pathToTemplateFile = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("PathToTemplateFile"));
            }
        }
        public string TemplateName {
            get {
                return _templateName;
            }
            set {
                _templateName = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("TemplateName"));
            }
        }
        public int SourceRecordNumericPK {
            get {
                return _sourceRecordNumericPK;
            }
            set {
                _sourceRecordNumericPK = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("SourceRecordNumericPK"));
            }
        }
        public string TemplateFileDirectory {
            get {
                return _pathToTemporarilyExtractedProjectMdb;
            }
            set {
                _pathToTemporarilyExtractedProjectMdb = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("PathToTempExtractedProjectMdb"));
            }
        }
        public Ionic.Zip.ZipFile TemplateZipObject {
            get {
                return _templateZipObject;
            }
            set {
                _templateZipObject = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("TemplateZipObject"));
            }
        }
        public ObservableCollection<AddressBookContact> TemplateContacts {
            get {
                if (_templateContacts == null) {
                    _templateContacts = new ObservableCollection<AddressBookContact>();
                }
                return _templateContacts;
            }
            set {
                _templateContacts = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("TemplateContacts"));
            }
        }

        #endregion Public Properties

        #region Events

        public event PropertyChangedEventHandler PropertyChanged;

        public event CorruptTemplateDelegate CorruptTemplateLocated;

        public delegate void CorruptTemplateDelegate(object sender, CorruptTemplateEventArgs e);


        #endregion Events

        #region Constructors

        public TemplateContact(string pathToTplFile) {

            if (string.IsNullOrEmpty(pathToTplFile)) {
                throw new Exception(string.Format("Path to template was {0}", pathToTplFile == null ? "[Null]" : "[Empty]"));
            }

            if (File.Exists(pathToTplFile) == false) {
                throw new FileNotFoundException(string.Format("File does not exist or is inaccessible: {0}", pathToTplFile));
            }

            this.PathToTemplateFile = pathToTplFile;
            this.TemplateName = Path.GetFileNameWithoutExtension(pathToTplFile);

            this.TemplateFileDirectory = Path.GetDirectoryName(pathToTplFile) + @"\" + this.TemplateName;
            string templateFileNameOnly = Path.GetFileName(pathToTplFile);



            this.TemplateZipObject = new Ionic.Zip.ZipFile(pathToTplFile);
            this.TemplateZipObject.Password = "ITWCA321";
            this.TemplateZipObject.Encryption = Ionic.Zip.EncryptionAlgorithm.PkzipWeak;



            string tempDatabaseFile = Path.Combine(this.TemplateFileDirectory, "project.mdb");

            if (File.Exists(tempDatabaseFile)) {
                File.Delete(tempDatabaseFile);
            }

            if (File.Exists(tempDatabaseFile + ".tmp")) {
                File.Delete(tempDatabaseFile + ".tmp");
            }

            if (Directory.Exists(TemplateFileDirectory) == false) {
                Directory.CreateDirectory(TemplateFileDirectory);
            }

            foreach (Ionic.Zip.ZipEntry curZipEntry in TemplateZipObject.Entries) {

                if (curZipEntry.FileName == "project.mdb") {

                    curZipEntry.Password = "ITWCA321";

                    //This is a custom overload of this method--added to use Password. DotNetZip default passed null as password for some reason
                    try {

                        curZipEntry.Extract(this.TemplateFileDirectory, Ionic.Zip.ExtractExistingFileAction.OverwriteSilently, "ITWCA321");
                    }
                    catch {
                        if (File.Exists(Path.Combine(this.TemplateFileDirectory, "project.mdb"))) {
                            try {
                                File.Delete(Path.Combine(this.TemplateFileDirectory, "project.mdb"));
                            }
                            catch {

                            }
                        }
                    }
                
                    break;
                }
            }

            this._pathToTemporarilyExtractedProjectMdb = Path.Combine(TemplateFileDirectory, this.TemplateName + "_cfgtmp");

            if (File.Exists(this._pathToTemporarilyExtractedProjectMdb)) {
                File.Delete(this._pathToTemporarilyExtractedProjectMdb);
            }

            if (File.Exists(tempDatabaseFile) == false) {
                throw new Exception(string.Format("Could not extract contents of template file: {0}\nZip Entry was corrupt.", this.TemplateName));
            }

            File.Move(tempDatabaseFile, this._pathToTemporarilyExtractedProjectMdb);

            this._privateOleConnString = string.Format("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = {0}", this._pathToTemporarilyExtractedProjectMdb);

            populateTemplateContacts(this._privateOleConnString);
        }

        #endregion Constructors

        #region Private Functions

        private void populateTemplateContacts(string oleDbConnString) {

            if (string.IsNullOrEmpty(oleDbConnString)) {
                throw new FileNotFoundException(string.Format("Connection string passed to populateTemplateContacts was {0}", oleDbConnString == null ? "[Null]" : "[Empty]"));
            }

            using (OleDbConnection curOleConn = new OleDbConnection(oleDbConnString))
            using (OleDbCommand curOleCommand = curOleConn.CreateCommand())
            using (DataTable templateContactDT = new DataTable()) {

                try {
                    curOleConn.Open();
                }
                catch {
#if DEBUG
                    System.Windows.MessageBox.Show(string.Format("Unable to open Jet 4.0 connection to database: {0}", oleDbConnString));
#endif
                    throw;
                }

                curOleCommand.CommandText = "SELECT * FROM Info";

                templateContactDT.Load(curOleCommand.ExecuteReader());

                foreach (DataRow curRow in templateContactDT.Rows) {
                    this.TemplateContacts.Add(new AddressBookContact(curRow));
                }
            }
        }

        private string getAllPkValuesFromContactsForUseWithSqlWhereInClause(ObservableCollection<AddressBookContact> input) {

            string returnString = "(";

            foreach (AddressBookContact curContact in input) {
                if (curContact.SourceRecordIntPK > 0 && string.IsNullOrWhiteSpace(curContact.Name) == false) {
                    returnString += curContact.SourceRecordIntPK.ToString() + ", ";
                }
            }

            returnString += ")";

            return returnString;

        }
        
        private void OnPropertyChanged(object sender, PropertyChangedEventArgs e) {
            if (PropertyChanged != null) {
                PropertyChanged(sender, e);
            }
        }

        public void OnCorruptTemplateLocated(object sender, CorruptTemplateEventArgs e) {

            if (this.CorruptTemplateLocated != null) {
                CorruptTemplateLocated(sender, e);
            }
        }

        #endregion Private Functions

        #region Public Functions

        public void SaveTemplate(bool copyContactLogoToITpipesDirectory = true) {

            if (string.IsNullOrEmpty(this.PathToTemplateFile) ||
                string.IsNullOrEmpty(this._pathToTemporarilyExtractedProjectMdb)) {

                throw new Exception("Cannot save a template whose PathToTemplateFile or _pathToTemporarilyExtractedProjectMdb is null or empty");
            }

            if (File.Exists(this.PathToTemplateFile) == false ||
                File.Exists(this._pathToTemporarilyExtractedProjectMdb) == false) {
                throw new FileNotFoundException(string.Format("Template file or temporary file not found:\nTemplate: {0}\nTemp File: {1}", this.PathToTemplateFile, this._pathToTemporarilyExtractedProjectMdb));
            }

            using (OleDbConnection curOleConn = new OleDbConnection(this._privateOleConnString))
            using (OleDbCommand curOleCommand = curOleConn.CreateCommand()) {

                curOleConn.Open();

                //Need to clear out all contacts which have been removed from the contact list in the template:

                string pkValuesInContactList = getAllPkValuesFromContactsForUseWithSqlWhereInClause(this.TemplateContacts);

                if (pkValuesInContactList != "()") {

                    curOleCommand.CommandText = "DELETE FROM `Info` WHERE [Info_ID] NOT IN " + pkValuesInContactList;
                    curOleCommand.ExecuteNonQuery();
                }
                else {
                    curOleCommand.CommandText = "DELETE FROM `Info`";
                    curOleCommand.ExecuteNonQuery();
                }

                foreach (AddressBookContact curContact in this.TemplateContacts) {
                    if (curContact == null || string.IsNullOrWhiteSpace(curContact.Name) || curContact.ContactHasBeenModified == false) {
                        continue;
                    }

                    AddressBookContact.CopyLogoToITpipesDirectory(curContact, this.TemplateName);

                    if (curContact.SourceRecordIntPK > 0) {
                        AddressBookContact.LoadCommandToUpdateExistingContactRecord(curOleCommand, curContact);
                    }

                    else {
                        AddressBookContact.LoadCommandToInsertNewContactRecord(curOleCommand, curContact);
                    }

                    curOleCommand.ExecuteNonQuery();

                    if (curContact.SourceRecordIntPK < 1) {
                        curOleCommand.Parameters.Clear();
                        curOleCommand.CommandText = "SELECT TOP 1 [Info_ID] FROM Info ORDER BY [Info_ID] DESC";
                        curContact.SourceRecordIntPK = (int)curOleCommand.ExecuteScalar();
                    }
                }
            }
            
            string pathToTempMdb = Path.Combine(Path.GetTempPath(), "project.mdb");

            if (File.Exists(pathToTempMdb)) {
                File.Delete(pathToTempMdb);
            }

            File.Copy(this._pathToTemporarilyExtractedProjectMdb, pathToTempMdb);

            Ionic.Zip.ZipEntry projectMdbEntry = null;

            foreach (var curEntry in TemplateZipObject.Entries) {
                if (curEntry.FileName == "project.mdb") {
                    projectMdbEntry = curEntry;
                }
            }

            if (projectMdbEntry != null) {
                TemplateZipObject.RemoveEntry(projectMdbEntry);
            }

            projectMdbEntry = null;


            TemplateZipObject.Encryption = Ionic.Zip.EncryptionAlgorithm.PkzipWeak;
            TemplateZipObject.AddFile(pathToTempMdb, "\\").Password = "ITWCA321";

            string tempFolderPath = Path.GetTempPath() + Path.GetFileNameWithoutExtension(TemplateZipObject.Name);
            if (Directory.Exists(tempFolderPath) == false) {
                Directory.CreateDirectory(tempFolderPath);
            }

            TemplateZipObject.Save();
            TemplateZipObject.Reset(false); //If you don't reset the zip object, project.mdb becomes corrupt after the second consecutive save--not sure why.

        }
        
        public void Dispose() {
            //Only the ZipFile object is disposable right now
            this.TemplateZipObject.Dispose();
            try {

                if (Directory.Exists(Path.GetDirectoryName(this._pathToTemporarilyExtractedProjectMdb))) {
                    Directory.Delete(Path.GetDirectoryName(this._pathToTemporarilyExtractedProjectMdb), true);
                }
            }
            catch {
                //no need to do a thing
            }
        }
        
        #endregion Public Functions

    }

    class AddressBookContact : INotifyPropertyChanged, IDisposable {

        #region Private Variables

        private string _name = string.Empty;
        private string _street = string.Empty;
        private string _city = string.Empty;
        private string _contactState = string.Empty;
        private string _zipcode = string.Empty;
        private string _phoneNumber = string.Empty;
        private string _pathToLogo = string.Empty;
        private ContactType _contactType;
        private string _department = string.Empty;
        private string _contactEmail = string.Empty;
        private string _contactFax = string.Empty;
        private string _contactMobileNumber = string.Empty;
        private string _responsible = string.Empty;
        private int _sourceRecordIntPK;
        private ContactSourceType _sourceFileType;
        private BitmapImage _logoBitmap;
        private bool _contactHasBeenModified; //used to determine if the template/address book needs to be saved. IO is expensive.

        #endregion Private Variables

        #region Public Properties

        public string Name {
            get {
                return _name;
            }
            set {
                _name = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("Name"));

                this.ContactHasBeenModified = true;
            }
        }
        public string Street {
            get {
                return _street;
            }
            set {
                _street = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("Street"));

                this.ContactHasBeenModified = true;
            }
        }
        public string City {
            get {
                return _city;
            }
            set {
                _city = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("City"));

                this.ContactHasBeenModified = true;
            }
        }
        public string ContactState {
            get {
                return _contactState;
            }
            set {
                _contactState = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("ContactState"));

                this.ContactHasBeenModified = true;
            }
        }
        public string Zipcode {
            get {
                return _zipcode;
            }
            set {
                _zipcode = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("Zipcode"));

                this.ContactHasBeenModified = true;
            }
        }
        public string PhoneNumber {
            get {
                return _phoneNumber;
            }
            set {
                _phoneNumber = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("PhoneNumber"));

                this.ContactHasBeenModified = true;
            }
        }
        public string PathToLogo {
            get {
                return _pathToLogo;
            }
            set {

                if (File.Exists(value)) {
                    _pathToLogo = value;
                    OnPropertyChanged(this, new PropertyChangedEventArgs("PathToLogo"));

                    this.ContactHasBeenModified = true;
                }
            }
        }
        public ContactType ContactType {
            get {
                return _contactType;
            }
            set {
                _contactType = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("ContactType"));

                this.ContactHasBeenModified = true;
            }
        }
        public string Department {
            get {
                return _department;
            }
            set {
                _department = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("Department"));

                this.ContactHasBeenModified = true;
            }
        }
        public string ContactEmail {
            get {
                return _contactEmail;
            }
            set {
                _contactEmail = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("ContactEmail"));

                this.ContactHasBeenModified = true;
            }
        }
        public string ContactFax {
            get {
                return _contactFax;
            }
            set {
                _contactFax = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("ContactFax"));

                this.ContactHasBeenModified = true;
            }
        }
        public string ContactMobileNumber {
            get {
                return _contactMobileNumber;
            }
            set {
                _contactMobileNumber = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("ContactMobileNumber"));

                this.ContactHasBeenModified = true;
            }
        }
        public string Responsible {
            get {
                return _responsible;
            }
            set {
                _responsible = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("Responsible"));

                this.ContactHasBeenModified = true;
            }
        }
        public int SourceRecordIntPK {
            get {
                return _sourceRecordIntPK;
            }
            set {
                _sourceRecordIntPK = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("SourceRecordIntPK"));

                this.ContactHasBeenModified = true;
            }
        }
        public ContactSourceType SourceFileType {
            get {
                return _sourceFileType;
            }
            set {
                _sourceFileType = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("SourceFileType"));

                this.ContactHasBeenModified = true;
            }
        }
        public BitmapImage LogoBitmap {
            get {
                return _logoBitmap;
            }
            set {
                _logoBitmap = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("LogoBitmap"));
                this.ContactHasBeenModified = true;
            }
        }
        public bool ContactHasBeenModified {
            get {
                return _contactHasBeenModified;
            }
            set {
                _contactHasBeenModified = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("ContactHasBeenModified"));
            }
        }

        #endregion Public Properties

        #region Events

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion Events

        #region Constructors

        public AddressBookContact() {

        }

        public AddressBookContact(SharedEnums.ContactSourceType targetContactSourceType) {

            this.SourceFileType = targetContactSourceType;
            this.City = string.Empty;
            this.ContactEmail = string.Empty;
            this.ContactFax = string.Empty;
            this.ContactMobileNumber = string.Empty;
            this.ContactState = string.Empty;
            this.ContactType = ContactType.Contractor;
            this.Department = string.Empty;
            this.Name = string.Empty;
            this.PathToLogo = string.Empty;
            this.PhoneNumber = string.Empty;
            this.Responsible = string.Empty;
            this.SourceRecordIntPK = -1;
            this.Street = string.Empty;
            this.Zipcode = string.Empty;
        }

        public AddressBookContact(DataRow sourceDR) {

            this.Name = getValueIfAvailable(sourceDR, "Contact_Name");
            this.Street = getValueIfAvailable(sourceDR, "Contact_Street");
            this.City = getValueIfAvailable(sourceDR, "Contact_City");
            this.ContactState = getValueIfAvailable(sourceDR, "Contact_State");
            this.Zipcode = getValueIfAvailable(sourceDR, "Contact_Zip");
            this.PhoneNumber = getValueIfAvailable(sourceDR, "Contact_Phone");
            string logoPathRelativeToItpDir = getValueIfAvailable(sourceDR, "Logo");
            this.PathToLogo =
                logoPathRelativeToItpDir == null ? 
                null : 
                Path.Combine(AddressBookModel.Path_To_Default_ITPipes_Install_Directory, logoPathRelativeToItpDir);

            this.ContactType = getContactTypeEnumFromString(getValueIfAvailable(sourceDR, "Category"));
            this.Department = getValueIfAvailable(sourceDR, "Department");
            this.ContactEmail = getValueIfAvailable(sourceDR, "Contact_Email");
            this.ContactFax = getValueIfAvailable(sourceDR, "Contact_Fax");
            this.ContactMobileNumber = getValueIfAvailable(sourceDR, "Contact_Mobile");
            this.Responsible = getValueIfAvailable(sourceDR, "Responsible");

            if (sourceDR.Table.Columns.Contains("Address_ID")) {
                this.SourceRecordIntPK = (int)sourceDR["Address_ID"];
                this.SourceFileType = ContactSourceType.AddressBook;
            }
            else if (sourceDR.Table.Columns.Contains("Info_ID")) {
                this.SourceRecordIntPK = (int)sourceDR["Info_ID"];
                this.SourceFileType = ContactSourceType.TemplateFile;
            }
            else {
                throw new Exception("AddressBookContact Constructor was passed an invalid datarow: " +
                    "The datarow did not contain the primary key column for a template contact ([Info_ID]) or for the address book ([Address_ID])");
            }

            this.ContactHasBeenModified = false;
        }

        #endregion Constructors

        #region Private Functions

        private static string getContactTypeStringFromEnum(SharedEnums.ContactType input) {
            if (input == ContactType.Contractor) {
                return "Contractor";
            }
            else if (input == ContactType.Client) {
                return "Client";
            }
            else if (input == ContactType.None) {
                return string.Empty;
            }
            else {
                throw new Exception(string.Format("Unrecognized Enum Value: ", input.ToString()));
            }
        }
        
        private void OnPropertyChanged(object sender, PropertyChangedEventArgs e) {
            if (PropertyChanged != null) {
                PropertyChanged(sender, e);
            }
        }

        private ContactType getContactTypeEnumFromString(string input) {

            if (string.IsNullOrEmpty(input)) {
                return ContactType.None;
            }

            if (input.ToUpper() == "CONTRACTOR") {
                return ContactType.Contractor;
            }
            else if (input.ToUpper() == "CLIENT") {
                return ContactType.Client;
            }

            return ContactType.None;
        }

        private string getValueIfAvailable(DataRow sourceRow, string columnName) {

            if (sourceRow == null) {
                throw new NullReferenceException("sourceRow for setValueIfAvailable in AddressBookContact constructor was null.");
            }


            if (sourceRow.Table.Columns.Contains(columnName)) {

                try {

                    if (sourceRow[columnName] == DBNull.Value) {
                        return string.Empty;
                    }

                    return (string)sourceRow[columnName];
                }
                catch (Exception ex) {
#if DEBUG
                    System.Windows.MessageBox.Show(
                        string.Format(
                            "Address Book field does not exist or datatype did not match the type of target property:\nColumn Name: {0}\nColumn Datatype: {1}\nSpecific Exception: {3}",
                            columnName,
                            sourceRow.Table.Columns[sourceRow.Table.Columns.IndexOf(columnName)].DataType.ToString(),
                            ex.Message));
#endif
                }
            }
            return string.Empty;
        }
        
        #endregion Private Functions

        #region Public Functions

        public AddressBookContact Clone() {

            AddressBookContact returnContact = new AddressBookContact();

            returnContact.City = this.City;
            returnContact.ContactEmail = this.ContactEmail;
            returnContact.ContactFax = this.ContactFax;
            returnContact.ContactMobileNumber = this.ContactMobileNumber;
            returnContact.ContactState = this.ContactState;
            returnContact.ContactType = this.ContactType;
            returnContact.Department = this.Department;
            returnContact.Name = this.Name;
            returnContact.PathToLogo = this.PathToLogo;
            returnContact.PhoneNumber = this.PhoneNumber;
            returnContact.Responsible = this.Responsible;
            returnContact.Street = this.Street;
            returnContact.Zipcode = this.Zipcode;

            //We deliberately DO NOT clone the primary key to the new object.

            return returnContact;

        }

        public static void LoadCommandToInsertNewContactRecord(OleDbCommand curOleCommand, AddressBookContact newContact) {

            if (curOleCommand == null || string.IsNullOrEmpty(curOleCommand.CommandText) || newContact == null) {
                throw new NullReferenceException("_loadCommandToInsertNewContactRecord was passed NULL or empty parameters");
            }

            if (newContact.SourceRecordIntPK > 0) {
                throw new Exception("_loadCommandToInsertNewContactRecord was passed an address book record which already has a primary key assigned. " +
                    "This method should only be called to populate OleDbCommand parameters for AddressBookContacts which do not yet exist in the target template.");
            }

            string tableName = null;

            if (newContact.SourceFileType == ContactSourceType.AddressBook) {
                tableName = "Address";
            }
            else if (newContact.SourceFileType == ContactSourceType.TemplateFile) {
                tableName = "Info";
            }
            else {
                throw new Exception(string.Format("ContactSourceType Enum was not a recognized value: {0}", newContact.SourceFileType.ToString()));
            }


            curOleCommand.CommandText = "INSERT INTO `" + tableName + "` (" +
                "[Contact_Name], " +
                "[Contact_Street], " +
                "[Contact_City], " +
                "[Contact_State], " +
                "[Contact_Zip], " +
                "[Contact_Phone], " +
                "[Logo], " +
                "[Category], " +
                "[Department], " +
                "[Contact_Email], " +
                "[Contact_Fax], " +
                "[Contact_Mobile], " +
                "[Responsible]) " +
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);";

            curOleCommand.Parameters.Clear();

            curOleCommand.Parameters.AddWithValue("Contact_Name", newContact.Name);
            curOleCommand.Parameters.AddWithValue("Contact_Street", newContact.Street);
            curOleCommand.Parameters.AddWithValue("Contact_City", newContact.City);
            curOleCommand.Parameters.AddWithValue("Contact_State", newContact.ContactState);
            curOleCommand.Parameters.AddWithValue("Contact_Zip", newContact.Zipcode);
            curOleCommand.Parameters.AddWithValue("Contact_Phone", newContact.PhoneNumber);
            curOleCommand.Parameters.AddWithValue("Logo", newContact.PathToLogo == null ? (object)DBNull.Value : string.Format(@"Logos\{0}", Path.GetFileName(newContact.PathToLogo)));
            curOleCommand.Parameters.AddWithValue("Category", getContactTypeStringFromEnum(newContact.ContactType));
            curOleCommand.Parameters.AddWithValue("Department", newContact.Department);
            curOleCommand.Parameters.AddWithValue("Contact_Email", newContact.ContactEmail);
            curOleCommand.Parameters.AddWithValue("Contact_Fax", newContact.ContactEmail);
            curOleCommand.Parameters.AddWithValue("Contact_Mobile", newContact.ContactMobileNumber);
            curOleCommand.Parameters.AddWithValue("Responsible", newContact.Responsible);

        }

        public static void LoadCommandToUpdateExistingContactRecord(OleDbCommand curOleCommand, AddressBookContact existingContact) {

            if (curOleCommand == null || string.IsNullOrEmpty(curOleCommand.CommandText) || existingContact == null) {
                throw new NullReferenceException("_populateUpdateTemplateCommandWithParameters was passed NULL or empty parameters");
            }

            if (existingContact.SourceRecordIntPK < 1) {
                throw new IndexOutOfRangeException("AddressBookContact passed to _populateUpdateTemplateCommandWithParameters did not have a valid primary key parameter.\n" +
                    "This method should only be called to populate OleDbCommand parameters for AddressBookContacts which already exist in the target template.");
            }

            string tableName = null;
            string pkColumnName = null;

            if (existingContact.SourceFileType == ContactSourceType.AddressBook) {
                tableName = "Address";
                pkColumnName = "Address_ID";
            }
            else if (existingContact.SourceFileType == ContactSourceType.TemplateFile) {
                tableName = "Info";
                pkColumnName = "Info_ID";
            }
            else {
                throw new Exception(string.Format("ContactSourceType Enum was not a recognized value: {0}", existingContact.SourceFileType.ToString()));
            }

            curOleCommand.CommandText = "UPDATE `" + tableName + "` SET " +
                "[Contact_Name] = ?, " +
                "[Contact_Street] = ?, " +
                "[Contact_City] = ?, " +
                "[Contact_State] = ?, " +
                "[Contact_Zip] = ?, " +
                "[Contact_Phone] = ?, " +
                "[Logo] = ?, " +
                "[Category] = ?, " +
                "[Department] = ?, " +
                "[Contact_Email] = ?, " +
                "[Contact_Fax] = ?, " +
                "[Contact_Mobile] = ?, " +
                "[Responsible] = ? " +
                "WHERE [" + pkColumnName + "] = ?;";

            curOleCommand.Parameters.Clear();

            curOleCommand.Parameters.AddWithValue("Contact_Name", existingContact.Name);
            curOleCommand.Parameters.AddWithValue("Contact_Street", existingContact.Street);
            curOleCommand.Parameters.AddWithValue("Contact_City", existingContact.City);
            curOleCommand.Parameters.AddWithValue("Contact_State", existingContact.ContactState);
            curOleCommand.Parameters.AddWithValue("Contact_Zip", existingContact.Zipcode);
            curOleCommand.Parameters.AddWithValue("Contact_Phone", existingContact.PhoneNumber);
            //The logo files are stored in the ITpipes program directory, and have to be referenced as a relative path--learned that when the logos weren't working. :-\
            curOleCommand.Parameters.AddWithValue("Logo", existingContact.PathToLogo == null ? (object)DBNull.Value : string.Format(@"Logos\{0}", Path.GetFileName(existingContact.PathToLogo)));
            curOleCommand.Parameters.AddWithValue("Category", getContactTypeStringFromEnum(existingContact.ContactType));
            curOleCommand.Parameters.AddWithValue("Department", existingContact.Department);
            curOleCommand.Parameters.AddWithValue("Contact_Email", existingContact.ContactEmail);
            curOleCommand.Parameters.AddWithValue("Contact_Fax", existingContact.ContactFax);
            curOleCommand.Parameters.AddWithValue("Contact_Mobile", existingContact.ContactMobileNumber);
            curOleCommand.Parameters.AddWithValue("Responsible", existingContact.Responsible);
            curOleCommand.Parameters.AddWithValue("pkColumnName", existingContact.SourceRecordIntPK);

        }

        public static void CopyLogoToITpipesDirectory(AddressBookContact sourceContact, string logoPrefix) {
            if (File.Exists(sourceContact.PathToLogo) == false) {
                return;
            }

            if (Directory.Exists(@"C:\Program Files\InspectIT\Logos") == false) {
                Directory.CreateDirectory(@"C:\Program Files\InspectIT\Logos");
            }

            System.Text.RegularExpressions.Regex filenameRegex = new System.Text.RegularExpressions.Regex("[^a-z^A-Z^0-9]");

            string newFileName = string.Format(@"C:\Program Files\InspectIT\Logos\{0}.{1}.bmp", logoPrefix, filenameRegex.Replace(sourceContact.Name, string.Empty));

            if (newFileName == sourceContact.PathToLogo) {
                return;
            }

            if (File.Exists(newFileName)) {
                File.Delete(newFileName);
            }

            try {

                if (Path.GetExtension(sourceContact.PathToLogo).ToUpper().Contains("BMP")) {

                    File.Copy(sourceContact.PathToLogo, newFileName, true);
                }

                else {

                    Image convertImage = Image.FromFile(sourceContact.PathToLogo);
                    convertImage.Save(newFileName, System.Drawing.Imaging.ImageFormat.Bmp);
                    convertImage.Dispose();
                }

                sourceContact.PathToLogo = newFileName;

            }
            catch (System.IO.IOException ex) {

                System.Windows.MessageBox.Show(string.Format("Failed to copy logo file to ITpipes program directory. Please verify you have write access to C:\\Program Files\\InspectIT\n{0}", ex.Message));

            }

        }

        public void Dispose() {

            //Currently, there isn't actually anything to dispose.

        }
        
        #endregion Public Functions
    }


    public class CorruptTemplateEventArgs : EventArgs {

        public string TemplateName { get; set; }
        public string PathToTemplate { get; set; }
        public Exception AssociatedException { get; set; }

        public CorruptTemplateEventArgs(string name, string tplFilePath, Exception innerException = null) {

            this.TemplateName = name;
            this.PathToTemplate = tplFilePath;
            this.AssociatedException = innerException;

        }

    }

}

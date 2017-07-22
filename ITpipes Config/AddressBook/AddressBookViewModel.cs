using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.IO;
using System.ComponentModel;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Windows.Media.Imaging;

namespace ITpipes_Config.AddressBook {
    class AddressBookViewModel : INotifyPropertyChanged, IDisposable {
        
        private AddressBookModel _curAddressBookModel;
        private AddressBookContact _selectedAddrBookContact;
        private TemplateContact _selectedTemplate;
        private Array _validContactTypes = Enum.GetValues(typeof(SharedEnums.ContactType));

        public AddressBookModel CurAddressBookModel {
            get {
                if (_curAddressBookModel == null) {
                    _curAddressBookModel = new AddressBookModel();
                }
                return _curAddressBookModel;
            }
            set {
                _curAddressBookModel = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("CurAddressBookModel"));
            }
        }
        public AddressBookContact SelectedAddrBookContact {
            get {
                return _selectedAddrBookContact;
            }
            set {

                _selectedAddrBookContact = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("SelectedAddrBookContact"));
            }
        }
        public TemplateContact SelectedTemplate {
            get {
                return _selectedTemplate;
            }
            set {
                _selectedTemplate = value;
                OnPropertyChanged(this, new PropertyChangedEventArgs("SelectedTemplate"));
            }
        }
        public Array ValidContactTypes {
            get {
                return _validContactTypes;
            }
        }

        
        public ICommand SelectNewLogo {
            get {
                return new GenericCommandCanAlwaysExecute(_selectANewLogo);
            }
        }

        public ICommand CopyTemplateContactToAddressBookCommand {
            get {
                return new GenericCommandCanAlwaysExecute(_copySelectedTemplateContactToAddressBook);
            }
        }

        public ICommand CopyAddressBookContactToSelectedTemplateCommand {
            get {
                return new GenericCommandCanAlwaysExecute(_copySelectedAddressBookContactToSelectedTemplate);
            }
        }

        public ICommand AddNewContactCommand {
            get {
                return new GenericCommandCanAlwaysExecute(_addNewContact);
            }
        }

        public ICommand DeleteSelectedContactCommand {
            get {
                return new GenericCommandCanAlwaysExecute(_deleteSelectedContact);
            }
        }

        public ICommand SaveChangesToAddressBookCommand {
            get {
                return new GenericCommandCanAlwaysExecute(_saveChangesToAddressBook);
            }
        }

        public ICommand SaveChangesToSelectedTemplateCommand {
            get {
                return new GenericCommandCanAlwaysExecute(_saveChangesToSelectedTemplate);
            }
        }

        public ICommand SaveSelectedContactToAllTemplatesAndAddressBook {
            get {
                return new GenericCommandCanAlwaysExecute(_saveSelectedContactToAllTemplatesAndAddressBook);
            }
        }

        private void _saveSelectedContactToAllTemplatesAndAddressBook() {

            if (SelectedAddrBookContact == null || CurAddressBookModel == null || CurAddressBookModel.AddressBookContactList == null || CurAddressBookModel.TemplateContacts == null) {
                return;
            }

            if (CurAddressBookModel.AddressBookContactList.Contains(SelectedAddrBookContact) == false) {

                bool foundContactWithMatchingNameInAddressBook = false;

                //using a for because you can't replace an object in a foreach
                for (int i = 0; i < CurAddressBookModel.AddressBookContactList.Count; i++) {

                    AddressBookContact curContact = CurAddressBookModel.AddressBookContactList[i];
                    
                    if (curContact.Name.ToUpper() == SelectedAddrBookContact.Name.ToUpper()) {

                        foundContactWithMatchingNameInAddressBook = true;
                        _cloneValuesAcrossAddressBookContacts(curContact, SelectedAddrBookContact, SharedEnums.ContactSourceType.AddressBook);

                    }
                }

                if (foundContactWithMatchingNameInAddressBook == false) {
                    AddressBookContact newContact = new AddressBookContact();
                    _cloneValuesAcrossAddressBookContacts(newContact, SelectedAddrBookContact, SharedEnums.ContactSourceType.AddressBook);
                    CurAddressBookModel.AddressBookContactList.Add(newContact);
                }
            }

            foreach(TemplateContact curTemplate in CurAddressBookModel.TemplateContacts) {

                if (curTemplate.TemplateContacts.Contains(SelectedAddrBookContact) == false) {

                    bool foundContactWithMatchingNameInTemplate = false;

                    for (int i = 0; i < curTemplate.TemplateContacts.Count; i++) {

                        AddressBookContact curContact = curTemplate.TemplateContacts[i];

                        if (curContact.Name.ToUpper() == SelectedAddrBookContact.Name.ToUpper()) {
                            
                            foundContactWithMatchingNameInTemplate = true;
                            _cloneValuesAcrossAddressBookContacts(curContact, SelectedAddrBookContact, SharedEnums.ContactSourceType.TemplateFile);
                            break;
                        }
                    }

                    if (foundContactWithMatchingNameInTemplate == false) {
                        AddressBookContact newContact = new AddressBookContact();
                        _cloneValuesAcrossAddressBookContacts(newContact, SelectedAddrBookContact, SharedEnums.ContactSourceType.TemplateFile);
                        curTemplate.TemplateContacts.Add(newContact);
                    }
                }
            }
        }
        
        private void _cloneValuesAcrossAddressBookContacts(AddressBookContact targetContact, AddressBookContact sourceContact, SharedEnums.ContactSourceType newContactSourceType) {

            //Have to do some convoluted new string gibberish to prevent the object reference being set instead of the value.
            //There has to be a better way of doing this. I may end up just making addressbookcontacts a struct. *shrug*

            targetContact.City = new string(sourceContact.City.ToCharArray());
            targetContact.ContactEmail = new string(sourceContact.ContactEmail.ToCharArray());
            targetContact.ContactFax = new string(sourceContact.ContactFax.ToCharArray());
            targetContact.ContactMobileNumber = new string(sourceContact.ContactMobileNumber.ToCharArray());
            targetContact.ContactState = new string(sourceContact.ContactState.ToCharArray());
            targetContact.ContactType = sourceContact.ContactType;
            targetContact.Department = new string(sourceContact.Department.ToCharArray());
            targetContact.Name = new string(sourceContact.Name.ToCharArray());
            targetContact.PathToLogo = new string(sourceContact.PathToLogo.ToCharArray());
            targetContact.PhoneNumber = new string(sourceContact.PhoneNumber.ToCharArray());
            targetContact.Responsible = new string(sourceContact.Responsible.ToCharArray());
            targetContact.Street = new string(sourceContact.Street.ToCharArray());
            targetContact.Zipcode = new string(sourceContact.Zipcode.ToCharArray());
            targetContact.SourceFileType = newContactSourceType;

        }

        private void _saveChangesToSelectedTemplate() {
            if (this.CurAddressBookModel.TemplateContacts != null && this.SelectedTemplate.TemplateContacts.Contains(this.SelectedAddrBookContact)) {
                this.SelectedTemplate.SaveTemplate();
            }
        }

        private void _saveChangesToAddressBook() {
            this.CurAddressBookModel.SaveChangesToFile();
        }

        private void _addNewContact() {
            if (this.CurAddressBookModel != null && this.CurAddressBookModel.AddressBookContactList != null) {
                AddressBookContact newContact = new AddressBookContact(SharedEnums.ContactSourceType.AddressBook);
                this.CurAddressBookModel.AddressBookContactList.Add(newContact);
                this.SelectedAddrBookContact = newContact;
            }
        }

        private void _deleteSelectedContact() {
            if (this.SelectedTemplate != null && this.SelectedTemplate.TemplateContacts != null && this.SelectedTemplate.TemplateContacts.Contains(this.SelectedAddrBookContact)) {
                this.SelectedTemplate.TemplateContacts.Remove(this.SelectedAddrBookContact);
            }
            else if (this.CurAddressBookModel != null && this.CurAddressBookModel.AddressBookContactList != null && this.CurAddressBookModel.AddressBookContactList.Contains(this.SelectedAddrBookContact)) {
                this.CurAddressBookModel.AddressBookContactList.Remove(this.SelectedAddrBookContact);
            }
        }

        private void _copySelectedAddressBookContactToSelectedTemplate() {

            if (CurAddressBookModel.AddressBookContactList.Contains(this.SelectedAddrBookContact) == false ||
                this.SelectedTemplate == null) {

                //In this case, the selected contact is a template contact or there is no selected template--maybe put in some kind of notification later if I feel like it.
                return;
            }

            AddressBookContact clonedContact = this.SelectedAddrBookContact.Clone();
            clonedContact.SourceFileType = SharedEnums.ContactSourceType.TemplateFile;

            this.SelectedTemplate.TemplateContacts.Add(clonedContact);
        }

        private void _copySelectedTemplateContactToAddressBook() {

            if (SelectedTemplate == null || SelectedTemplate.TemplateContacts.Contains(this.SelectedAddrBookContact) == false ||
                this.SelectedTemplate == null) {

                //In this case, the selected contact is not a template contact or there is no selected template--maybe put in some kind of notification later if I feel like it.
                return;
            }

            AddressBookContact clonedContact = this.SelectedAddrBookContact.Clone();
            clonedContact.SourceFileType = SharedEnums.ContactSourceType.AddressBook;

            this.CurAddressBookModel.AddressBookContactList.Add(clonedContact);
        }

        private void _selectANewLogo() {
            if (this.SelectedAddrBookContact == null) {
                return;
            }

            var dlg = new Microsoft.Win32.OpenFileDialog();

            dlg.DefaultExt = ".bmp";
            dlg.Filter = "Image|*.bmp;*.jpg;*.jpeg;*.gif;*.png";
            bool? result = dlg.ShowDialog();

            if (result != null && result == true) {
                string newImagePath = dlg.FileName;
                if (File.Exists(newImagePath) == false) {
                    return;
                }

                SelectedAddrBookContact.PathToLogo = newImagePath;
            }
        }


        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(object sender, PropertyChangedEventArgs e) {

            if (PropertyChanged != null) {
                PropertyChanged(sender, e);
            }
        }

        public void Dispose() {

            this.CurAddressBookModel.Dispose();
            
        }
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
            this._action = action;
        }
    }

    public class ConvertMissingImageToMissingImagePicture : System.Windows.Data.IValueConverter {
        
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {

            string newImagePath = value as string;

            if (newImagePath == null || File.Exists(newImagePath) == false) {
                return "/ITpipes Config;component/Not-Quite-Art_Assets/No_Image_Available.png";
            }

            //Have to actually grab a BitmapImage from the file instead of passing the file path
            //If the file path is passed, the WPF control ends up locking the file, and attempting to delete the file throws an exception.
            return getBmpFromFile(newImagePath);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
            return value;
        }

        private static BitmapImage getBmpFromFile(string sourcePath) {
            BitmapImage returnImage = new BitmapImage();
            using (FileStream imgFileStream = File.OpenRead(sourcePath)) {
                returnImage.BeginInit();
                returnImage.CacheOption = BitmapCacheOption.OnLoad;
                returnImage.StreamSource = imgFileStream;
                returnImage.EndInit();
            }

            return returnImage;
        }
    }

    /*
    //Code Region Template:


    #region Read-Only Resources



    #endregion Read-Only Resources

    #region Private Variables



    #endregion Private Variables

    #region Public Properties



    #endregion Public Properties

    #region Events



    #endregion Events

    #region Constructors



    #endregion Constructors

    #region Private Functions



    #endregion Private Functions

    #region Public Functions



    #endregion Public Functions

    */


}

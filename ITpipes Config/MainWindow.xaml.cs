using System.Windows;
using System.ComponentModel;
using ITpipes_Config.AddressBook;

namespace ITpipes_Config {
    public partial class MainWindow : Window {

        public MainWindow() {

            InitializeComponent();
        }
        
        //Really didn't want to use any code-behind, but because of ITpipes' address book/template wonkiness, I'm adding an event to reselect even if the item is selected.
        //this shouldn't be too awful, since it's contained to the sender, can only be applied to listboxitems, and doesn't rely on the state of my model:
        public void HandleItemReselectedInListBox(object sender, System.Windows.Input.MouseButtonEventArgs e) {

            if (sender == null) {
                return;
            }

            System.Windows.Controls.ListBoxItem clickeditem = sender as System.Windows.Controls.ListBoxItem;

            //clickedItem will be null if it can't cast to a ListBoxItem. Thanks, StackOverflow, for knowing everything!
            if (clickeditem != null &&
                clickeditem.IsSelected) {

                clickeditem.IsSelected = false;
                clickeditem.IsSelected = true;
            }
        }
        
        //There's no decent way I know of to handle window closing events in MVVM. Oh well.
        private void Window_Closing(object sender, CancelEventArgs e) {

            this.Hide(); //don't need people to see the window appear to freeze while the address book contacts and the like are being saved.

            var curViewModel = this.DataContext as ITpipesSettingsViewModel;

            if (curViewModel != null) {

                curViewModel.Settings.SaveSettings();
                curViewModel.Dispose();
            }
            
            if (addrBookViewModelInstance != null) {

                if (addrBookViewModelInstance.CurAddressBookModel != null) {

                    addrBookViewModelInstance.CurAddressBookModel.SaveChangesToFile();

                    foreach(TemplateContact curTemplate in addrBookViewModelInstance.CurAddressBookModel.TemplateContacts) {

                        curTemplate.SaveTemplate();
                    }
                }

                addrBookViewModelInstance.Dispose();
            }
        }
    }
}
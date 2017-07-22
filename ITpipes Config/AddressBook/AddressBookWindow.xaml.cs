using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ITpipes_Config.AddressBook {
    /// <summary>
    /// Interaction logic for AddressBookWindow.xaml
    /// </summary>
    public partial class AddressBookWindow : Window {
        public AddressBookWindow() {
            InitializeComponent();
        }


        //Really didn't want to use any code-behind, but because of our weird address book/template wonkiness, I'm adding an event to reselect even if the item is selected.
        //this shouldn't be too awful, since it's contained to the sender, can only be applied to listboxitems, and doesn't rely on the state of my model:
        public void HandleItemReselectedInListBox(object sender, MouseButtonEventArgs e) {

            if (sender == null) {
                return;
            }

            ListBoxItem clickeditem = sender as ListBoxItem;
            
            //clickedItem will be null if it can't cast to a ListBoxItem. Thanks, StackOverflow, for knowing everything!
            if (clickeditem != null &&
                clickeditem.IsSelected) {

                clickeditem.IsSelected = false;
                clickeditem.IsSelected = true;
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e) {

            //Yeah, yeah--code behind bad.
            //There's no decent way I know of to handle window closing events in MVVM.

            AddressBookViewModel curViewModel = this.DataContext as AddressBookViewModel;
            this.imgSelectedLogo.Source = null; //need to set image source to null to avoid the image control locking the file while it's being deleted/replaced.

            if (curViewModel != null) {
                
                if (curViewModel.CurAddressBookModel != null) {

                    curViewModel.CurAddressBookModel.SaveChangesToFile();
                    
                }

                if (curViewModel.CurAddressBookModel.TemplateContacts != null) {

                    foreach(TemplateContact curContact in curViewModel.CurAddressBookModel.TemplateContacts) {

                        curContact.SaveTemplate();
                    }
                }


                curViewModel.Dispose();
            }
        }
    }
}

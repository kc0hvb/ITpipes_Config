using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ITpipes_Config {
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class AudioRecordingInstructionsWindow : Window {
        public AudioRecordingInstructionsWindow() {
            InitializeComponent();

            string[] testResources = Assembly.GetExecutingAssembly().GetManifestResourceNames();
            Stream logFileStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ITpipes_Config.Resources.ITpipes Text-to-Speech Recording Device Setup.rtf");
            rtBoxAudioCaptureSetupInstructions.Selection.Load(logFileStream, DataFormats.Rtf);
            rtBoxAudioCaptureSetupInstructions.CaretPosition = rtBoxAudioCaptureSetupInstructions.CaretPosition.DocumentStart;
            

            rtBoxAudioCaptureSetupInstructions.PreviewMouseLeftButtonDown += HyperLink_Clicked;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e) {

            //if (sender.Equals(this)) {

            //    this.Hide();

            //    e.Cancel = true;
            //}
            //else {
            //    e.Cancel = false;
            //}

        }

        public void HyperLink_Clicked(object sender, MouseEventArgs e) {

            if (sender.GetType() == typeof(Hyperlink)) {
                var hyperlinkAddress = (Hyperlink)sender;

                System.Diagnostics.Process.Start(hyperlinkAddress.NavigateUri.ToString());
            }
        }

    }
}

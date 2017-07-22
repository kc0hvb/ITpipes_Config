using System;
using System.Windows;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AForge.Video.DirectShow;

namespace ITpipes_Config {
    public class VideoCaptureDevice : DependencyObject {

        public string VideoCaptureDeviceName {
            get { return (string)GetValue(videoCaptureDeviceNameProperty); }
            set { SetValue(videoCaptureDeviceNameProperty, value); }
        }
        public static readonly DependencyProperty videoCaptureDeviceNameProperty =
            DependencyProperty.Register("VideoCaptureDeviceName", typeof(string), typeof(ITpipesSettingsObj), new PropertyMetadata(string.Empty));

        public string VideoCaptureDevicePin {
            get { return (string)GetValue(VideoCaptureDevicePinProperty); }
            set { SetValue(VideoCaptureDevicePinProperty, value); }
        }
        public static readonly DependencyProperty VideoCaptureDevicePinProperty =
            DependencyProperty.Register("VideoCaptureDevicePin", typeof(string), typeof(ITpipesSettingsObj), new PropertyMetadata(string.Empty));


        public FilterInfoCollection AvailableAforgeVideoInputSources {      
            get { return (FilterInfoCollection)GetValue(AvailableAforgeVideoInputSourcesProperty); }
            set { SetValue(AvailableAforgeVideoInputSourcesProperty, value); }
        }
        public static readonly DependencyProperty AvailableAforgeVideoInputSourcesProperty =
            DependencyProperty.Register("AvailableAforgeVideoInputSources", typeof(FilterInfoCollection), typeof(VideoCaptureDevice));


    }
}

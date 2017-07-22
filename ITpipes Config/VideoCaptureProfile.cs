using System;
using System.Windows;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ITpipes_Config.SharedEnums;

namespace ITpipes_Config
{

    public class VideoCaptureProfile : DependencyObject {

        public static char[] NUMERIC_CHARS = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };

        //Public Properties


        public string FriendlyName {
            get { return (string)GetValue(FriendlyNameProperty); }
            set { SetValue(FriendlyNameProperty, value); }
        }
        public static readonly DependencyProperty FriendlyNameProperty =
            DependencyProperty.Register("FriendlyName", typeof(string), typeof(VideoCaptureProfile));
        
        public videoCaptureFormats VideoCaptureFormat {
            get { return (videoCaptureFormats)GetValue(VideoCaptureFormatProperty); }
            set { SetValue(VideoCaptureFormatProperty, value); }
        }

        public string DatabaseSpecifiedName {
            get { return (string)GetValue(DatabaseSpecifiedNameProperty); }
            set { SetValue(DatabaseSpecifiedNameProperty, value); }
        }
        public static readonly DependencyProperty DatabaseSpecifiedNameProperty =
            DependencyProperty.Register("DatabaseSpecifiedName", typeof(string), typeof(VideoCaptureProfile));


        public static readonly DependencyProperty VideoCaptureFormatProperty =
            DependencyProperty.Register("VideoCaptureFormat", typeof(videoCaptureFormats), typeof(VideoCaptureProfile));
        
        public videoCaptureResolutions VideoCaptureResolution {
            get { return (videoCaptureResolutions)GetValue(VideoCaptureResolutionProperty); }
            set { SetValue(VideoCaptureResolutionProperty, value); }
        }
        public static readonly DependencyProperty VideoCaptureResolutionProperty =
            DependencyProperty.Register("VideoCaptureResolution", typeof(videoCaptureResolutions), typeof(VideoCaptureProfile));
        
        public bool UseVariableBitrate {
            get { return (bool)GetValue(UseVariableBitrateProperty); }
            set { SetValue(UseVariableBitrateProperty, value); }
        }
        public static readonly DependencyProperty UseVariableBitrateProperty =
            DependencyProperty.Register("UseVariableBitrate", typeof(bool), typeof(VideoCaptureProfile));
        
        public string VideoCaptureConstantBitrate {
            get { return (string)GetValue(VideoCaptureConstantBitrateProperty); }
            set { SetValue(VideoCaptureConstantBitrateProperty, value); }
        }
        public static readonly DependencyProperty VideoCaptureConstantBitrateProperty =
            DependencyProperty.Register("VideoCaptureConstantBitrate", typeof(string), typeof(ITpipesSettingsObj));

        public int QualityPercent {
            get { return (int)GetValue(QualityPercentProperty); }
            set { SetValue(QualityPercentProperty, value); }
        }
        public static readonly DependencyProperty QualityPercentProperty =
            DependencyProperty.Register("QualityPercent", typeof(int), typeof(VideoCaptureProfile));


        #region Constructors

        public VideoCaptureProfile(string videoCaptureFormat, string videoResolutionProfile, bool useVariableBitrate, int rawBitrate, int qualityPercent) {
            try {
                this.VideoCaptureFormat = getVidCapFormatFromStr(videoCaptureFormat);
                this.VideoCaptureResolution = getVidCapResolutionFromStr(videoResolutionProfile);
                this.UseVariableBitrate = useVariableBitrate;
                this.VideoCaptureConstantBitrate = (Convert.ToSingle(rawBitrate)/1000000f).ToString("0.00");
                this.QualityPercent = qualityPercent;
                this.FriendlyName = getFriendlyName(videoCaptureFormat, this.VideoCaptureResolution);
                this.DatabaseSpecifiedName = videoCaptureFormat + " (" + videoResolutionProfile + ")";
            }
            catch {
                throw;
            }

        }

        #endregion

        //internal functions

        private static videoCaptureFormats getVidCapFormatFromStr(string inputVideoFormatStr) {
            switch (inputVideoFormatStr) {
                case ("WMV"):
                    {
                        return videoCaptureFormats.WMV7;
                    }
                case ("WMV9"):
                    {
                        return videoCaptureFormats.WMV9;
                    }
                case ("MPEG 1"):
                    {
                        return videoCaptureFormats.MPEG1;
                    }
                case ("MPEG 2"):
                    {
                        return videoCaptureFormats.MPEG2;
                    }
                case ("MPEG 4"):
                    {
                        return videoCaptureFormats.MPEG4;
                    }
                default:
                    {
                        throw new Exception(string.Format("Video Format Not Recognized:\n{0}", inputVideoFormatStr));
                    }
            }
        }
        
        private static videoCaptureResolutions getVidCapResolutionFromStr(string inputVideoResProfile) {
            switch (inputVideoResProfile) {
                case ("CD"):
                    {
                        return videoCaptureResolutions.CD;
                    }
                case ("DVD"):
                    {
                        return videoCaptureResolutions.DVD;
                    }
                default:
                    {
                        throw new Exception(string.Format("Video Capture Resolution Not Recognized:\n{0}", inputVideoResProfile));
                    }
            }
        }

        private static string getFriendlyName(string inpVideoFormat, videoCaptureResolutions inpVideoResolution) {
            string resolutionComponentName;
            
            switch (inpVideoResolution) {
                case (videoCaptureResolutions.CD):
                    {
                        resolutionComponentName = "(320x240)";
                        break;
                    }
                case (videoCaptureResolutions.DVD):
                    {
                        resolutionComponentName = "(640x480)";
                        break;
                    }
                default:
                    {
                        throw new Exception(string.Format("videoCaptureResolutions enum was of unrecognized value:\n{0}", inpVideoResolution.ToString()));
                    }
            }

            return (string.Format("{0} {1}", inpVideoFormat.Trim(), resolutionComponentName));
        }

        //Validation Callbacks:

        public static bool IsValidCbrFloat(object inputValue) {

            if (inputValue == null) {
                return false;
            }

            float inputValueCasted = -1234.56f;

            string inputValueString;

            try {
                inputValueString = inputValue.ToString();
            }
            catch {
                return false;
            } 

            //Adding a 0 onto the end because otherwise "1." would evaluate to false, meaning decimals could only be placed before the last char in the string.
            //This also means that in the writeout for this property "x." needs to write out as "x"
            if (float.TryParse(inputValue.ToString() + '0', out inputValueCasted) && inputValueCasted < 10.01 && inputValueCasted > 0.5) {
                return true;
            }

            return false;
        }

        public static bool IsValidWholeNumberPercent(object inputValue) {

            if (inputValue == null) {
                return false;
            }

            int inputCastToInt;

            if (int.TryParse(inputValue.ToString(), out inputCastToInt)) {
                
                if (inputCastToInt < 101 && inputCastToInt >= 0) {
                    return true;
                }

            }

            return false;
        }

        public static void WholeNumberPercentPropertyChangedCallback(DependencyObject d, DependencyPropertyChangedEventArgs e) {

            if (e.NewValue == null) {
                d.SetValue(e.Property, (int)0);
                return;
            }

            int parsedIntValue = -1;

            if (int.TryParse(e.NewValue.ToString(), out parsedIntValue) &&
                parsedIntValue >= 0 &&
                parsedIntValue <= 100 ) {

                d.SetValue(e.Property, e.NewValue);
            }

        }
        


    }
}

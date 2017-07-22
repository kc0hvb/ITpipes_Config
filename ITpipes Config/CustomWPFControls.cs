using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ITpipes_Config.CustomControls
{

    public class TextBoxWholeNumberPercent : TextBox {

        string lastValidValue {
            get; set;
        } = "0";

        public static char[] NUMERIC_CHARS = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };

        public TextBoxWholeNumberPercent() {
            
            this.PreviewKeyDown += PreviewKeyDownHandler;
            this.PreviewTextInput += PreviewTextInputHandler;
            this.TextChanged += TextChangedHandler;

            DataObject.AddPastingHandler(this, PasteHandler);
        }

        public static void PreviewKeyDownHandler(object sender, KeyEventArgs e) {

            if (e.Key == Key.Space) {
                e.Handled = true;
            }

        }

        public static void TextChangedHandler(object sender, TextChangedEventArgs e) {

            //validation on input is handled in the PreviewTextInput event, so all this needs to do is set to 0 if text has been changed to null--null is not valid percent.

            TextBoxWholeNumberPercent thisControl = sender as TextBoxWholeNumberPercent;

            int caretIndexAtStart = thisControl.CaretIndex;

            if (thisControl.Text.Length == 0) {
                thisControl.Text = "0";
                thisControl.CaretIndex = 1;
                thisControl.lastValidValue = "0";
                e.Handled = true;
            } 
            
            else {
                string modifiedText = new string(thisControl.Text.ToCharArray());

                while (modifiedText[0] == '0') {
                    if (modifiedText.Length == 1) {
                        break;
                    }

                    modifiedText = modifiedText.Substring(1);

                }

                int parsedNewText = -1;

                if (int.TryParse(modifiedText, out parsedNewText)) {

                    if (parsedNewText > 100) {
                        thisControl.Text = "100";
                        thisControl.CaretIndex = 3;
                        thisControl.lastValidValue = "100";
                        //e.Handled = true;
                        return;
                    }

                    else if (parsedNewText < 0) {
                        thisControl.Text = "0";
                        thisControl.CaretIndex = 1;
                        thisControl.lastValidValue = "0";
                        //e.Handled = true;
                        return;
                    }

                    thisControl.Text = modifiedText;
                    
                    thisControl.lastValidValue = modifiedText;
                    e.Handled = true;
                    return;
                }

                else {
                    thisControl.Text = thisControl.lastValidValue;
                    thisControl.CaretIndex = thisControl.lastValidValue.Length;
                }
            }

        }

        public static void PreviewTextInputHandler(object sender, TextCompositionEventArgs e) {



        }

        public void PasteHandler(object sender, DataObjectPastingEventArgs e) {

            string textBeingPasted = e.DataObject.GetData(DataFormats.Text) as string;

            if (IsValidPositiveWholeNumberPercentValue(textBeingPasted) == false) {
                e.Handled = true;
            }

        }
        
        private static bool IsValidPositiveWholeNumberPercentValue(string input) {

            if (input == null || input.Length == 0) {
                return false;
            }

            for (int i = 0; i < input.Length; i++) {
                if (NUMERIC_CHARS.Contains(input[i]) == false) {
                    return false;
                }
            }

            int parsedInput = int.Parse(input);

            if (parsedInput > 100) { //parsedInput can't be negative at this point--the minus symbol would be filtered out in the NUMERIC_CHARS comparison.
                return false;
            }

            return true;
        }

    }



    public class TextBoxSerialPortNumber : TextBox {


        string lastValidValue {
            get; set;
        }
        = "0";

        public static char[] NUMERIC_CHARS = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };

        public TextBoxSerialPortNumber() {

            this.PreviewKeyDown += PreviewKeyDownHandler;
            this.PreviewTextInput += PreviewTextInputHandler;
            this.TextChanged += TextChangedHandler;

            DataObject.AddPastingHandler(this, PasteHandler);
        }

        public static void PreviewKeyDownHandler(object sender, KeyEventArgs e) {

            if (e.Key == Key.Space) {
                e.Handled = true;
            }

        }

        public static void TextChangedHandler(object sender, TextChangedEventArgs e) {

            //validation on input is handled in the PreviewTextInput event, so all this needs to do is set to 0 if text has been changed to null--null is not valid percent.

            TextBoxSerialPortNumber thisControl = sender as TextBoxSerialPortNumber;

            int caretIndexAtStart = thisControl.CaretIndex;

            //if (thisControl.Text.Length == 0) {
            //    thisControl.Text = "1";
            //    thisControl.CaretIndex = 1;
            //    thisControl.lastValidValue = "1";
            //    e.Handled = true;
            //}
            
                string modifiedText = new string(thisControl.Text.ToCharArray());

            if (modifiedText.Length > 0) {

                while (modifiedText[0] == '0') {
                    if (modifiedText.Length == 1) {
                        break;
                    }

                    modifiedText = modifiedText.Substring(1);

                }

                int parsedNewText = -1;

                if (int.TryParse(modifiedText, out parsedNewText)) {

                    if (parsedNewText > 100) {
                        thisControl.Text = "1";
                        thisControl.CaretIndex = 1;
                        thisControl.lastValidValue = "1";
                        //e.Handled = true;
                        return;
                    }

                    else if (parsedNewText < 1) {
                        thisControl.Text = "1";
                        thisControl.CaretIndex = 1;
                        thisControl.lastValidValue = "1";
                        //e.Handled = true;
                        return;
                    }

                    thisControl.Text = modifiedText;

                    thisControl.lastValidValue = modifiedText;
                    e.Handled = true;
                    return;
                }

                else {
                    thisControl.Text = thisControl.lastValidValue;
                    thisControl.CaretIndex = thisControl.lastValidValue.Length;
                }
            }

        }

        public static void PreviewTextInputHandler(object sender, TextCompositionEventArgs e) {



        }

        public void PasteHandler(object sender, DataObjectPastingEventArgs e) {

            string textBeingPasted = e.DataObject.GetData(DataFormats.Text) as string;

            if (IsValidSerialPortNumber(textBeingPasted) == false) {
                e.Handled = true;
            }

        }

        private static bool IsValidSerialPortNumber(string input) {

            if (input == null || input.Length == 0 || input.Length > 2) {
                return false;
            }

            for (int i = 0; i < input.Length; i++) {
                if (NUMERIC_CHARS.Contains(input[i]) == false) {
                    return false;
                }
            }

            int parsedInput = int.Parse(input);

            if (parsedInput > 100 || parsedInput < 1) { //parsedInput can't be negative at this point--the minus symbol would be filtered out in the NUMERIC_CHARS comparison.
                return false;
            }

            return true;
        }

    }












    //public class TextBoxInt : TextBox
    //{
    //    public static string[] validNumberChars = new string[] { "", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" };

    //    protected override void OnTextInput(TextCompositionEventArgs e)
    //    {
    //        if (validNumberChars.Contains(e.Text) == false)
    //        {
    //            e.Handled = true;
    //        }
    //        else
    //        {
    //            base.OnTextInput(e);
    //        }
    //    }

    //    protected override void OnTextChanged(TextChangedEventArgs e)
    //    {

    //        int tempInt = 0;
    //        if (int.TryParse(this.Text, out tempInt))
    //        {
    //            base.OnTextChanged(e);
    //        }
    //        else
    //        {
    //            e.Handled = true;
    //            return;
    //        }

    //    }
    //}

    //public class TextBoxFloat : TextBox
    //{
    //    public static string[] validText = new string[] { "", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "."};

    //    protected override void OnPreviewTextInput(TextCompositionEventArgs e)
    //    {
    //        TextBoxFloat curTextBoxFloat = e.OriginalSource as TextBoxFloat;

    //        if (e.Text == "." && curTextBoxFloat.Text.Contains('.') == true)
    //        {
    //            e.Handled = true;
    //            return;
    //        }

    //        if (validText.Contains(e.Text) == false)
    //        {
    //            e.Handled = true;
    //            return;
    //        }

    //        base.OnTextInput(e);
    //    }

    //    protected override void OnTextChanged(TextChangedEventArgs e)
    //    {
    //        float convertedTextValue = 0F;
    //            TextBoxFloat curTextBoxFloat = this as TextBoxFloat;
    //            if (float.TryParse(curTextBoxFloat.Text + "0", out convertedTextValue) == true)
    //            {
    //                base.OnTextChanged(e);
    //            }
    //            else
    //            {
    //                curTextBoxFloat.Text = "";
    //                e.Handled = true;
    //                return;
    //            }
    //    }
    //}
}

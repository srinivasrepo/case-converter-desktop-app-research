using System;
using System.Windows;
using System.Windows.Forms;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices;
using Clipboard = System.Windows.Clipboard;
using System.Threading.Tasks;

namespace App
{
    class CaseConverter
    {
        //private IKeyboardMouseEvents globalMouseHook;

        [STAThread]
        static void Main(string[] args)
        {
            // See https://aka.ms/new-console-template for more information

            // Save clipboard's current content to restore it later.
            System.Windows.IDataObject tmpClipboard = Clipboard.GetDataObject();

            Clipboard.Clear();

            //// I think a small delay will be more safe.
            //// You could remove it, but be careful.

            Task.Delay(1000);
            //// Send Ctrl+C, which is "copy"
            //System.Windows.Forms.SendKeys.SendWait("%~");
            System.Windows.Forms.SendKeys.SendWait("^c");

            //// Same as above. But this is more important.
            //// In some softwares like Word, the mouse double click will not svelect the word you clicked immediately.
            //// If you remove it, you will not get the text you selected.
            //await Task.Delay(50);

            if (System.Windows.Clipboard.ContainsText())
            {
                string text = System.Windows.Clipboard.GetText();

                // Your code
                System.Windows.Clipboard.SetText(text.ToUpper());

            }
            else
            {
                // Restore the Clipboard.
                System.Windows.Clipboard.SetText("no text in selected");
            }

            System.Windows.Forms.SendKeys.SendWait("^v");

        }

    }




}









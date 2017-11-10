using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using Microsoft.Office.Interop.Word;
namespace WpfApplication1
{
    class ImageToWordHelper //:IDisposable
    {
        object miss = System.Reflection.Missing.Value;
        object otrue = true;
        object ofalse = false;
         Microsoft.Office.Interop.Word.Application WordApp = new Microsoft.Office.Interop.Word.Application();
        Document doc = null;
        
        public ImageToWordHelper(){
            try
            {
                WordApp.Documents.Add();
                WordApp.Visible = true;
                doc = WordApp.ActiveDocument;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(" " + ex.Message);
            }
        }
        public void PastetoWord(System.Drawing.Image image, System.Drawing.Imaging.ImageFormat format)
        {
            try
            {
                //Byte[] imageCapture = ToByteArray(bitmap, System.Drawing.Imaging.ImageFormat.Bmp);
                
                System.Windows.Clipboard.SetDataObject(image, true);
                object start = 0;
                object end = 0;
                Microsoft.Office.Interop.Word.Range rng = doc.Range(ref start, ref end);
                rng.Paste();
                if (image != null )
                {
                    ((IDisposable)image).Dispose();
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(" " + ex.Message);
            }
        }
        public void SaveWordAndClose()
        {
            doc.Save();
            System.Windows.MessageBox.Show("Word created ");
            WordApp.Quit(ref otrue, ref miss, ref miss);
            /*if (WordApp != null || doc!=null)
            {
                ((IDisposable)WordApp).Dispose();
                ((IDisposable)doc).Dispose();
            }*/
            WordApp = null;
            doc = null;
        }
    }
}

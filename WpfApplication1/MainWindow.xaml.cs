using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using System.Drawing;
using System.Threading;
using System.IO;
using Microsoft.Office.Interop.Word;
namespace WpfApplication1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public  partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        public Bitmap imageCapture()
        {
            Bitmap bitmap = new Bitmap(Screen.PrimaryScreen.Bounds.Width,
                                    Screen.PrimaryScreen.Bounds.Height);
            Graphics graphics = Graphics.FromImage(bitmap);
            graphics.CopyFromScreen(0, 0, 0, 0, bitmap.Size);
            this.WindowStyle = WindowStyle.SingleBorderWindow;
            this.WindowState = (WindowState)FormWindowState.Normal;
            return bitmap;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            CaptureButtonOverride();
         }
        ImageToWordHelper wordDocCreate = null;
        private void CaptureButtonOverride()
        {
            try
            {
                minimizeWindow();
            }
            finally
            {
                Bitmap bitmapCapture = imageCapture();
                if (wordDocCreate == null)
                {
                    wordDocCreate = new ImageToWordHelper();
                    wordDocCreate.PastetoWord(bitmapCapture, System.Drawing.Imaging.ImageFormat.Bmp);
                }
                else
                {
                    wordDocCreate.PastetoWord(bitmapCapture, System.Drawing.Imaging.ImageFormat.Bmp);
                }
            }
        }
        private void SaveButtonOverride(ImageToWordHelper wordDocCreate)
        {
            wordDocCreate.SaveWordAndClose();
            /*if (wordDocCreate != null)
            {
                ((IDisposable)wordDocCreate).Dispose();
            }*/
            
        }
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            SaveButtonOverride(wordDocCreate);
            wordDocCreate = null;
        }
        public void minimizeWindow()
        {
            this.WindowStyle = WindowStyle.None;
            this.WindowState = (WindowState)FormWindowState.Minimized;
        }
       
    }
    
}

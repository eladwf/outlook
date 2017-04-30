using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        private Office.CommandBar menuBar;
        private Office.CommandBarPopup newMenuBar;
        private Office.CommandBarButton buttonOne;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            AddMenuBar();
        }

        private void AddMenuBar()
        {
            try
            {
                menuBar = this.Application.ActiveExplorer().CommandBars.ActiveMenuBar;
                newMenuBar = (Office.CommandBarPopup)menuBar.Controls.Add(
                    Office.MsoControlType.msoControlPopup, missing,
                    missing, missing, true);
                if (newMenuBar != null)
                {
                    buttonOne = (Office.CommandBarButton)
                        newMenuBar.Controls.
                    Add(Office.MsoControlType.msoControlButton, System.
                        Type.Missing, System.Type.Missing, 1, true);
                    newMenuBar.Caption = "Highlight";
                    buttonOne.Style = Office.MsoButtonStyle.msoButtonIconAndCaption;
                    buttonOne.Caption = "Highlight elad";
                    buttonOne.FaceId = 100;
                    buttonOne.Tag = "c123";
                    buttonOne.Picture = getImage();
                    buttonOne.Click += ButtonOne_Click;
                    newMenuBar.Visible = true;
                }
            }
        
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private void ButtonOne_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                if (this.Application.ActiveExplorer().Selection.Count > 0)
                {
                    Object selObject = this.Application.ActiveExplorer().Selection[1];
                    if (selObject is Outlook.MailItem)
                    {
                        Outlook.MailItem mailItem =
                            (selObject as Outlook.MailItem);
                        FindAndHighlight(mailItem, "elad");
                    }
                }
            }
            catch{ return; }
            
        }
        sealed public class ConvertImage : System.Windows.Forms.AxHost
        {
            private ConvertImage()
                : base(null)
            {
            }

            public static stdole.IPictureDisp Convert
                (System.Drawing.Image image)
            {
                return (stdole.IPictureDisp)System.
                    Windows.Forms.AxHost
                    .GetIPictureDispFromPicture(image);
            }
        }
        private stdole.IPictureDisp getImage()
        {
            stdole.IPictureDisp tempImage = null;
            try
            {
                System.Drawing.Image newIcon =
                    Properties.Resources.star_icon;

                System.Windows.Forms.ImageList newImageList =
                    new System.Windows.Forms.ImageList();
                newImageList.Images.Add(newIcon);
                tempImage = ConvertImage.Convert(newImageList.Images[0]);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            return tempImage;
        }

        public void FindAndHighlight(Outlook.MailItem mailItem, string text)
        {
            if (mailItem != null)
            {
                if (mailItem.Body.Contains(text))
                {
                    string output = @"<font color=""red"">" + text + "</font>";
                    mailItem.HTMLBody = mailItem.HTMLBody.Replace("elad", output);
                    Outlook.Folder fd = (Outlook.Folder)Application.ActiveExplorer().CurrentFolder;
                    var customCat = "Elad";
                    if (Application.Session.Categories[customCat] == null)
                        Application.Session.Categories.Add(customCat, Outlook.OlCategoryColor.olCategoryColorDarkRed);
                    mailItem.Categories = customCat;
                    mailItem.Save();
                }
            }
        }
    
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
        }
        
        #endregion
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn_OneClickReSend
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility のメンバー

        public string GetCustomUI(string ribbonID)
        {
            string ribbonXML = String.Empty;

            //メール一覧画面のみ、アドインを表示する。
            if (ribbonID == "Microsoft.Outlook.Explorer")
            {
                return GetResourceText("OutlookAddIn_OneClickReSend.Ribbon.xml");
            }

            return ribbonXML;
        }

        #endregion

        #region リボンのコールバック
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnResendButton(Office.IRibbonControl control)
        {
            //再送ボタン押下時の処理
            var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
            foreach (var item in explorer.Selection)
            {
                //選択数分処理
                if (item is Outlook.MailItem)
                {
                    //再送対象のMailItemを取得
                    var selectMailItem = item as Outlook.MailItem;

                    //送信用のMailItemを作成
                    Outlook.MailItem sendMailItem = Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem);

                    //再送対象MailItemの要素を送信用MailItemの要素にコピー
                    sendMailItem.To = selectMailItem.To;
                    sendMailItem.CC = selectMailItem.CC;
                    sendMailItem.BCC = selectMailItem.BCC;
                    sendMailItem.Subject = selectMailItem.Subject;
                    sendMailItem.BodyFormat = selectMailItem.BodyFormat;
                    sendMailItem.Recipients.ResolveAll();

                    if (selectMailItem.BodyFormat == Outlook.OlBodyFormat.olFormatPlain)
                    {
                        //本文がテキスト形式の場合、Body要素をコピーする。
                        sendMailItem.Body = selectMailItem.Body;
                    }
                    else if (selectMailItem.BodyFormat == Outlook.OlBodyFormat.olFormatHTML)
                    {
                        //本文がHTML形式の場合、HTMLBody要素をコピーする。
                        sendMailItem.HTMLBody = selectMailItem.HTMLBody;
                    }
                    else
                    {
                        //本文がリッチテキスト形式や指定なしの場合は非サポート
                        ;
                    }

                    //コピーが終わったら送信するMailItemを表示
                    sendMailItem.Display(false);
                }
            }
        }

        #endregion

        #region ヘルパー

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}

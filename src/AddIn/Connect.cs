// Word plaginining kirish nuqtasi: Word uni ishga tushganda yuklaydi va lentaga
// (Home yorlig'i) "Yangi lotinga" tugmasini qo'shadi.
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Alifbo.Word
{
    [ComVisible(true)]
    [Guid(Connect.ClassId)]
    [ProgId(Connect.ProgIdValue)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        public const string ClassId = "5D920E66-8009-4A2B-8956-325639112835";
        public const string ProgIdValue = "Alifbo.Connect";

        private object _application;

        private const string RibbonXml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
            "<ribbon><tabs><tab idMso=\"TabHome\">" +
            "<group id=\"AlifboGroup\" label=\"Alifbo\" insertAfterMso=\"GroupEditing\">" +
            "<button id=\"AlifboConvert\" size=\"large\" imageMso=\"ReplaceDialog\" keytip=\"YL\"" +
            " label=\"Yangi lotinga\"" +
            " screentip=\"Yangi lotin alifbosiga o\u02BBgirish\"" +
            " supertip=\"Bir bosishda: matn belgilangan bo\u02BBlsa \u2014 faqat o\u02BBsha qism, aks holda butun hujjat " +
            "eski lotin yoki kirilldan yangi lotinga (\u00F6, \u011F, \u015F, \u00E7) o\u02BBgiriladi. " +
            "Formatlar saqlanadi. Bekor qilish: Ctrl+Z.\"" +
            " onAction=\"OnConvert\"/>" +
            "</group></tab></tabs></ribbon></customUI>";

        // ---------------------------------------------------------- IDTExtensibility2

        public void OnConnection(object application, ExtConnectMode connectMode, object addInInst, ref Array custom)
        {
            _application = application;
        }

        public void OnDisconnection(ExtDisconnectMode removeMode, ref Array custom)
        {
            if (_application != null && Marshal.IsComObject(_application))
            {
                try { Marshal.ReleaseComObject(_application); } catch { }
            }
            _application = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom) { }

        public void OnBeginShutdown(ref Array custom) { }

        // ------------------------------------------------------- IRibbonExtensibility

        public string GetCustomUI(string ribbonId)
        {
            return RibbonXml;
        }

        // Lenta tugmasi (Office uni IDispatch orqali nomi bo'yicha chaqiradi).
        public void OnConvert(object control)
        {
            try
            {
                if (_application == null) return;
                new WordConverter(_application).Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Xatolik yuz berdi: " + ex.Message +
                                "\r\nQilingan o\u02BBzgarishlarni Ctrl+Z bilan bekor qilishingiz mumkin.",
                                "Alifbo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (control != null && Marshal.IsComObject(control))
                {
                    try { Marshal.ReleaseComObject(control); } catch { }
                }
            }
        }
    }
}

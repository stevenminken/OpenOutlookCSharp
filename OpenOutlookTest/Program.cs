
using Stimulsoft.Report;
using Stimulsoft.Report.Export;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

Outlook.Application application = GetApplicationObject();
CreateSendItem(application);

/*
 *https://admin.stimulsoft.com/documentation/classreference-dbs/html/ca9a51c1-d3a5-79f9-70f1-a55470295f76.htm
 *Exporteert Stimulsoft mrt bestand naar een rtf bestand en verstuurd deze via outlook
 */
void CreateSendItem(Outlook.Application oApp)
{
    Outlook.MailItem mailItem = null;
    Outlook.Recipients mailRecipients = null;
    Outlook.Recipient mailRecipient = null;
    StiReport report;

    try
    {
        //creert nieuwe instantie van StiReport
        report = new StiReport();
        report.Load("c:\\Users\\Steven\\Documents\\Reports\\KleinBevestigingDatumTijd.mrt");
        report.Render(false);

        //Schrijft rtf file naar bestand
        StiRtfExportSettings rtfSettings = new StiRtfExportSettings();
        MemoryStream memoryStream = new MemoryStream();
        rtfSettings.ImageQuality = 1.0f;
        report.ExportDocument(StiExportFormat.Rtf, memoryStream, rtfSettings);

        Console.WriteLine("The export action is complete.", "Export Report");

        mailItem = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
        mailItem.To = "stevenminken@hotmail.com";
        mailItem.Subject = "A programatically generated e-mail";
        mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatRichText;
        mailItem.RTFBody = memoryStream.ToArray();
        memoryStream.Flush();
        mailItem.Display();

        //HTML code ter info, rendered niet goed
        //StiHtmlExportSettings htmlSettings = new StiHtmlExportSettings();
        //report.ExportDocument(StiExportFormat.Html, "c:\\Users\\Steven\\Documents\\Reports\\KleinBevestigingDatumTijd.html", htmlSettings);

        //mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
        //mailItem.HTMLBody = File.ReadAllText("c:\\Users\\Steven\\Documents\\Reports\\KleinBevestigingDatumTijd.html");
    }
    catch (Exception ex)
    {
        Console.WriteLine(ex.ToString());
    }
    finally
    {
        if (mailRecipient != null) Marshal.ReleaseComObject(mailRecipient);
        if (mailRecipients != null) Marshal.ReleaseComObject(mailRecipients);
        if (mailItem != null) Marshal.ReleaseComObject(mailItem);
    }
}

/*
 * Vraagt de actieve outlook instantie op en retourneert deze
 */
Outlook.Application GetApplicationObject()
{
    Outlook.Application application = null;

    // Check whether there is an Outlook process running.
    if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
    {

        // If so, use the GetActiveObject method to obtain the process and cast it to an Application object.
        // application = (Outlook.Application)Marshal.GetActiveObject("Outlook.Application");
        application = new Outlook.Application();
    }

    else
    {
        // If not, create a new instance of Outlook and sign in to the default profile.
        application = new Outlook.Application();
        Outlook.NameSpace nameSpace = application.GetNamespace("MAPI");
        nameSpace.Logon("", "", Missing.Value, Missing.Value);
        nameSpace = null;
    }

    // Return the Outlook Application object.
    return application;
}

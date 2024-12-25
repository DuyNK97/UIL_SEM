using System;
using System.Collections.Generic;
using LabelManager2;
using NPOI.SS.Formula.PTG;
using QRCoder;
using static ClosedXML.Excel.XLPredefinedFormat;

namespace CIM
{
    public class LabelPrintingService
    {
        private Application lblApp;
        private Document lblDoc;

        public LabelPrintingService()
        {
            lblApp = new Application();
            lblDoc = lblApp.Documents.Open(@"C:\Label\label.Lab", false); // Open label file
        }

        //public bool PrintLabel(string labPath, string fileName, int copies, List<LabPrintVar> labVariableList, bool isClearDefault)
        //{
        //    try
        //    {
        //        labPath = labPath.Trim().EndsWith("\\") ? labPath.Trim() : $"{labPath.Trim()}\\";
        //        Document docLabel = lblApp.Documents.Open(labPath + fileName);
        //        if (docLabel == null)
        //        {
        //            Console.WriteLine("Label Path not found");
        //            return false;
        //        }

        //        // Connect to printer (assuming ZebraZT411 is hardcoded for now)
        //        docLabel.Printer.SwitchTo("ZebraZT411", "TCP/IP", true);

        //        // Prepare variables
        //        Dictionary<string, string> varsInLocalLabelDic = new Dictionary<string, string>();
        //        foreach (LabPrintVar labVar in labVariableList)
        //        {
        //            varsInLocalLabelDic[labVar.LabVar] = labVar.Value;
        //        }

        //        // Iterate through form variables in the label document
        //        int count = docLabel.Variables.FormVariables.Count;
        //        for (int j = 1; j <= count; j++)
        //        {
        //            string varName = docLabel.Variables.FormVariables.Item(j).Name;

        //            // Clear default values if needed
        //            if (isClearDefault && varsInLocalLabelDic.ContainsKey(varName))
        //            {
        //                docLabel.Variables.FormVariables.Item(varName).Value = "";
        //            }

        //            // Set variable values based on input list
        //            if (varsInLocalLabelDic.ContainsKey(varName))
        //            {
        //                docLabel.Variables.FormVariables.Item(varName).Value = varsInLocalLabelDic[varName];
        //            }
        //        }

        //        // Print the document
        //        if (!Print(copies, docLabel))
        //        {
        //            return false;
        //        }

        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Error printing label: " + ex.Message);
        //        return false;
        //    }
        //    finally
        //    {
        //        // Close the label document
        //        if (lblDoc != null)
        //        {
        //            lblDoc.Close(false); // Close document without saving changes
        //            lblDoc = null;
        //        }
        //    }
        //}

        public bool PrintLabel(string labPath, string fileName, int copies, List<LabPrintVar> labVariableList, bool isClearDefault)
        {
            try
            {
                labPath = labPath.Trim().EndsWith("\\") ? labPath.Trim() : $"{labPath.Trim()}\\";
                Document docLabel = lblApp.Documents.Open(labPath + fileName);
                if (docLabel == null)
                {
                    Console.WriteLine("Label Path not found");
                    return false;
                }

                // Connect to printer at 192.168.3.106, port 9100
                docLabel.Printer.SwitchTo("Printer411", "9100", true);

                // Prepare variables
                Dictionary<string, string> varsInLocalLabelDic = new Dictionary<string, string>();
                foreach (LabPrintVar labVar in labVariableList)
                {
                    varsInLocalLabelDic[labVar.LabVar] = labVar.Value;
                }

                // Iterate through form variables in the label document
                int count = docLabel.Variables.FormVariables.Count;
                for (int j = 1; j <= count; j++)
                {
                    string varName = docLabel.Variables.FormVariables.Item(j).Name;

                    // Clear default values if needed
                    if (isClearDefault && varsInLocalLabelDic.ContainsKey(varName))
                    {
                        docLabel.Variables.FormVariables.Item(varName).Value = "";
                    }

                    // Set variable values based on input list
                    if (varsInLocalLabelDic.ContainsKey(varName))
                    {
                        docLabel.Variables.FormVariables.Item(varName).Value = varsInLocalLabelDic[varName];
                    }
                }

                // Print the document
                if (!Print(copies, docLabel))
                {
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error printing label: " + ex.Message);
                return false;
            }
            finally
            {
                // Close the label document
                if (lblDoc != null)
                {
                    lblDoc.Close(false); // Close document without saving changes
                    lblDoc = null;
                }
            }
        }


        private bool Print(int copies, Document docLabel)
        {
            try
            {

                docLabel.PrintDocument(copies);
                return true;


            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi trong quá trình in: " + ex.Message);
                return false;
            }
        }


        public void Close()
        {
            if (lblDoc != null)
            {
                lblDoc.Close(false); // Close document without saving changes
                lblDoc = null;
            }
            if (lblApp != null)
            {
                lblApp.Quit();
                lblApp = null;
            }
        }
    }
}

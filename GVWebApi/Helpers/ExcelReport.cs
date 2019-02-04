﻿using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2013.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using GVWebapi.RemoteData;
using GVWebapi.Models;
using System.Data.SqlClient;

namespace GVWebapi.Helpers.Reporting
{
    public class ExcelReport
    {

        // Static Var's
        //
        private int _contractID;
        private string _period;
        private string _customer;
        private int _customerID;
        private string _startDate;
        private int _invoiceID;
        private string _revisionTotalString;
        private string _rolloverTotalString;
        private string _costavoidanceTotalString;
        private string _replacementTotalsString;
        private string _overrideDate;
        private UInt32Value _summaryindex;
        // Creates a SpreadsheetDocument.
        public void CreatePackage(string filePath,int customerid,string customer, string period, int contractID,string startDate,int invoiceID,string overrideDate)
        {
            _customerID = customerid;
            _period = period;
            _customer = customer;
            _contractID = contractID;
            _startDate = startDate;
            _invoiceID = invoiceID;
            _overrideDate = overrideDate;
         
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                CreateParts(package);
                FixFormating(package);
                package.Close();
            }
        }
        private void FixFormating(SpreadsheetDocument document)
        {
            
                WorkbookPart workbookPart = document.WorkbookPart;
                IEnumerable<string> worksheetIds = workbookPart.Workbook.Descendants<Sheet>().Select(w => w.Id.Value);
                WorksheetPart worksheetPart;
                foreach (string worksheetId in worksheetIds)
                {
                    worksheetPart = ((WorksheetPart)workbookPart.GetPartById(worksheetId));
                    PageSetup pageSetup = worksheetPart.Worksheet.Descendants<PageSetup>().FirstOrDefault();
                    if (pageSetup != null)
                    {
                        pageSetup.Orientation = OrientationValues.Landscape;
                    }
                    else
                    {
                        pageSetup = new PageSetup();
                        pageSetup.Orientation = OrientationValues.Landscape;
                        pageSetup.FitToWidth = 1;
                        worksheetPart.Worksheet.AppendChild(pageSetup);
                    }
                    worksheetPart.Worksheet.Save();
                }
                workbookPart.Workbook.Save();
            
        }
        // Adds child parts and generates content of the specified part.
        private void CreateParts(SpreadsheetDocument document)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPart1Content(workbookPart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId8");
            GenerateWorkbookStylesPart1Content(workbookStylesPart1);
//Volume Trend
//
            WorksheetPart worksheetPart122 = workbookPart1.AddNewPart<WorksheetPart>("rId18");
            GenerateWorksheetVolumeTrend(worksheetPart122);

            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart122 = worksheetPart122.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
            GenerateSpreadsheetPrinterSettingsPart1Content(spreadsheetPrinterSettingsPart122);

            WorksheetCommentsPart worksheetCommentsPart122 = worksheetPart122.AddNewPart<WorksheetCommentsPart>("rId4");
            GenerateWorksheetCommentsPart1Content(worksheetCommentsPart122);


            VmlDrawingPart vmlDrawingPart122 = worksheetPart122.AddNewPart<VmlDrawingPart>("rId3");
            GenerateVmlDrawingPart10Content(vmlDrawingPart122);

            ImagePart imagePart122 = vmlDrawingPart122.AddNewPart<ImagePart>("image/jpeg", "rId2");
            GenerateImagePart1Content(imagePart122);

            ImagePart imagePart222 = vmlDrawingPart122.AddNewPart<ImagePart>("image/jpeg", "rId1");
            GenerateImagePart2Content(imagePart222);

            VmlDrawingPart vmlDrawingPart222 = worksheetPart122.AddNewPart<VmlDrawingPart>("rId2");
            GenerateVmlDrawingPart10Content(vmlDrawingPart222);

            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart1 = worksheetPart122.AddNewPart<SpreadsheetPrinterSettingsPart>("rId18");
            GenerateSpreadsheetPrinterSettingsPart1Content(spreadsheetPrinterSettingsPart1);
//Cost Avoidance
//
            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId11");
            GenerateWorksheetPartCostAvoidance(worksheetPart1);

            VmlDrawingPart vmlDrawingPart1 = worksheetPart1.AddNewPart<VmlDrawingPart>("rId3");
          

            ImagePart imagePart1 = vmlDrawingPart1.AddNewPart<ImagePart>("image/jpeg", "rId2");
            GenerateImagePart1Content(imagePart1);

            ImagePart imagePart2 = vmlDrawingPart1.AddNewPart<ImagePart>("image/jpeg", "rId1");
            GenerateImagePart2Content(imagePart2);

            VmlDrawingPart vmlDrawingPart2 = worksheetPart1.AddNewPart<VmlDrawingPart>("rId2");
            GenerateVmlDrawingPart10Content(vmlDrawingPart2);
            GenerateVmlDrawingPart10Content(vmlDrawingPart1);
            ThemePart themePart1 = workbookPart1.AddNewPart<ThemePart>("rId7");
            GenerateThemePart1Content(themePart1);
//Replacement
//
            WorksheetPart worksheetPart2 = workbookPart1.AddNewPart<WorksheetPart>("rId101");
            GenerateWorksheetPartReplacement(worksheetPart2);

            VmlDrawingPart vmlDrawingPart3 = worksheetPart2.AddNewPart<VmlDrawingPart>("rId2");
            vmlDrawingPart3.AddPart(imagePart1, "rId2");

            vmlDrawingPart3.AddPart(imagePart2, "rId1");
            GenerateVmlDrawingPart10Content(vmlDrawingPart3);

            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart2 = worksheetPart2.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
            GenerateSpreadsheetPrinterSettingsPart3Content(spreadsheetPrinterSettingsPart2);
//Page Setup
//
            WorksheetPart worksheetPart3 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerateWorksheetPartSetupPage(worksheetPart3);

            VmlDrawingPart vmlDrawingPart4 = worksheetPart3.AddNewPart<VmlDrawingPart>("rId2");
            vmlDrawingPart4.AddPart(imagePart1, "rId2");
            vmlDrawingPart4.AddPart(imagePart2, "rId1");

            GenerateVmlDrawingPart10Content(vmlDrawingPart4);

            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart3 = worksheetPart3.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
            GenerateSpreadsheetPrinterSettingsPart3Content(spreadsheetPrinterSettingsPart3);
//Survey Detail
//
            WorksheetPart worksheetPart4 = workbookPart1.AddNewPart<WorksheetPart>("rId6");
            GenerateWorksheetPartSurveyDetail(worksheetPart4);

            VmlDrawingPart vmlDrawingPart5 = worksheetPart4.AddNewPart<VmlDrawingPart>("rId2");
           
            vmlDrawingPart5.AddPart(imagePart1, "rId2");
            vmlDrawingPart5.AddPart(imagePart2, "rId1");
            GenerateVmlDrawingPart10Content(vmlDrawingPart5);

            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart4 = worksheetPart4.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
            GenerateSpreadsheetPrinterSettingsPart3Content(spreadsheetPrinterSettingsPart4);
//Survey Summary
//
            WorksheetPart worksheetPart5 = workbookPart1.AddNewPart<WorksheetPart>("rId5");
            GenerateWorksheetPartSurveySummary(worksheetPart5);

            VmlDrawingPart vmlDrawingPart6 = worksheetPart5.AddNewPart<VmlDrawingPart>("rId2");

            vmlDrawingPart6.AddPart(imagePart1, "rId2");
            vmlDrawingPart6.AddPart(imagePart2, "rId1");

            GenerateVmlDrawingPart10Content(vmlDrawingPart6);


            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart5 = worksheetPart5.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
            GenerateSpreadsheetPrinterSettingsPart3Content(spreadsheetPrinterSettingsPart5);

            CalculationChainPart calculationChainPart1 = workbookPart1.AddNewPart<CalculationChainPart>("rId10");
            GenerateCalculationChainPart1Content(calculationChainPart1);
//Easylink
//
            WorksheetPart worksheetPart6 = workbookPart1.AddNewPart<WorksheetPart>("rId4");
            GenerateWorksheetPartEasylink(worksheetPart6);

            VmlDrawingPart vmlDrawingPart7 = worksheetPart6.AddNewPart<VmlDrawingPart>("rId2");

            vmlDrawingPart7.AddPart(imagePart1, "rId2");
            vmlDrawingPart7.AddPart(imagePart2, "rId1"); 
            GenerateVmlDrawingPart10Content(vmlDrawingPart7);

            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart6 = worksheetPart6.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
            GenerateSpreadsheetPrinterSettingsPart3Content(spreadsheetPrinterSettingsPart6);

            SharedStringTablePart sharedStringTablePart1 = workbookPart1.AddNewPart<SharedStringTablePart>("rId9");
            GenerateSharedStringTablePart1Content(sharedStringTablePart1);
//Service History
//
            WorksheetPart worksheetPart113 = workbookPart1.AddNewPart<WorksheetPart>("rId19");
            GenerateWorksheetPartServiceHistory(worksheetPart113);

            VmlDrawingPart vmlDrawingPart14 = worksheetPart113.AddNewPart<VmlDrawingPart>("rId2");

            vmlDrawingPart14.AddPart(imagePart1, "rId2");
            vmlDrawingPart14.AddPart(imagePart2, "rId1");

            GenerateVmlDrawingPart10Content(vmlDrawingPart14);

//            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart113 = worksheetPart113.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
//            GenerateSpreadsheetPrinterSettingsPart3Content(spreadsheetPrinterSettingsPart113);
//Quarterly History
//
            WorksheetPart worksheetPart114 = workbookPart1.AddNewPart<WorksheetPart>("rId20");
            GenerateWorksheetPartQuarterlyHistory(worksheetPart114);

            VmlDrawingPart vmlDrawingPart15 = worksheetPart114.AddNewPart<VmlDrawingPart>("rId2");

            vmlDrawingPart15.AddPart(imagePart1, "rId2");
            vmlDrawingPart15.AddPart(imagePart2, "rId1");

            GenerateVmlDrawingPart10Content(vmlDrawingPart15);

            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart13 = worksheetPart114.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
            GenerateSpreadsheetPrinterSettingsPart3Content(spreadsheetPrinterSettingsPart13);

//Rollover History
 //
            WorksheetPart worksheetPart15 = workbookPart1.AddNewPart<WorksheetPart>("rId21");
            GenerateWorksheetPartRolloverHistory(worksheetPart15);

            VmlDrawingPart vmlDrawingPart16 = worksheetPart15.AddNewPart<VmlDrawingPart>("rId2");

            vmlDrawingPart16.AddPart(imagePart1, "rId2");
            vmlDrawingPart16.AddPart(imagePart2, "rId1");

            GenerateVmlDrawingPart10Content(vmlDrawingPart16);

            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart16 = worksheetPart15.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
            GenerateSpreadsheetPrinterSettingsPart3Content(spreadsheetPrinterSettingsPart16);
//ModelMatrix
 //

            
            WorksheetPart worksheetPart18 = workbookPart1.AddNewPart<WorksheetPart>("rId22");
            GenerateWorksheetPartModelMatrix(worksheetPart18);

            VmlDrawingPart vmlDrawingPart17 = worksheetPart18.AddNewPart<VmlDrawingPart>("rId2");
            vmlDrawingPart17.AddPart(imagePart1, "rId2");
            vmlDrawingPart17.AddPart(imagePart2, "rId1");
            GenerateVmlDrawingPart10Content(vmlDrawingPart17);
            
//Revision Summary
//
            WorksheetPart worksheetPart120 = workbookPart1.AddNewPart<WorksheetPart>("rId23");
            GenerateWorksheetPartRevisionSummary(worksheetPart120);

            VmlDrawingPart vmlDrawingPart18 = worksheetPart120.AddNewPart<VmlDrawingPart>("rId2");


            vmlDrawingPart18.AddPart(imagePart1, "rId2");
            vmlDrawingPart18.AddPart(imagePart2, "rId1");

            GenerateVmlDrawingPart10Content(vmlDrawingPart18);

            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart18 = worksheetPart120.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
            GenerateSpreadsheetPrinterSettingsPart3Content(spreadsheetPrinterSettingsPart18);

//Revision History
//
            WorksheetPart worksheetPart121 = workbookPart1.AddNewPart<WorksheetPart>("rId24");
            GenerateWorksheetPartRevisionHistory(worksheetPart121);

            VmlDrawingPart vmlDrawingPart21 = worksheetPart121.AddNewPart<VmlDrawingPart>("rId2");

            vmlDrawingPart21.AddPart(imagePart1, "rId2");

            vmlDrawingPart21.AddPart(imagePart2, "rId1");
            GenerateVmlDrawingPart10Content(vmlDrawingPart21);

            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart19 = worksheetPart121.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
            GenerateSpreadsheetPrinterSettingsPart3Content(spreadsheetPrinterSettingsPart19);
//Executive Summary
//
            WorksheetPart worksheetPart123 = workbookPart1.AddNewPart<WorksheetPart>("rId2");
            GenerateWorksheetPartExecutiveSummary(worksheetPart123);

     
            VmlDrawingPart vmlDrawingPart23 = worksheetPart123.AddNewPart<VmlDrawingPart>("rId2");
         
            vmlDrawingPart23.AddPart(imagePart1, "rId2");
            vmlDrawingPart23.AddPart(imagePart2, "rId1");
            GenerateVmlDrawingPart10Content(vmlDrawingPart23);
            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart20 = worksheetPart123.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
            GenerateSpreadsheetPrinterSettingsPart3Content(spreadsheetPrinterSettingsPart20);

            SetPackageProperties(document);

            
        }
        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Excel";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)4U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Worksheets";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "6";

            variant2.Append(vTInt321);

            Vt.Variant variant3 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Named Ranges";

            variant3.Append(vTLPSTR2);

            Vt.Variant variant4 = new Vt.Variant();
            Vt.VTInt32 vTInt322 = new Vt.VTInt32();
            vTInt322.Text = "3";

            variant4.Append(vTInt322);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);
            vTVector1.Append(variant3);
            vTVector1.Append(variant4);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)9U };
            Vt.VTLPSTR vTLPSTR3 = new Vt.VTLPSTR();
            vTLPSTR3.Text = "SETUP";
            Vt.VTLPSTR vTLPSTR4 = new Vt.VTLPSTR();
            vTLPSTR4.Text = "Replacements";
            Vt.VTLPSTR vTLPSTR5 = new Vt.VTLPSTR();
            vTLPSTR5.Text = "Cost Avoidance";
            Vt.VTLPSTR vTLPSTR6 = new Vt.VTLPSTR();
            vTLPSTR6.Text = "Easylink";
            Vt.VTLPSTR vTLPSTR7 = new Vt.VTLPSTR();
            vTLPSTR7.Text = "Surveys";
            Vt.VTLPSTR vTLPSTR8 = new Vt.VTLPSTR();
            vTLPSTR8.Text = "New Survey";
            Vt.VTLPSTR vTLPSTR9 = new Vt.VTLPSTR();
            vTLPSTR9.Text = "Easylink!Print_Area";
            Vt.VTLPSTR vTLPSTR10 = new Vt.VTLPSTR();
            vTLPSTR10.Text = "Easylink!Print_Titles";
            Vt.VTLPSTR vTLPSTR11 = new Vt.VTLPSTR();
            vTLPSTR11.Text = "Replacements!Print_Titles";

            vTVector2.Append(vTLPSTR3);
            vTVector2.Append(vTLPSTR4);
            vTVector2.Append(vTLPSTR5);
            vTVector2.Append(vTLPSTR6);
            vTVector2.Append(vTLPSTR7);
            vTVector2.Append(vTLPSTR8);
            vTVector2.Append(vTLPSTR9);
            vTVector2.Append(vTLPSTR10);
            vTVector2.Append(vTLPSTR11);

            titlesOfParts1.Append(vTVector2);
            Ap.Company company1 = new Ap.Company();
            company1.Text = "Hewlett-Packard Company";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "14.0300";

            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }
        // Build Worksheet
        // Generates content of workbookPart1.
        private void GenerateWorkbookPart1Content(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook();
            workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            FileVersion fileVersion1 = new FileVersion() { ApplicationName = "xl", LastEdited = "5", LowestEdited = "5", BuildVersion = "9303" };
            WorkbookProperties workbookProperties1 = new WorkbookProperties() { CodeName = "ThisWorkbook", HidePivotFieldList = true, DefaultThemeVersion = (UInt32Value)124226U };
          
            BookViews bookViews1 = new BookViews();
            WorkbookView workbookView1 = new WorkbookView() {  XWindow = 480, YWindow = 315, WindowWidth = (UInt32Value)18720U, WindowHeight = (UInt32Value)7530U, TabRatio = (UInt32Value)890U, ActiveTab = (UInt32Value)0U };
           
            bookViews1.Append(workbookView1);

            Sheets sheets1 = new Sheets();
            Sheet sheet1 = new Sheet() { Name = "SETUP", SheetId = (UInt32Value)12U, Id = "rId1",  };
            Sheet sheet2 = new Sheet() { Name = "Executive Summary", SheetId = (UInt32Value)22U, Id = "rId2" };
            Sheet sheet3 = new Sheet() { Name = "Model Matrix", SheetId = (UInt32Value)18U, Id = "rId22" };
            Sheet sheet4 = new Sheet() { Name = "Revision Summary", SheetId = (UInt32Value)20U, Id = "rId23" };
            Sheet sheet5 = new Sheet() { Name = "Revision History", SheetId = (UInt32Value)21U, Id = "rId24" };
            Sheet sheet6 = new Sheet() { Name = "Quarterly History", SheetId = (UInt32Value)16U, Id = "rId20" };
            Sheet sheet7 = new Sheet() { Name = "Rollover History", SheetId = (UInt32Value)17U, Id = "rId21" };
            Sheet sheet8 = new Sheet() { Name = "Volume Trend", SheetId = (UInt32Value)2U, Id = "rId18" };
            Sheet sheet9 = new Sheet() { Name = "Service History", SheetId = (UInt32Value)15U, Id = "rId19" };
            Sheet sheet10 = new Sheet() { Name = "Replacements", SheetId = (UInt32Value)5U, Id = "rId101" };
            Sheet sheet11 = new Sheet() { Name = "Cost Avoidance", SheetId = (UInt32Value)7U, Id = "rId11" };
            Sheet sheet12 = new Sheet() { Name = "Easylink", SheetId = (UInt32Value)9U, Id = "rId4" };
            Sheet sheet13 = new Sheet() { Name = "Surveys", SheetId = (UInt32Value)10U, Id = "rId5" };
            Sheet sheet14 = new Sheet() { Name = "New Survey", SheetId = (UInt32Value)14U, Id = "rId6" };



            sheets1.Append(sheet1);     
            sheets1.Append(sheet2);
            sheets1.Append(sheet3);
            sheets1.Append(sheet4);
            sheets1.Append(sheet5);
            sheets1.Append(sheet6);
            sheets1.Append(sheet7);
            sheets1.Append(sheet8);
            sheets1.Append(sheet9);
            sheets1.Append(sheet10);
            sheets1.Append(sheet11);
            sheets1.Append(sheet12);
            sheets1.Append(sheet13);
            sheets1.Append(sheet14);

            DefinedNames definedNames1 = new DefinedNames();
            DefinedName definedName15 = new DefinedName() { Name = "_xlnm.Print_Area", LocalSheetId = (UInt32Value)11U };
            definedName15.Text = "\'Rollover History\'!$A$1:$F$32";
           
            DefinedName definedName3 = new DefinedName() { Name = "_xlnm.Print_Titles", LocalSheetId = (UInt32Value)11U };
            definedName3.Text = "Replacements!$7:$7";

            DefinedName definedName8 = new DefinedName() { Name = "_xlnm.Print_Titles", LocalSheetId = (UInt32Value)7U };
            definedName8.Text = "\'Volume Trend\'!$5:$5";
            definedNames1.Append(definedName8);
           
           DefinedName definedName13 = new DefinedName() { Name = "_xlnm.Print_Area", LocalSheetId = (UInt32Value)7U };
           definedName13.Text = "\'Volume Trend\'!$A$1:$X$230";
           definedNames1.Append(definedName13);

            DefinedName definedName9 = new DefinedName() { Name = "_xlnm.Print_Titles", LocalSheetId = (UInt32Value)8U };
            definedName9.Text = "\'Service History\'!$11:$11";
            definedNames1.Append(definedName9);

            //DefinedName definedName10 = new DefinedName() { Name = "_xlnm.Print_Titles",  LocalSheetId = (UInt32Value)2U };
            //definedName10.Text = "\'Model Matrix\'!$5:$5";
            //DefinedNames definedNames1 = new DefinedNames();
            DefinedName definedName1 = new DefinedName() { Name = "_xlnm.Print_Titles", LocalSheetId = (UInt32Value)2U };
            definedName1.Text = "\'Model Matrix\'!$5:$5";
            DefinedName definedName2 = new DefinedName() { Name = "_xlnm.Print_Titles" };
            definedName2.Text = "\'Model Matrix\'!$5:$5";
            DefinedName definedName11 = new DefinedName() { Name = "_xlnm.Print_Titles", LocalSheetId = (UInt32Value)11U };
            definedName11.Text = "Easylink!$8:$8";
            DefinedName definedName12 = new DefinedName() { Name = "_xlnm.Print_Titles"};
            definedName12.Text = "Easylink!$8:$8";

            definedNames1.Append(definedName1);
            definedNames1.Append(definedName2);
            definedNames1.Append(definedName3);
          
            definedNames1.Append(definedName11);
            definedNames1.Append(definedName12);

            definedNames1.Append(definedName15);
            CalculationProperties calculationProperties1 = new CalculationProperties() { CalculationId = (UInt32Value)145621U };
            FileRecoveryProperties fileRecoveryProperties1 = new FileRecoveryProperties() { RepairLoad = true };



           // definedNames1.Append(definedName1);
           // definedNames1.Append(definedName2);
          
            //definedNames1.Append(definedName8);
            //definedNames1.Append(definedName9);
            //definedNames1.Append(definedName10);
            //CalculationProperties calculationProperties1 = new CalculationProperties() { CalculationId = (UInt32Value)145621U };

            workbook1.Append(fileVersion1);
            workbook1.Append(workbookProperties1);
            workbook1.Append(bookViews1);
            workbook1.Append(sheets1);
            workbook1.Append(definedNames1);
            workbook1.Append(calculationProperties1);
            workbook1.Append(fileRecoveryProperties1);

            workbookPart1.Workbook = workbook1;
        }
        // Build Style Sheets
        //
        private void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            NumberingFormats numberingFormats1 = new NumberingFormats() { Count = (UInt32Value)9U };
            NumberingFormat numberingFormat1 = new NumberingFormat() { NumberFormatId = (UInt32Value)164U, FormatCode = "[$-409]mmmm\\ d\\,\\ yyyy;@" };
            NumberingFormat numberingFormat2 = new NumberingFormat() { NumberFormatId = (UInt32Value)165U, FormatCode = "$#,##0;[Red]($#,##0)" };
            NumberingFormat numberingFormat3 = new NumberingFormat() { NumberFormatId = (UInt32Value)166U, FormatCode = "$#,##0;[Red]($#,##0)" };
            NumberingFormat numberingFormat4 = new NumberingFormat() { NumberFormatId = (UInt32Value)167U,  FormatCode = "0%;[Red](0%)" };
            NumberingFormat numberingFormat5 = new NumberingFormat() { NumberFormatId = (UInt32Value)168U, FormatCode = "#,##0;[Red](#,##0)" };
            NumberingFormat numberingFormat6 = new NumberingFormat() { NumberFormatId = (UInt32Value)169U, FormatCode = "0%;[Red](0%)"};
            NumberingFormat numberingFormat7 = new NumberingFormat() { NumberFormatId = (UInt32Value)170U, FormatCode = "###\"-\"###\"-\"####" };
            NumberingFormat numberingFormat8 = new NumberingFormat() { NumberFormatId = (UInt32Value)171U, FormatCode = "#,##0;[Red](#,##0)" };
            NumberingFormat numberingFormat9 = new NumberingFormat() { NumberFormatId = (UInt32Value)172U, FormatCode = "$0.00000;[Red]($0.00000)" };
            NumberingFormat numberingFormat10 = new NumberingFormat() { NumberFormatId = (UInt32Value)173U, FormatCode = "\"$\" #,##0.00;[Red](\"$\" #,##0.00)" };
            NumberingFormat numberingFormat11 = new NumberingFormat() { NumberFormatId = (UInt32Value)174U, FormatCode = "##0.00%;[Red](##0.00%)" };
            NumberingFormat numberingFormat12 = new NumberingFormat() { NumberFormatId = (UInt32Value)175U, FormatCode = "##0.0%;[Red](##0.0%)" };
            numberingFormats1.Append(numberingFormat1);
            numberingFormats1.Append(numberingFormat2);
            numberingFormats1.Append(numberingFormat3);
            numberingFormats1.Append(numberingFormat4);
            numberingFormats1.Append(numberingFormat5);
            numberingFormats1.Append(numberingFormat6);
            numberingFormats1.Append(numberingFormat7);
            numberingFormats1.Append(numberingFormat8);
            numberingFormats1.Append(numberingFormat9);
            numberingFormats1.Append(numberingFormat10);
            numberingFormats1.Append(numberingFormat11);
            numberingFormats1.Append(numberingFormat12);

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)26U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            Font font2 = new Font();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            Color color2 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName2 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };

            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);

            Font font3 = new Font();
            FontSize fontSize3 = new FontSize() { Val = 8D };
            Color color3 = new Color() { Indexed = (UInt32Value)8U };
            FontName fontName3 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 2 };
          //  FontScheme fontScheme2a = new FontScheme() { Val = FontSchemeValues.Minor };
            font3.Append(fontSize3);
            font3.Append(color3);
            font3.Append(fontName3);
            font3.Append(fontFamilyNumbering3);
        //    font3.Append(fontScheme2a);
            Font font4 = new Font();
            FontSize fontSize4 = new FontSize() { Val = 8D };
            Color color4 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName4 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering4 = new FontFamilyNumbering() { Val = 2 };
         //   FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font4.Append(fontSize4);
            font4.Append(color4);
            font4.Append(fontName4);
            font4.Append(fontFamilyNumbering4);
        //    font4.Append(fontScheme2);

            Font font5 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize5 = new FontSize() { Val = 8D };
            Color color5 = new Color() { Indexed = (UInt32Value)8U };
            FontName fontName5 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering5 = new FontFamilyNumbering() { Val = 2 };

            font5.Append(bold1);
            font5.Append(fontSize5);
            font5.Append(color5);
            font5.Append(fontName5);
            font5.Append(fontFamilyNumbering5);

            Font font6 = new Font();
            Bold bold2 = new Bold();
            Underline underline1 = new Underline();
            FontSize fontSize6 = new FontSize() { Val = 8D };
            Color color6 = new Color() { Indexed = (UInt32Value)8U };
            FontName fontName6 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering6 = new FontFamilyNumbering() { Val = 2 };

            font6.Append(bold2);
            font6.Append(underline1);
            font6.Append(fontSize6);
            font6.Append(color6);
            font6.Append(fontName6);
            font6.Append(fontFamilyNumbering6);

            Font font7 = new Font();
            Bold bold3 = new Bold();
            FontSize fontSize7 = new FontSize() { Val = 12D };
            Color color7 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName7 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering7 = new FontFamilyNumbering() { Val = 2 };

            font7.Append(bold3);
            font7.Append(fontSize7);
            font7.Append(color7);
            font7.Append(fontName7);
            font7.Append(fontFamilyNumbering7);

            Font font8 = new Font();
            Bold bold4 = new Bold();
            Italic italic1 = new Italic();
            FontSize fontSize8 = new FontSize() { Val = 11D };
            Color color8 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName8 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering8 = new FontFamilyNumbering() { Val = 2 };

            font8.Append(bold4);
            font8.Append(italic1);
            font8.Append(fontSize8);
            font8.Append(color8);
            font8.Append(fontName8);
            font8.Append(fontFamilyNumbering8);

            Font font9 = new Font();
            FontSize fontSize9 = new FontSize() { Val = 11D };
            Color color9 = new Color() { Theme = (UInt32Value)0U };
            FontName fontName9 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering9 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme3 = new FontScheme() { Val = FontSchemeValues.Minor };

            font9.Append(fontSize9);
            font9.Append(color9);
            font9.Append(fontName9);
            font9.Append(fontFamilyNumbering9);
            font9.Append(fontScheme3);

            Font font10 = new Font();
            Bold bold5 = new Bold();
            FontSize fontSize10 = new FontSize() { Val = 11D };
            Color color10 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName10 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering10 = new FontFamilyNumbering() { Val = 2 };

            font10.Append(bold5);
            font10.Append(fontSize10);
            font10.Append(color10);
            font10.Append(fontName10);
            font10.Append(fontFamilyNumbering10);

            Font font11 = new Font();
            Bold bold6 = new Bold();
            FontSize fontSize11 = new FontSize() { Val = 14D };
            Color color11 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName11 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering11 = new FontFamilyNumbering() { Val = 2 };

            font11.Append(bold6);
            font11.Append(fontSize11);
            font11.Append(color11);
            font11.Append(fontName11);
            font11.Append(fontFamilyNumbering11);

            Font font12 = new Font();
            Bold bold7 = new Bold();
            FontSize fontSize12 = new FontSize() { Val = 12D };
            Color color12 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName12 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering12 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme4 = new FontScheme() { Val = FontSchemeValues.Minor };

            font12.Append(bold7);
            font12.Append(fontSize12);
            font12.Append(color12);
            font12.Append(fontName12);
            font12.Append(fontFamilyNumbering12);
            font12.Append(fontScheme4);

            Font font13 = new Font();
            FontSize fontSize13 = new FontSize() { Val = 10D };
            Color color13 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName13 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering13 = new FontFamilyNumbering() { Val = 2 };

            font13.Append(fontSize13);
            font13.Append(color13);
            font13.Append(fontName13);
            font13.Append(fontFamilyNumbering13);

            Font font14 = new Font();
            Bold bold8 = new Bold();
            FontSize fontSize14 = new FontSize() { Val = 9D };
            FontName fontName14 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering14 = new FontFamilyNumbering() { Val = 2 };

            font14.Append(bold8);
            font14.Append(fontSize14);
            font14.Append(fontName14);
            font14.Append(fontFamilyNumbering14);

            Font font15 = new Font();
            Bold bold9 = new Bold();
            Italic italic2 = new Italic();
            FontSize fontSize15 = new FontSize() { Val = 10D };
            Color color14 = new Color() { Theme = (UInt32Value)0U, Tint = -0.34998626667073579D };
            FontName fontName15 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering15 = new FontFamilyNumbering() { Val = 2 };

            font15.Append(bold9);
            font15.Append(italic2);
            font15.Append(fontSize15);
            font15.Append(color14);
            font15.Append(fontName15);
            font15.Append(fontFamilyNumbering15);

            Font font16 = new Font();
            Bold bold10 = new Bold();
            FontSize fontSize16 = new FontSize() { Val = 9D };
            Color color15 = new Color() { Indexed = (UInt32Value)81U };
            FontName fontName16 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering16 = new FontFamilyNumbering() { Val = 2 };

            font16.Append(bold10);
            font16.Append(fontSize16);
            font16.Append(color15);
            font16.Append(fontName16);
            font16.Append(fontFamilyNumbering16);

            Font font17 = new Font();
            FontSize fontSize17 = new FontSize() { Val = 11D };
            Color color16 = new Color() { Rgb = "FFFF0000" };
            FontName fontName17 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering17 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme5 = new FontScheme() { Val = FontSchemeValues.Minor };

            font17.Append(fontSize17);
            font17.Append(color16);
            font17.Append(fontName17);
            font17.Append(fontFamilyNumbering17);
            font17.Append(fontScheme5);

            Font font18 = new Font();
            Bold bold11 = new Bold();
            FontSize fontSize18 = new FontSize() { Val = 10D };
            FontName fontName18 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering18 = new FontFamilyNumbering() { Val = 2 };

            font18.Append(bold11);
            font18.Append(fontSize18);
            font18.Append(fontName18);
            font18.Append(fontFamilyNumbering18);

            Font font19 = new Font();
            Bold bold12 = new Bold();
            FontSize fontSize19 = new FontSize() { Val = 10D };
            Color color17 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName19 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering19 = new FontFamilyNumbering() { Val = 2 };

            font19.Append(bold12);
            font19.Append(fontSize19);
            font19.Append(color17);
            font19.Append(fontName19);
            font19.Append(fontFamilyNumbering19);

            Font font20 = new Font();
            FontSize fontSize20 = new FontSize() { Val = 8D };
            Color color18 = new Color() { Rgb = "FFFF0000" };
            FontName fontName20 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering20 = new FontFamilyNumbering() { Val = 2 };

            font20.Append(fontSize20);
            font20.Append(color18);
            font20.Append(fontName20);
            font20.Append(fontFamilyNumbering20);

            Font font21 = new Font();
            Italic italic3 = new Italic();
            FontSize fontSize21 = new FontSize() { Val = 8D };
            Color color19 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName21 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering21 = new FontFamilyNumbering() { Val = 2 };

            font21.Append(italic3);
            font21.Append(fontSize21);
            font21.Append(color19);
            font21.Append(fontName21);
            font21.Append(fontFamilyNumbering21);

            Font font22 = new Font();
            FontSize fontSize22 = new FontSize() { Val = 10D };
            Color color20 = new Color() { Theme = (UInt32Value)0U };
            FontName fontName22 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering22 = new FontFamilyNumbering() { Val = 2 };

            font22.Append(fontSize22);
            font22.Append(color20);
            font22.Append(fontName22);
            font22.Append(fontFamilyNumbering22);

            Font font23 = new Font();
            Bold bold13 = new Bold();
            FontSize fontSize23 = new FontSize() { Val = 9D };
            Color color21 = new Color() { Rgb = "FFFF0000" };
            FontName fontName23 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering23 = new FontFamilyNumbering() { Val = 2 };

            font23.Append(bold13);
            font23.Append(fontSize23);
            font23.Append(color21);
            font23.Append(fontName23);
            font23.Append(fontFamilyNumbering23);

            Font font24 = new Font();
            Bold bold14 = new Bold();
            FontSize fontSize24 = new FontSize() { Val = 7D };
            Color color22 = new Color() { Rgb = "FFC00000" };
            FontName fontName24 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering24 = new FontFamilyNumbering() { Val = 2 };

            font24.Append(bold14);
            font24.Append(fontSize24);
            font24.Append(color22);
            font24.Append(fontName24);
            font24.Append(fontFamilyNumbering24);

            Font font25 = new Font();
            FontSize fontSize25 = new FontSize() { Val = 14D };
      
            FontName fontName25 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering25 = new FontFamilyNumbering() { Val = 2 };

            Font font26 = new Font();
            FontSize fontSize26 = new FontSize() { Val = 6D };
            FontName fontName26 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering26 = new FontFamilyNumbering() { Val = 2 };

            Font font27 = new Font();
            FontSize fontSize27 = new FontSize() { Val = 11D };
            Color color55 = new Color() { Rgb = "FFCC0000" };
            FontName fontName27 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering27 = new FontFamilyNumbering() { Val = 2 };


  
            font27.Append(fontSize27);
            font27.Append(color55);
            font27.Append(fontName27);
            font27.Append(fontFamilyNumbering27);


            Font font28 = new Font();
            FontSize fontSize28 = new FontSize() { Val = 9D };
            FontName fontName28 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering28 = new FontFamilyNumbering() { Val = 2 };

            font28.Append(fontSize28);
            font28.Append(fontName28);
            font28.Append(fontFamilyNumbering28);

      
            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);
            fonts1.Append(font5);
            fonts1.Append(font6);
            fonts1.Append(font7);
            fonts1.Append(font8);
            fonts1.Append(font9);
            fonts1.Append(font10);
            fonts1.Append(font11);
            fonts1.Append(font12);
            fonts1.Append(font13);
            fonts1.Append(font14);
            fonts1.Append(font15);
            fonts1.Append(font16);
            fonts1.Append(font17);
            fonts1.Append(font18);
            fonts1.Append(font19);
            fonts1.Append(font20);
            fonts1.Append(font21);
            fonts1.Append(font22);
            fonts1.Append(font23);
            fonts1.Append(font24);
            fonts1.Append(font25);
            fonts1.Append(font26);
            fonts1.Append(font27);
            fonts1.Append(font28);
            Fills fills1 = new Fills() { Count = (UInt32Value)5U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            Fill fill3 = new Fill();

            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Theme = (UInt32Value)0U };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill3.Append(foregroundColor1);
            patternFill3.Append(backgroundColor1);

            fill3.Append(patternFill3);

            Fill fill4 = new Fill();

            GradientFill gradientFill1 = new GradientFill() { Degree = 90D };

            GradientStop gradientStop1 = new GradientStop() { Position = 0D };
            Color color23 = new Color() { Theme = (UInt32Value)4U, Tint = 0.80001220740379042D };

            gradientStop1.Append(color23);

            GradientStop gradientStop2 = new GradientStop() { Position = 1D };
            Color color24 = new Color() { Theme = (UInt32Value)4U, Tint = 0.40000610370189521D };

            gradientStop2.Append(color24);

            gradientFill1.Append(gradientStop1);
            gradientFill1.Append(gradientStop2);

            fill4.Append(gradientFill1);

            Fill fill5 = new Fill();

            PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor2 = new ForegroundColor() { Theme = (UInt32Value)4U, Tint = 0.79998168889431442D };
            BackgroundColor backgroundColor2 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill4.Append(foregroundColor2);
            patternFill4.Append(backgroundColor2);

            fill5.Append(patternFill4);

            fills1.Append(fill1);
            fills1.Append(fill2);
            fills1.Append(fill3);
            fills1.Append(fill4);
            fills1.Append(fill5);

            Borders borders1 = new Borders() { Count = (UInt32Value)39U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color25 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color25);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color26 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color26);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color27 = new Color() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color27);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color28 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color28);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            Border border3 = new Border();

            LeftBorder leftBorder3 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color29 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder3.Append(color29);
            RightBorder rightBorder3 = new RightBorder();

            TopBorder topBorder3 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color30 = new Color() { Indexed = (UInt32Value)64U };

            topBorder3.Append(color30);

            BottomBorder bottomBorder3 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color31 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder3.Append(color31);
            DiagonalBorder diagonalBorder3 = new DiagonalBorder();

            border3.Append(leftBorder3);
            border3.Append(rightBorder3);
            border3.Append(topBorder3);
            border3.Append(bottomBorder3);
            border3.Append(diagonalBorder3);

            Border border4 = new Border();
            LeftBorder leftBorder4 = new LeftBorder();
            RightBorder rightBorder4 = new RightBorder();

            TopBorder topBorder4 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color32 = new Color() { Indexed = (UInt32Value)64U };

            topBorder4.Append(color32);

            BottomBorder bottomBorder4 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color33 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder4.Append(color33);
            DiagonalBorder diagonalBorder4 = new DiagonalBorder();

            border4.Append(leftBorder4);
            border4.Append(rightBorder4);
            border4.Append(topBorder4);
            border4.Append(bottomBorder4);
            border4.Append(diagonalBorder4);

            Border border5 = new Border();
            LeftBorder leftBorder5 = new LeftBorder();

            RightBorder rightBorder5 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color34 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder5.Append(color34);

            TopBorder topBorder5 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color35 = new Color() { Indexed = (UInt32Value)64U };

            topBorder5.Append(color35);

            BottomBorder bottomBorder5 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color36 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder5.Append(color36);
            DiagonalBorder diagonalBorder5 = new DiagonalBorder();

            border5.Append(leftBorder5);
            border5.Append(rightBorder5);
            border5.Append(topBorder5);
            border5.Append(bottomBorder5);
            border5.Append(diagonalBorder5);

            Border border6 = new Border();

            LeftBorder leftBorder6 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color37 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder6.Append(color37);

            RightBorder rightBorder6 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color38 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder6.Append(color38);
            TopBorder topBorder6 = new TopBorder();

            BottomBorder bottomBorder6 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color39 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder6.Append(color39);
            DiagonalBorder diagonalBorder6 = new DiagonalBorder();

            border6.Append(leftBorder6);
            border6.Append(rightBorder6);
            border6.Append(topBorder6);
            border6.Append(bottomBorder6);
            border6.Append(diagonalBorder6);

            Border border7 = new Border();
            LeftBorder leftBorder7 = new LeftBorder();
            RightBorder rightBorder7 = new RightBorder();
            TopBorder topBorder7 = new TopBorder();

            BottomBorder bottomBorder7 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color40a = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder7.Append(color40a);
            DiagonalBorder diagonalBorder7 = new DiagonalBorder();

            border7.Append(leftBorder7);
            border7.Append(rightBorder7);
            border7.Append(topBorder7);
            border7.Append(bottomBorder7);
            border7.Append(diagonalBorder7);

            Border border8 = new Border();

            LeftBorder leftBorder8 = new LeftBorder() { Style = BorderStyleValues.Medium };
            Color color41 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder8.Append(color41);
            RightBorder rightBorder8 = new RightBorder();
            TopBorder topBorder8 = new TopBorder();
            BottomBorder bottomBorder8 = new BottomBorder();
            DiagonalBorder diagonalBorder8 = new DiagonalBorder();

            border8.Append(leftBorder8);
            border8.Append(rightBorder8);
            border8.Append(topBorder8);
            border8.Append(bottomBorder8);
            border8.Append(diagonalBorder8);

            Border border9 = new Border();
            LeftBorder leftBorder9 = new LeftBorder();

            RightBorder rightBorder9 = new RightBorder() { Style = BorderStyleValues.Medium };
            Color color42 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder9.Append(color42);
            TopBorder topBorder9 = new TopBorder();
            BottomBorder bottomBorder9 = new BottomBorder();
            DiagonalBorder diagonalBorder9 = new DiagonalBorder();

            border9.Append(leftBorder9);
            border9.Append(rightBorder9);
            border9.Append(topBorder9);
            border9.Append(bottomBorder9);
            border9.Append(diagonalBorder9);

            Border border10 = new Border();

            LeftBorder leftBorder10 = new LeftBorder() { Style = BorderStyleValues.Medium };
            Color color43 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder10.Append(color43);
            RightBorder rightBorder10 = new RightBorder();
            TopBorder topBorder10 = new TopBorder();

            BottomBorder bottomBorder10 = new BottomBorder() { Style = BorderStyleValues.Medium };
            Color color44 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder10.Append(color44);
            DiagonalBorder diagonalBorder10 = new DiagonalBorder();

            border10.Append(leftBorder10);
            border10.Append(rightBorder10);
            border10.Append(topBorder10);
            border10.Append(bottomBorder10);
            border10.Append(diagonalBorder10);

            Border border11 = new Border();
            LeftBorder leftBorder11 = new LeftBorder();
            RightBorder rightBorder11 = new RightBorder();
            TopBorder topBorder11 = new TopBorder();

            BottomBorder bottomBorder11 = new BottomBorder() { Style = BorderStyleValues.Medium };
            Color color45a = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder11.Append(color45a);
            DiagonalBorder diagonalBorder11 = new DiagonalBorder();

            border11.Append(leftBorder11);
            border11.Append(rightBorder11);
            border11.Append(topBorder11);
            border11.Append(bottomBorder11);
            border11.Append(diagonalBorder11);

            Border border12 = new Border();
            LeftBorder leftBorder12 = new LeftBorder();

            RightBorder rightBorder12 = new RightBorder() { Style = BorderStyleValues.Medium };
            Color color46 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder12.Append(color46);
            TopBorder topBorder12 = new TopBorder();

            BottomBorder bottomBorder12 = new BottomBorder() { Style = BorderStyleValues.Medium };
            Color color47 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder12.Append(color47);
            DiagonalBorder diagonalBorder12 = new DiagonalBorder();

            border12.Append(leftBorder12);
            border12.Append(rightBorder12);
            border12.Append(topBorder12);
            border12.Append(bottomBorder12);
            border12.Append(diagonalBorder12);

            Border border13 = new Border();

            LeftBorder leftBorder13 = new LeftBorder() { Style = BorderStyleValues.Medium };
            Color color48 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder13.Append(color48);
            RightBorder rightBorder13 = new RightBorder();

            TopBorder topBorder13 = new TopBorder() { Style = BorderStyleValues.Medium };
            Color color49 = new Color() { Indexed = (UInt32Value)64U };

            topBorder13.Append(color49);
            BottomBorder bottomBorder13 = new BottomBorder();
            DiagonalBorder diagonalBorder13 = new DiagonalBorder();

            border13.Append(leftBorder13);
            border13.Append(rightBorder13);
            border13.Append(topBorder13);
            border13.Append(bottomBorder13);
            border13.Append(diagonalBorder13);

            Border border14 = new Border();
            LeftBorder leftBorder14 = new LeftBorder();

            RightBorder rightBorder14 = new RightBorder() { Style = BorderStyleValues.Medium };
            Color color50 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder14.Append(color50);

            TopBorder topBorder14 = new TopBorder() { Style = BorderStyleValues.Medium };
            Color color51 = new Color() { Indexed = (UInt32Value)64U };

            topBorder14.Append(color51);
            BottomBorder bottomBorder14 = new BottomBorder();
            DiagonalBorder diagonalBorder14 = new DiagonalBorder();

            border14.Append(leftBorder14);
            border14.Append(rightBorder14);
            border14.Append(topBorder14);
            border14.Append(bottomBorder14);
            border14.Append(diagonalBorder14);

            Border border15 = new Border();

            LeftBorder leftBorder15 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color52 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder15.Append(color52);
            RightBorder rightBorder15 = new RightBorder();

            TopBorder topBorder15 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color53 = new Color() { Indexed = (UInt32Value)64U };

            topBorder15.Append(color53);
            BottomBorder bottomBorder15 = new BottomBorder();
            DiagonalBorder diagonalBorder15 = new DiagonalBorder();

            border15.Append(leftBorder15);
            border15.Append(rightBorder15);
            border15.Append(topBorder15);
            border15.Append(bottomBorder15);
            border15.Append(diagonalBorder15);

            Border border16 = new Border();
            LeftBorder leftBorder16 = new LeftBorder();

            RightBorder rightBorder16 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color54 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder16.Append(color54);

            TopBorder topBorder16 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color551 = new Color() { Indexed = (UInt32Value)64U };

            topBorder16.Append(color551);
            BottomBorder bottomBorder16 = new BottomBorder();
            DiagonalBorder diagonalBorder16 = new DiagonalBorder();

            border16.Append(leftBorder16);
            border16.Append(rightBorder16);
            border16.Append(topBorder16);
            border16.Append(bottomBorder16);
            border16.Append(diagonalBorder16);

            Border border17 = new Border();

            LeftBorder leftBorder17 = new LeftBorder() { Style = BorderStyleValues.Medium };
            Color color56 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder17.Append(color56);
            RightBorder rightBorder17 = new RightBorder();

            TopBorder topBorder17 = new TopBorder() { Style = BorderStyleValues.Medium };
            Color color57 = new Color() { Indexed = (UInt32Value)64U };

            topBorder17.Append(color57);

            BottomBorder bottomBorder17 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color58 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder17.Append(color58);
            DiagonalBorder diagonalBorder17 = new DiagonalBorder();

            border17.Append(leftBorder17);
            border17.Append(rightBorder17);
            border17.Append(topBorder17);
            border17.Append(bottomBorder17);
            border17.Append(diagonalBorder17);

            Border border18 = new Border();
            LeftBorder leftBorder18 = new LeftBorder();
            RightBorder rightBorder18 = new RightBorder();

            TopBorder topBorder18 = new TopBorder() { Style = BorderStyleValues.Medium };
            Color color59 = new Color() { Indexed = (UInt32Value)64U };

            topBorder18.Append(color59);

            BottomBorder bottomBorder18 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color60 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder18.Append(color60);
            DiagonalBorder diagonalBorder18 = new DiagonalBorder();

            border18.Append(leftBorder18);
            border18.Append(rightBorder18);
            border18.Append(topBorder18);
            border18.Append(bottomBorder18);
            border18.Append(diagonalBorder18);

            Border border19 = new Border();
            LeftBorder leftBorder19 = new LeftBorder();

            RightBorder rightBorder19 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color61 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder19.Append(color61);

            TopBorder topBorder19 = new TopBorder() { Style = BorderStyleValues.Medium };
            Color color62 = new Color() { Indexed = (UInt32Value)64U };

            topBorder19.Append(color62);

            BottomBorder bottomBorder19 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color63 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder19.Append(color63);
            DiagonalBorder diagonalBorder19 = new DiagonalBorder();

            border19.Append(leftBorder19);
            border19.Append(rightBorder19);
            border19.Append(topBorder19);
            border19.Append(bottomBorder19);
            border19.Append(diagonalBorder19);

            Border border20 = new Border();

            LeftBorder leftBorder20 = new LeftBorder() { Style = BorderStyleValues.Medium };
            Color color64 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder20.Append(color64);
            RightBorder rightBorder20 = new RightBorder();

            TopBorder topBorder20 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color65 = new Color() { Indexed = (UInt32Value)64U };

            topBorder20.Append(color65);

            BottomBorder bottomBorder20 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color66 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder20.Append(color66);
            DiagonalBorder diagonalBorder20 = new DiagonalBorder();

            border20.Append(leftBorder20);
            border20.Append(rightBorder20);
            border20.Append(topBorder20);
            border20.Append(bottomBorder20);
            border20.Append(diagonalBorder20);

            Border border21 = new Border();

            LeftBorder leftBorder21 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color67 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder21.Append(color67);

            RightBorder rightBorder21 = new RightBorder() { Style = BorderStyleValues.Medium };
            Color color68 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder21.Append(color68);

            TopBorder topBorder21 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color69 = new Color() { Indexed = (UInt32Value)64U };

            topBorder21.Append(color69);

            BottomBorder bottomBorder21 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color70 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder21.Append(color70);
            DiagonalBorder diagonalBorder21 = new DiagonalBorder();

            border21.Append(leftBorder21);
            border21.Append(rightBorder21);
            border21.Append(topBorder21);
            border21.Append(bottomBorder21);
            border21.Append(diagonalBorder21);

            Border border22 = new Border();
            LeftBorder leftBorder22 = new LeftBorder();
            RightBorder rightBorder22 = new RightBorder();

            TopBorder topBorder22 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color71 = new Color() { Indexed = (UInt32Value)64U };

            topBorder22.Append(color71);

            BottomBorder bottomBorder22 = new BottomBorder() { Style = BorderStyleValues.Medium };
            Color color72 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder22.Append(color72);
            DiagonalBorder diagonalBorder22 = new DiagonalBorder();

            border22.Append(leftBorder22);
            border22.Append(rightBorder22);
            border22.Append(topBorder22);
            border22.Append(bottomBorder22);
            border22.Append(diagonalBorder22);

            Border border23 = new Border();

            LeftBorder leftBorder23 = new LeftBorder() { Style = BorderStyleValues.Medium };
            Color color73 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder23.Append(color73);
            RightBorder rightBorder23 = new RightBorder();
            TopBorder topBorder23 = new TopBorder();

            BottomBorder bottomBorder23 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color74 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder23.Append(color74);
            DiagonalBorder diagonalBorder23 = new DiagonalBorder();

            border23.Append(leftBorder23);
            border23.Append(rightBorder23);
            border23.Append(topBorder23);
            border23.Append(bottomBorder23);
            border23.Append(diagonalBorder23);

            Border border24 = new Border();

            LeftBorder leftBorder24 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color75 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder24.Append(color75);

            RightBorder rightBorder24 = new RightBorder() { Style = BorderStyleValues.Medium };
            Color color76 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder24.Append(color76);
            TopBorder topBorder24 = new TopBorder();

            BottomBorder bottomBorder24 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color77 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder24.Append(color77);
            DiagonalBorder diagonalBorder24 = new DiagonalBorder();

            border24.Append(leftBorder24);
            border24.Append(rightBorder24);
            border24.Append(topBorder24);
            border24.Append(bottomBorder24);
            border24.Append(diagonalBorder24);

            Border border25 = new Border();

            LeftBorder leftBorder25 = new LeftBorder() { Style = BorderStyleValues.Medium };
            Color color78 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder25.Append(color78);
            RightBorder rightBorder25 = new RightBorder();

            TopBorder topBorder25 = new TopBorder() { Style = BorderStyleValues.Medium };
            Color color79 = new Color() { Indexed = (UInt32Value)64U };

            topBorder25.Append(color79);

            BottomBorder bottomBorder25 = new BottomBorder() { Style = BorderStyleValues.Medium };
            Color color80 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder25.Append(color80);
            DiagonalBorder diagonalBorder25 = new DiagonalBorder();

            border25.Append(leftBorder25);
            border25.Append(rightBorder25);
            border25.Append(topBorder25);
            border25.Append(bottomBorder25);
            border25.Append(diagonalBorder25);

            Border border26 = new Border();
            LeftBorder leftBorder26 = new LeftBorder();
            RightBorder rightBorder26 = new RightBorder();

            TopBorder topBorder26 = new TopBorder() { Style = BorderStyleValues.Medium };
            Color color81 = new Color() { Indexed = (UInt32Value)64U };

            topBorder26.Append(color81);

            BottomBorder bottomBorder26 = new BottomBorder() { Style = BorderStyleValues.Medium };
            Color color82 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder26.Append(color82);
            DiagonalBorder diagonalBorder26 = new DiagonalBorder();

            border26.Append(leftBorder26);
            border26.Append(rightBorder26);
            border26.Append(topBorder26);
            border26.Append(bottomBorder26);
            border26.Append(diagonalBorder26);

            Border border27 = new Border();
            LeftBorder leftBorder27 = new LeftBorder();

            RightBorder rightBorder27 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color83 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder27.Append(color83);

            TopBorder topBorder27 = new TopBorder() { Style = BorderStyleValues.Medium };
            Color color84 = new Color() { Indexed = (UInt32Value)64U };

            topBorder27.Append(color84);

            BottomBorder bottomBorder27 = new BottomBorder() { Style = BorderStyleValues.Medium };
            Color color85 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder27.Append(color85);
            DiagonalBorder diagonalBorder27 = new DiagonalBorder();

            border27.Append(leftBorder27);
            border27.Append(rightBorder27);
            border27.Append(topBorder27);
            border27.Append(bottomBorder27);
            border27.Append(diagonalBorder27);

            Border border28 = new Border();

            LeftBorder leftBorder28 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color86 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder28.Append(color86);

            RightBorder rightBorder28 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color87 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder28.Append(color87);

            TopBorder topBorder28 = new TopBorder() { Style = BorderStyleValues.Medium };
            Color color88 = new Color() { Indexed = (UInt32Value)64U };

            topBorder28.Append(color88);

            BottomBorder bottomBorder28 = new BottomBorder() { Style = BorderStyleValues.Medium };
            Color color89 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder28.Append(color89);
            DiagonalBorder diagonalBorder28 = new DiagonalBorder();

            border28.Append(leftBorder28);
            border28.Append(rightBorder28);
            border28.Append(topBorder28);
            border28.Append(bottomBorder28);
            border28.Append(diagonalBorder28);

            Border border29 = new Border();

            LeftBorder leftBorder29 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color90 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder29.Append(color90);

            RightBorder rightBorder29 = new RightBorder() { Style = BorderStyleValues.Medium };
            Color color91 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder29.Append(color91);

            TopBorder topBorder29 = new TopBorder() { Style = BorderStyleValues.Medium };
            Color color92 = new Color() { Indexed = (UInt32Value)64U };

            topBorder29.Append(color92);

            BottomBorder bottomBorder29 = new BottomBorder() { Style = BorderStyleValues.Medium };
            Color color93 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder29.Append(color93);
            DiagonalBorder diagonalBorder29 = new DiagonalBorder();

            border29.Append(leftBorder29);
            border29.Append(rightBorder29);
            border29.Append(topBorder29);
            border29.Append(bottomBorder29);
            border29.Append(diagonalBorder29);

            Border border30 = new Border();

            LeftBorder leftBorder30 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color94 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder30.Append(color94);

            RightBorder rightBorder30 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color95 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder30.Append(color95);
            TopBorder topBorder30 = new TopBorder();
            BottomBorder bottomBorder30 = new BottomBorder();
            DiagonalBorder diagonalBorder30 = new DiagonalBorder();

            border30.Append(leftBorder30);
            border30.Append(rightBorder30);
            border30.Append(topBorder30);
            border30.Append(bottomBorder30);
            border30.Append(diagonalBorder30);

            Border border31 = new Border();

            LeftBorder leftBorder31 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color96 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder31.Append(color96);

            RightBorder rightBorder31 = new RightBorder() { Style = BorderStyleValues.Medium };
            Color color97 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder31.Append(color97);
            TopBorder topBorder31 = new TopBorder();
            BottomBorder bottomBorder31 = new BottomBorder();
            DiagonalBorder diagonalBorder31 = new DiagonalBorder();

            border31.Append(leftBorder31);
            border31.Append(rightBorder31);
            border31.Append(topBorder31);
            border31.Append(bottomBorder31);
            border31.Append(diagonalBorder31);

            Border border32 = new Border();
            LeftBorder leftBorder32 = new LeftBorder();

            RightBorder rightBorder32 = new RightBorder() { Style = BorderStyleValues.Medium };
            Color color98 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder32.Append(color98);

            TopBorder topBorder32 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color99 = new Color() { Indexed = (UInt32Value)64U };

            topBorder32.Append(color99);

            BottomBorder bottomBorder32 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color100 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder32.Append(color100);
            DiagonalBorder diagonalBorder32 = new DiagonalBorder();

            border32.Append(leftBorder32);
            border32.Append(rightBorder32);
            border32.Append(topBorder32);
            border32.Append(bottomBorder32);
            border32.Append(diagonalBorder32);

            Border border33 = new Border();

            LeftBorder leftBorder33 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color101 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder33.Append(color101);
            RightBorder rightBorder33 = new RightBorder();

            TopBorder topBorder33 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color102 = new Color() { Indexed = (UInt32Value)64U };

            topBorder33.Append(color102);

            BottomBorder bottomBorder33 = new BottomBorder() { Style = BorderStyleValues.Medium };
            Color color103 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder33.Append(color103);
            DiagonalBorder diagonalBorder33 = new DiagonalBorder();

            border33.Append(leftBorder33);
            border33.Append(rightBorder33);
            border33.Append(topBorder33);
            border33.Append(bottomBorder33);
            border33.Append(diagonalBorder33);

            Border border34 = new Border();
            LeftBorder leftBorder34 = new LeftBorder();

            RightBorder rightBorder34 = new RightBorder() { Style = BorderStyleValues.Medium };
            Color color104 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder34.Append(color104);

            TopBorder topBorder34 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color105 = new Color() { Indexed = (UInt32Value)64U };

            topBorder34.Append(color105);

            BottomBorder bottomBorder34 = new BottomBorder() { Style = BorderStyleValues.Medium };
            Color color106 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder34.Append(color106);
            DiagonalBorder diagonalBorder34 = new DiagonalBorder();

            border34.Append(leftBorder34);
            border34.Append(rightBorder34);
            border34.Append(topBorder34);
            border34.Append(bottomBorder34);
            border34.Append(diagonalBorder34);

            Border border35 = new Border();
            LeftBorder leftBorder35 = new LeftBorder();
            RightBorder rightBorder35 = new RightBorder();

            TopBorder topBorder35 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color107 = new Color() { Indexed = (UInt32Value)64U };

            topBorder35.Append(color107);
            BottomBorder bottomBorder35 = new BottomBorder();
            DiagonalBorder diagonalBorder35 = new DiagonalBorder();

            border35.Append(leftBorder35);
            border35.Append(rightBorder35);
            border35.Append(topBorder35);
            border35.Append(bottomBorder35);
            border35.Append(diagonalBorder35);

            Border border36 = new Border();

            LeftBorder leftBorder36 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color108 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder36.Append(color108);
            RightBorder rightBorder36 = new RightBorder();
            TopBorder topBorder36 = new TopBorder();
            BottomBorder bottomBorder36 = new BottomBorder();
            DiagonalBorder diagonalBorder36 = new DiagonalBorder();

            border36.Append(leftBorder36);
            border36.Append(rightBorder36);
            border36.Append(topBorder36);
            border36.Append(bottomBorder36);
            border36.Append(diagonalBorder36);

            Border border37 = new Border();
            LeftBorder leftBorder37 = new LeftBorder();

            RightBorder rightBorder37 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color109 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder37.Append(color109);
            TopBorder topBorder37 = new TopBorder();
            BottomBorder bottomBorder37 = new BottomBorder();
            DiagonalBorder diagonalBorder37 = new DiagonalBorder();

            border37.Append(leftBorder37);
            border37.Append(rightBorder37);
            border37.Append(topBorder37);
            border37.Append(bottomBorder37);
            border37.Append(diagonalBorder37);

            Border border38 = new Border();

            LeftBorder leftBorder38 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color110 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder38.Append(color110);
            RightBorder rightBorder38 = new RightBorder();
            TopBorder topBorder38 = new TopBorder();

            BottomBorder bottomBorder38 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color111 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder38.Append(color111);
            DiagonalBorder diagonalBorder38 = new DiagonalBorder();

            border38.Append(leftBorder38);
            border38.Append(rightBorder38);
            border38.Append(topBorder38);
            border38.Append(bottomBorder38);
            border38.Append(diagonalBorder38);

            Border border39 = new Border();
            LeftBorder leftBorder39 = new LeftBorder();

            RightBorder rightBorder39 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color112 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder39.Append(color112);
            TopBorder topBorder39 = new TopBorder();

            BottomBorder bottomBorder39 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color113 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder39.Append(color113);
            DiagonalBorder diagonalBorder39 = new DiagonalBorder();

            border39.Append(leftBorder39);
            border39.Append(rightBorder39);
            border39.Append(topBorder39);
            border39.Append(bottomBorder39);
            border39.Append(diagonalBorder39);

            Border border = new Border();
            LeftBorder leftBorder = new LeftBorder();
            RightBorder rightBorder = new RightBorder();
            TopBorder topBorder = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color114a = new Color() { Indexed = (UInt32Value)64U };
            topBorder.Append(color114a);
            BottomBorder bottomBorder = new BottomBorder();
            DiagonalBorder diagonalBorder = new DiagonalBorder();

            border.Append(leftBorder);
            border.Append(rightBorder);
            border.Append(topBorder);
            border.Append(bottomBorder);
            border.Append(diagonalBorder);

            borders1.Append(border1);
            borders1.Append(border2);
            borders1.Append(border3);
            borders1.Append(border4);
            borders1.Append(border5);
            borders1.Append(border6);
            borders1.Append(border7);
            borders1.Append(border8);
            borders1.Append(border9);
            borders1.Append(border10);
            borders1.Append(border11);
            borders1.Append(border12);
            borders1.Append(border13);
            borders1.Append(border14);
            borders1.Append(border15);
            borders1.Append(border16);
            borders1.Append(border17);
            borders1.Append(border18);
            borders1.Append(border19);
            borders1.Append(border20);
            borders1.Append(border21);
            borders1.Append(border22);
            borders1.Append(border23);
            borders1.Append(border24);
            borders1.Append(border25);
            borders1.Append(border26);
            borders1.Append(border27);
            borders1.Append(border28);
            borders1.Append(border29);
            borders1.Append(border30);
            borders1.Append(border31);
            borders1.Append(border32);
            borders1.Append(border33);
            borders1.Append(border34);
            borders1.Append(border35);
            borders1.Append(border36);
            borders1.Append(border37);
            borders1.Append(border38);
            borders1.Append(border39);
            borders1.Append(border);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)2U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);
            cellStyleFormats1.Append(cellFormat2);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)169U };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true };
            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true };

            CellFormat cellFormat6 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment1 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right };

            cellFormat6.Append(alignment1);

            CellFormat cellFormat7 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)9U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment2 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right };

            cellFormat7.Append(alignment2);

            CellFormat cellFormat8 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment3 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

            cellFormat8.Append(alignment3);

            CellFormat cellFormat9 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment4 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat9.Append(alignment4);

            CellFormat cellFormat10 = new CellFormat() { NumberFormatId = (UInt32Value)38U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment5 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat10.Append(alignment5);

            CellFormat cellFormat11 = new CellFormat() { NumberFormatId = (UInt32Value)38U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment6 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat11.Append(alignment6);

            CellFormat cellFormat12 = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment7 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat12.Append(alignment7);
            CellFormat cellFormat13 = new CellFormat() { NumberFormatId = (UInt32Value)166U, FontId = (UInt32Value)8U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true };

            CellFormat cellFormat14 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)11U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment8 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

            cellFormat14.Append(alignment8);

            CellFormat cellFormat15 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment9 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat15.Append(alignment9);

            CellFormat cellFormat16 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)13U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment10 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right };

            cellFormat16.Append(alignment10);

            CellFormat cellFormat17 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)13U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment11 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

            cellFormat17.Append(alignment11);

            CellFormat cellFormat17a = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)13U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)4U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment11a = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

            cellFormat17a.Append(alignment11a);

            CellFormat cellFormat18 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment12 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat18.Append(alignment12);

            CellFormat cellFormat19 = new CellFormat() { NumberFormatId = (UInt32Value)166U, FontId = (UInt32Value)3U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFill = true, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment13 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat19.Append(alignment13);

            CellFormat cellFormat20 = new CellFormat() { NumberFormatId = (UInt32Value)166U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment14 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat20.Append(alignment14);
            CellFormat cellFormat21 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)16U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true };

            CellFormat cellFormat22 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)7U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment15 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat22.Append(alignment15);
            CellFormat cellFormat23 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)12U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true };

            CellFormat cellFormat24 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment16 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, ShrinkToFit = true };

            cellFormat24.Append(alignment16);

            CellFormat cellFormat25 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)18U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment17 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

            cellFormat25.Append(alignment17);

            CellFormat cellFormat26 = new CellFormat() { NumberFormatId = (UInt32Value)165U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)8U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment18 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat26.Append(alignment18);

            CellFormat cellFormat27 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)8U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment19 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

            cellFormat27.Append(alignment19);

            CellFormat cellFormat28 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)3U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment20 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat28.Append(alignment20);

            CellFormat cellFormat29 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment21 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat29.Append(alignment21);

            CellFormat cellFormat30 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment22 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right };

            cellFormat30.Append(alignment22);
            CellFormat cellFormat31 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true };

            CellFormat cellFormat32 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)18U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment23 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right };

            cellFormat32.Append(alignment23);
            CellFormat cellFormat33 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)22U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true };

            CellFormat cellFormat34 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment24 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat34.Append(alignment24);

            CellFormat cellFormat35 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)3U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment25 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat35.Append(alignment25);

            CellFormat cellFormat36 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)31U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment26 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat36.Append(alignment26);

            CellFormat cellFormat37 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)32U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment27 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat37.Append(alignment27);

            CellFormat cellFormat38 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)21U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment28 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat38.Append(alignment28);

            CellFormat cellFormat39 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)33U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment29 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat39.Append(alignment29);

            CellFormat cellFormat40 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)20U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment30 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat40.Append(alignment30);

            CellFormat cellFormat41 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)13U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment31 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

            cellFormat41.Append(alignment31);
            CellFormat cellFormat42 = new CellFormat() { NumberFormatId = (UInt32Value)166U, FontId = (UInt32Value)14U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true };

            CellFormat cellFormat43 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)14U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment32 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right };

            cellFormat43.Append(alignment32);
            CellFormat cellFormat44 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)21U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true };
            CellFormat cellFormat45 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)8U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true };

            CellFormat cellFormat46 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)7U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment33 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

            cellFormat46.Append(alignment33);

            CellFormat cellFormat47 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)7U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment34 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

            cellFormat47.Append(alignment34);

            CellFormat cellFormat48 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)3U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment35 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, WrapText = true };

            cellFormat48.Append(alignment35);

            CellFormat cellFormat49 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)4U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment36 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, WrapText = true };

            cellFormat49.Append(alignment36);

            CellFormat cellFormat50 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment37 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, WrapText = true };

            cellFormat50.Append(alignment37);

            CellFormat cellFormat51 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)3U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment38 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = true };

            cellFormat51.Append(alignment38);

            CellFormat cellFormat52 = new CellFormat() { NumberFormatId = (UInt32Value)38U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)34U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment39 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat52.Append(alignment39);

            CellFormat cellFormat53 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)34U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment40 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat53.Append(alignment40);

            CellFormat cellFormat54 = new CellFormat() { NumberFormatId = (UInt32Value)38U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment41 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat54.Append(alignment41);

            CellFormat cellFormat55 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment42 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat55.Append(alignment42);

            CellFormat cellFormat56 = new CellFormat() { NumberFormatId = (UInt32Value)38U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment43 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat56.Append(alignment43);

            CellFormat cellFormat57 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment44 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, ShrinkToFit = true };

            cellFormat57.Append(alignment44);

            CellFormat cellFormat58 = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)14U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment45 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat58.Append(alignment45);

            CellFormat cellFormat59 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)34U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment46 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, ShrinkToFit = true };

            cellFormat59.Append(alignment46);
//Set with Text Left and Money Right on the boarderd side
            CellFormat cellFormat60 = new CellFormat() { NumberFormatId = (UInt32Value)166U, FontId = (UInt32Value)4U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)15U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment47 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat60.Append(alignment47);

            CellFormat cellFormat61 = new CellFormat() { NumberFormatId = (UInt32Value)173U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)35U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment48 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat61.Append(alignment48);

            CellFormat cellFormat62 = new CellFormat() { NumberFormatId = (UInt32Value)166U, FontId = (UInt32Value)4U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)36U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment49 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat62.Append(alignment49);

            CellFormat cellFormat63 = new CellFormat() { NumberFormatId = (UInt32Value)173U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)35U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment50 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat63.Append(alignment50);

            CellFormat cellFormat64 = new CellFormat() { NumberFormatId = (UInt32Value)166U, FontId = (UInt32Value)4U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)36U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment51 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };
            
            cellFormat64.Append(alignment51);

            CellFormat cellFormat65 = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)37U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment52 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat65.Append(alignment52);
//End Of Set
            CellFormat cellFormat66 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment53 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat66.Append(alignment53);

            CellFormat cellFormat67 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment54 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, ShrinkToFit = true };

            cellFormat67.Append(alignment54);

            CellFormat cellFormat68 = new CellFormat() { NumberFormatId = (UInt32Value)38U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment55 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat68.Append(alignment55);

            CellFormat cellFormat69 = new CellFormat() { NumberFormatId = (UInt32Value)166U, FontId = (UInt32Value)4U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)38U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment56 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat69.Append(alignment56);

            CellFormat cellFormat70 = new CellFormat() { NumberFormatId = (UInt32Value)166U, FontId = (UInt32Value)5U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)3U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment57 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, WrapText = true };

            cellFormat70.Append(alignment57);

            CellFormat cellFormat71 = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)37U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment58 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat71.Append(alignment58);

            CellFormat cellFormat72 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment59 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat72.Append(alignment59);

            CellFormat cellFormat73 = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment60 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat73.Append(alignment60);

            CellFormat cellFormat74 = new CellFormat() { NumberFormatId = (UInt32Value)166U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment61 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat74.Append(alignment61);

            CellFormat cellFormat75 = new CellFormat() { NumberFormatId = (UInt32Value)166U, FontId = (UInt32Value)4U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)38U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment62 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat75.Append(alignment62);

            CellFormat cellFormat76 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment63 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat76.Append(alignment63);

            CellFormat cellFormat77 = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment64 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat77.Append(alignment64);

            CellFormat cellFormat78 = new CellFormat() { NumberFormatId = (UInt32Value)166U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment65 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat78.Append(alignment65);

            CellFormat cellFormat79 = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)14U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment66 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat79.Append(alignment66);

            CellFormat cellFormat80 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)34U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment67 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat80.Append(alignment67);

            CellFormat cellFormat81 = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)34U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment68 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat81.Append(alignment68);

            CellFormat cellFormat82 = new CellFormat() { NumberFormatId = (UInt32Value)166U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)34U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment69 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat82.Append(alignment69);

            CellFormat cellFormat83 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)4U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)36U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment70 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat83.Append(alignment70);

            CellFormat cellFormat84 = new CellFormat() { NumberFormatId = (UInt32Value)38U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)34U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment71 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat84.Append(alignment71);

            CellFormat cellFormat85 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)34U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment72 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat85.Append(alignment72);

            CellFormat cellFormat86 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)4U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)15U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment73 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat86.Append(alignment73);

            CellFormat cellFormat87 = new CellFormat() { NumberFormatId = (UInt32Value)38U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment74 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat87.Append(alignment74);

            CellFormat cellFormat88 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment75 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat88.Append(alignment75);

            CellFormat cellFormat89 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)4U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)36U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment76 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat89.Append(alignment76);

            CellFormat cellFormat90 = new CellFormat() { NumberFormatId = (UInt32Value)38U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment77 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat90.Append(alignment77);

            CellFormat cellFormat91 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment78 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat91.Append(alignment78);

            CellFormat cellFormat92 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)4U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)38U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment79 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat92.Append(alignment79);

            CellFormat cellFormat93 = new CellFormat() { NumberFormatId = (UInt32Value)164U, FontId = (UInt32Value)3U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)35U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment80 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Top };

            cellFormat93.Append(alignment80);

            CellFormat cellFormat94 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)3U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)36U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment81 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat94.Append(alignment81);

            CellFormat cellFormat95 = new CellFormat() { NumberFormatId = (UInt32Value)164U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)37U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment82 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat95.Append(alignment82);

            CellFormat cellFormat96 = new CellFormat() { NumberFormatId = (UInt32Value)164U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)14U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment83 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat96.Append(alignment83);

            CellFormat cellFormat97 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)4U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)34U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment84 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat97.Append(alignment84);

            CellFormat cellFormat98 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)4U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)15U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment85 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat98.Append(alignment85);

            CellFormat cellFormat99 = new CellFormat() { NumberFormatId = (UInt32Value)171U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)35U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment86 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Top };

            cellFormat99.Append(alignment86);

            CellFormat cellFormat100 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment87 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat100.Append(alignment87);

            CellFormat cellFormat101 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)36U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment88 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Top, WrapText = true };

            cellFormat101.Append(alignment88);

            CellFormat cellFormat102 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)27U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment89 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat102.Append(alignment89);

            CellFormat cellFormat103 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)28U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment90 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat103.Append(alignment90);

            CellFormat cellFormat104 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)12U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment91 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = true };

            cellFormat104.Append(alignment91);

            CellFormat cellFormat105 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)13U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment92 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, WrapText = true };

            cellFormat105.Append(alignment92);

            CellFormat cellFormat106 = new CellFormat() { NumberFormatId = (UInt32Value)164U, FontId = (UInt32Value)17U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment93 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

            cellFormat106.Append(alignment93);

            CellFormat cellFormat107 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)12U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment94 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

            cellFormat107.Append(alignment94);

            CellFormat cellFormat108 = new CellFormat() { NumberFormatId = (UInt32Value)165U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)13U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment95 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat108.Append(alignment95);

            CellFormat cellFormat109 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)35U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment96 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

            cellFormat109.Append(alignment96);

            CellFormat cellFormat110 = new CellFormat() { NumberFormatId = (UInt32Value)165U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)8U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment97 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat110.Append(alignment97);

            CellFormat cellFormat111 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)9U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment98 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

            cellFormat111.Append(alignment98);

            CellFormat cellFormat112 = new CellFormat() { NumberFormatId = (UInt32Value)165U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)11U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment99 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat112.Append(alignment99);

            CellFormat cellFormat113 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment100 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat113.Append(alignment100);

            CellFormat cellFormat114 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)20U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment101 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat114.Append(alignment101);

            CellFormat cellFormat115 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)5U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment102 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat115.Append(alignment102);

            CellFormat cellFormat116 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)23U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment103 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat116.Append(alignment103);

            CellFormat cellFormat117 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)29U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment104 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat117.Append(alignment104);

            CellFormat cellFormat118 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)30U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment105 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat118.Append(alignment105);

            CellFormat cellFormat119 = new CellFormat() { NumberFormatId = (UInt32Value)164U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)34U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment106 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, ShrinkToFit = true };

            cellFormat119.Append(alignment106);

            CellFormat cellFormat120 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment107 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, ShrinkToFit = true };

            cellFormat120.Append(alignment107);

            CellFormat cellFormat121 = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment108 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, ShrinkToFit = true };

            cellFormat121.Append(alignment108);

            CellFormat cellFormat122 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment109 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, ShrinkToFit = true };

            cellFormat122.Append(alignment109);

            CellFormat cellFormat123 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)4U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment110 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat123.Append(alignment110);

            CellFormat cellFormat124 = new CellFormat() { NumberFormatId = (UInt32Value)3U, FontId = (UInt32Value)4U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)38U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment111 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat124.Append(alignment111);
            CellFormat cellFormat125 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)23U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true };

            CellFormat cellFormat126 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)7U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment112 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat126.Append(alignment112);

            CellFormat cellFormat127 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)7U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment113 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat127.Append(alignment113);

            CellFormat cellFormat128 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)6U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment114 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat128.Append(alignment114);

            CellFormat cellFormat129 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)6U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment115 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat129.Append(alignment115);

            CellFormat cellFormat130 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)12U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment116 = new Alignment() { Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat130.Append(alignment116);

            CellFormat cellFormat131 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, ApplyFont = true, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
            Alignment alignment117 = new Alignment() { Vertical = VerticalAlignmentValues.Bottom, WrapText = true };

            cellFormat131.Append(alignment117);
            CellFormat cellFormat132 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };

            CellFormat cellFormat133 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)12U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment118 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, WrapText = true };

            cellFormat133.Append(alignment118);

            CellFormat cellFormat134 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, ApplyFont = true, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true, ApplyBorder = true };
            Alignment alignment119 = new Alignment() { WrapText = true };

            cellFormat134.Append(alignment119);

            CellFormat cellFormat135 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
            Alignment alignment120 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat135.Append(alignment120);

            CellFormat cellFormat136 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)18U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)19U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment121 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat136.Append(alignment121);

            CellFormat cellFormat137 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)3U, FormatId = (UInt32Value)0U, ApplyFill = true, ApplyBorder = true, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment122 = new Alignment() { Vertical = VerticalAlignmentValues.Bottom };

            cellFormat137.Append(alignment122);

            CellFormat cellFormat138 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)4U, FormatId = (UInt32Value)0U, ApplyFill = true, ApplyBorder = true, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment123 = new Alignment() { Vertical = VerticalAlignmentValues.Bottom };

            cellFormat138.Append(alignment123);

            CellFormat cellFormat139 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)18U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)19U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment124 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat139.Append(alignment124);

            CellFormat cellFormat140 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)3U, FormatId = (UInt32Value)0U, ApplyFill = true, ApplyBorder = true, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment125 = new Alignment() { Vertical = VerticalAlignmentValues.Bottom };

            cellFormat140.Append(alignment125);

            CellFormat cellFormat141 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)4U, FormatId = (UInt32Value)0U, ApplyFill = true, ApplyBorder = true, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment126 = new Alignment() { Vertical = VerticalAlignmentValues.Bottom };

            cellFormat141.Append(alignment126);

            CellFormat cellFormat142 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)22U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment127 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat142.Append(alignment127);

            CellFormat cellFormat143 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)22U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment128 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat143.Append(alignment128);

            CellFormat cellFormat144 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)10U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment129 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat144.Append(alignment129);

            CellFormat cellFormat145 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)10U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment130 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat145.Append(alignment130);

            CellFormat cellFormat146 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)9U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment131 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat146.Append(alignment131);

            CellFormat cellFormat147 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)9U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment132 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat147.Append(alignment132);

            CellFormat cellFormat148 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)18U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)24U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment133 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat148.Append(alignment133);
            CellFormat cellFormat149 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)25U, FormatId = (UInt32Value)0U, ApplyFill = true, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            CellFormat cellFormat150 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)26U, FormatId = (UInt32Value)0U, ApplyFill = true, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };

            CellFormat cellFormat151 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)18U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)16U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment134 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat151.Append(alignment134);

            CellFormat cellFormat152 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)17U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment135 = new Alignment() { Vertical = VerticalAlignmentValues.Bottom };

            cellFormat152.Append(alignment135);

            CellFormat cellFormat153 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)18U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment136 = new Alignment() { Vertical = VerticalAlignmentValues.Bottom };

            cellFormat153.Append(alignment136);

            CellFormat cellFormat154 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)19U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)10U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment137 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat154.Append(alignment137);

            CellFormat cellFormat155 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)10U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment138 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat155.Append(alignment138);

            CellFormat cellFormat156 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)6U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment139 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

            cellFormat156.Append(alignment139);

            CellFormat cellFormat157 = new CellFormat() { NumberFormatId = (UInt32Value)167U, FontId = (UInt32Value)24U,   FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = false, ApplyAlignment = true };
            Alignment alignment140 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, ShrinkToFit = true };

            cellFormat157.Append(alignment140);

// Used by Volume Trend         
            CellFormat cellFormat158 = new CellFormat() { NumberFormatId = (UInt32Value)167U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment141 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, ShrinkToFit = true };
            cellFormat158.Append(alignment141);

            CellFormat cellFormat159 = new CellFormat() { NumberFormatId = (UInt32Value)168U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment142 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, ShrinkToFit = true };

            cellFormat159.Append(alignment142);

            CellFormat cellFormat160 = new CellFormat() { NumberFormatId = (UInt32Value)169U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment143 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, ShrinkToFit = true };

            cellFormat160.Append(alignment143);

            CellFormat cellFormat161 = new CellFormat() { NumberFormatId = (UInt32Value)167U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment144 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, ShrinkToFit = true };

            cellFormat161.Append(alignment144);

            CellFormat cellFormat162 = new CellFormat() { NumberFormatId = (UInt32Value)167U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment145 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center};

            cellFormat162.Append(alignment145);

            CellFormat cellFormat163 = new CellFormat() { NumberFormatId = (UInt32Value)168U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment146 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat163.Append(alignment146);

            CellFormat cellFormat164 = new CellFormat() { NumberFormatId = (UInt32Value)167U, FontId = (UInt32Value)13U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)3U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment147 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

            cellFormat164.Append(alignment147);

            CellFormat cellFormat165 = new CellFormat() { NumberFormatId = (UInt32Value)169U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = false, ApplyAlignment = true };
            Alignment alignment148 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center  };
            cellFormat165.Append(alignment148);

            CellFormat cellFormat166 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)13U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = false, ApplyAlignment = true };
            Alignment alignment149 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right };
            cellFormat166.Append(alignment149);

            CellFormat cellFormat167 = new CellFormat() { NumberFormatId = (UInt32Value)169U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment150 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };
            cellFormat167.Append(alignment150);


            CellFormat cellFormat168 = new CellFormat() { NumberFormatId = (UInt32Value)169U, FontId = (UInt32Value)13U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = false, ApplyAlignment = true };
            Alignment alignment151 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };
            cellFormat168.Append(alignment151);


            CellFormat cellFormat169 = new CellFormat() { NumberFormatId = (UInt32Value)166U, FontId = (UInt32Value)13U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = false, ApplyAlignment = true };
            Alignment alignment152 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };
            cellFormat169.Append(alignment152);

 //Set for String Left side Percet Right side
 //BLUE Lines
            CellFormat cellFormat170 = new CellFormat() { NumberFormatId = (UInt32Value)169U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)36U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment153 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };
            cellFormat170.Append(alignment153);


            CellFormat cellFormat171 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)35U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment154 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };
            cellFormat171.Append(alignment154);
//WHITE Lines
            CellFormat cellFormat172 = new CellFormat() { NumberFormatId = (UInt32Value)169U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)36U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = false, ApplyAlignment = true };
            Alignment alignment155 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };
            cellFormat172.Append(alignment155);


            CellFormat cellFormat173 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)35U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = false, ApplyAlignment = true };
            Alignment alignment156 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };
            cellFormat173.Append(alignment156);

//End of Set
//EasyLink Phone Numbers
            CellFormat cellFormat174 = new CellFormat() { NumberFormatId = (UInt32Value)170U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = false, ApplyAlignment = true };
            Alignment alignment157 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };
            cellFormat174.Append(alignment157);

            CellFormat cellFormat175 = new CellFormat() { NumberFormatId = (UInt32Value)170U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment158 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };
            cellFormat175.Append(alignment158);
//Top Line for subtotals
            CellFormat cellFormat176 = new CellFormat() { NumberFormatId = (UInt32Value)165U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)39U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment159 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };
            cellFormat176.Append(alignment159);

            CellFormat cellFormat177 = new CellFormat() { NumberFormatId = (UInt32Value)166U, FontId = (UInt32Value)13U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)39U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = false, ApplyAlignment = true };
            Alignment alignment160 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };
            cellFormat177.Append(alignment160);
//Fields with Red Braces
            CellFormat cellFormat178 = new CellFormat() { NumberFormatId = (UInt32Value)171U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = false, ApplyAlignment = true };
            Alignment alignment161 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };
            cellFormat178.Append(alignment161);

            CellFormat cellFormat179 = new CellFormat() { NumberFormatId = (UInt32Value)171U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment162 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };
            cellFormat179.Append(alignment162);
//Revision History Header Label
            CellFormat cellFormat180 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)3U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment163 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = false };

            cellFormat180.Append(alignment163);

// Used by Volume Trend  Blue
            CellFormat cellFormat181 = new CellFormat() { NumberFormatId = (UInt32Value)167U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment164 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, ShrinkToFit = true };
            cellFormat181.Append(alignment164);

            CellFormat cellFormat182 = new CellFormat() { NumberFormatId = (UInt32Value)168U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment165 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, ShrinkToFit = true };

            cellFormat182.Append(alignment165);

            CellFormat cellFormat183 = new CellFormat() { NumberFormatId = (UInt32Value)167U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment166 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, ShrinkToFit = true };

            cellFormat183.Append(alignment166);

            CellFormat cellFormat184 = new CellFormat() { NumberFormatId = (UInt32Value)169U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment167 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, ShrinkToFit = true };

            cellFormat184.Append(alignment167);

            CellFormat cellFormat185 = new CellFormat() { NumberFormatId = (UInt32Value)171U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment168 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat185.Append(alignment168);

            CellFormat cellFormat186 = new CellFormat() { NumberFormatId = (UInt32Value)168U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment169 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat186.Append(alignment169);


 // Used by Volume Trend Whte background   
            CellFormat cellFormat187 = new CellFormat() { NumberFormatId = (UInt32Value)167U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment170 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, ShrinkToFit = true };
            cellFormat187.Append(alignment170);

            CellFormat cellFormat188 = new CellFormat() { NumberFormatId = (UInt32Value)168U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment171 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, ShrinkToFit = true };

            cellFormat188.Append(alignment171);

            CellFormat cellFormat189 = new CellFormat() { NumberFormatId = (UInt32Value)169U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment172 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, ShrinkToFit = true };

            cellFormat189.Append(alignment172);

            CellFormat cellFormat190 = new CellFormat() { NumberFormatId = (UInt32Value)167U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment173 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, ShrinkToFit = true };

            cellFormat190.Append(alignment173);

            CellFormat cellFormat191 = new CellFormat() { NumberFormatId = (UInt32Value)171U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment174 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, ShrinkToFit = true };

            cellFormat191.Append(alignment174);

            CellFormat cellFormat192 = new CellFormat() { NumberFormatId = (UInt32Value)168U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment175 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, ShrinkToFit = true };

            cellFormat192.Append(alignment175);

//Non Numeric text
            CellFormat cellFormat193 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment176 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, ShrinkToFit = true };

            cellFormat193.Append(alignment176);

            CellFormat cellFormat194= new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment177 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, ShrinkToFit = true };

            cellFormat194.Append(alignment177);

            CellFormat cellFormat195 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)6U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)39U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment178 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = false };
            cellFormat195.Append(alignment178);

            CellFormat cellFormat196 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)13U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment179 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = false };
            cellFormat196.Append(alignment179);
// RED Asterisk volume trend
            CellFormat cellFormat197 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)22U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment180 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = false };
            cellFormat197.Append(alignment180);

            CellFormat cellFormat198 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)22U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment181 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = false };
            cellFormat198.Append(alignment181);
// CCP Format Setup Page
            CellFormat cellFormat199 = new CellFormat() { NumberFormatId = (UInt32Value)172U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)8U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment182 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, WrapText = false };
            cellFormat199.Append(alignment182);

            CellFormat cellFormat200 = new CellFormat() { NumberFormatId = (UInt32Value)172U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)8U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment183 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, WrapText = false };
            cellFormat200.Append(alignment183);
//Report Text in header

            CellFormat cellFormat201 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)27U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment184 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, WrapText = true };
            cellFormat201.Append(alignment184);


  // CCP Format  No Lines
            CellFormat cellFormat202 = new CellFormat() { NumberFormatId = (UInt32Value)172U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment185 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, WrapText = false };
            cellFormat202.Append(alignment185);

            CellFormat cellFormat203 = new CellFormat() { NumberFormatId = (UInt32Value)172U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment186 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, WrapText = false };
            cellFormat203.Append(alignment186);
//Executive summary End Cells with border
            CellFormat cellFormat204 = new CellFormat() { NumberFormatId = (UInt32Value)173U, FontId = (UInt32Value)2U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)36U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment187 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat204.Append(alignment187);
 

            CellFormat cellFormat205 = new CellFormat() { NumberFormatId = (UInt32Value)173U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)36U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment188 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat205.Append(alignment188);

            CellFormat cellFormat206 = new CellFormat() { NumberFormatId = (UInt32Value)173U, FontId = (UInt32Value)13U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment189 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat206.Append(alignment189);

            CellFormat cellFormat207 = new CellFormat() { NumberFormatId = (UInt32Value)171U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)39U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment190 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };
            cellFormat207.Append(alignment190);


            CellFormat cellFormat208 = new CellFormat() { NumberFormatId = (UInt32Value)174U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment191 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat208.Append(alignment191);

            CellFormat cellFormat209 = new CellFormat() { NumberFormatId = (UInt32Value)174U, FontId = (UInt32Value)3U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment192 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center  };

            cellFormat209.Append(alignment192);

            //Percent field
            CellFormat cellFormat210 = new CellFormat() { NumberFormatId = (UInt32Value)175U, FontId = (UInt32Value)14U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = false, ApplyAlignment = true };
            Alignment alignment193 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right };
            cellFormat210.Append(alignment193);

            CellFormat cellFormat211 = new CellFormat() { NumberFormatId = (UInt32Value)173U, FontId = (UInt32Value)13U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)39U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment194 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Top };
            cellFormat211.Append(alignment194);

            CellFormat cellFormat212 = new CellFormat() { NumberFormatId = (UInt32Value)173U, FontId = (UInt32Value)13U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment195 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right, Vertical = VerticalAlignmentValues.Top};
            cellFormat212.Append(alignment195);

            CellFormat cellFormat213 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)39U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment196 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

            cellFormat213.Append(alignment196);

            CellFormat cellFormat214 = new CellFormat() { NumberFormatId = (UInt32Value)173U, FontId = (UInt32Value)4U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)36U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment197 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right };
            cellFormat214.Append(alignment197);

            CellFormat cellFormat215 = new CellFormat() { NumberFormatId = (UInt32Value)173U, FontId = (UInt32Value)4U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)36U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment198 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right };
            cellFormat215.Append(alignment198);

            CellFormat cellFormat216 = new CellFormat() { NumberFormatId = (UInt32Value)173U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)35U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment199 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Bottom };
            cellFormat216.Append(alignment199);

            CellFormat cellFormat217 = new CellFormat() { NumberFormatId = (UInt32Value)164U, FontId = (UInt32Value)4U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)20U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment200 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Bottom };

            cellFormat217.Append(alignment200);

            CellFormat cellFormat218 = new CellFormat() { NumberFormatId = (UInt32Value)173U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)35U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment201 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };
            cellFormat218.Append(alignment201);

            CellFormat cellFormat219 = new CellFormat() { NumberFormatId = (UInt32Value)173U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment202 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };
            cellFormat219.Append(alignment202);

            CellFormat cellFormat220 = new CellFormat() { NumberFormatId = (UInt32Value)173U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyBorder = false, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment203 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };
            cellFormat220.Append(alignment203);

            cellFormats1.Append(cellFormat3); //0
            cellFormats1.Append(cellFormat4); //1
            cellFormats1.Append(cellFormat5); //2
            cellFormats1.Append(cellFormat6); //3
            cellFormats1.Append(cellFormat7); //4
            cellFormats1.Append(cellFormat8); //5
            cellFormats1.Append(cellFormat9);
            cellFormats1.Append(cellFormat10);
            cellFormats1.Append(cellFormat11);
            cellFormats1.Append(cellFormat12);
            cellFormats1.Append(cellFormat13);
            cellFormats1.Append(cellFormat14);
            cellFormats1.Append(cellFormat15);
            cellFormats1.Append(cellFormat16);
            cellFormats1.Append(cellFormat17);
            cellFormats1.Append(cellFormat18);
            cellFormats1.Append(cellFormat19);
            cellFormats1.Append(cellFormat20);
            cellFormats1.Append(cellFormat21);
            cellFormats1.Append(cellFormat22);
            cellFormats1.Append(cellFormat23);
            cellFormats1.Append(cellFormat24);
            cellFormats1.Append(cellFormat25);
            cellFormats1.Append(cellFormat26);
            cellFormats1.Append(cellFormat27);
            cellFormats1.Append(cellFormat28);
            cellFormats1.Append(cellFormat29);
            cellFormats1.Append(cellFormat30);
            cellFormats1.Append(cellFormat31);
            cellFormats1.Append(cellFormat32);
            cellFormats1.Append(cellFormat33);
            cellFormats1.Append(cellFormat34);
            cellFormats1.Append(cellFormat35);
            cellFormats1.Append(cellFormat36);
            cellFormats1.Append(cellFormat37);
            cellFormats1.Append(cellFormat38);
            cellFormats1.Append(cellFormat39);
            cellFormats1.Append(cellFormat40);
            cellFormats1.Append(cellFormat41);
            cellFormats1.Append(cellFormat42);
            cellFormats1.Append(cellFormat43);
            cellFormats1.Append(cellFormat44);
            cellFormats1.Append(cellFormat45);
            cellFormats1.Append(cellFormat46);
            cellFormats1.Append(cellFormat47);
            cellFormats1.Append(cellFormat48);
            cellFormats1.Append(cellFormat49);
            cellFormats1.Append(cellFormat50);
            cellFormats1.Append(cellFormat51);
            cellFormats1.Append(cellFormat52);
            cellFormats1.Append(cellFormat53);
            cellFormats1.Append(cellFormat54);
            cellFormats1.Append(cellFormat55);
            cellFormats1.Append(cellFormat56);
            cellFormats1.Append(cellFormat57);
            cellFormats1.Append(cellFormat58);
            cellFormats1.Append(cellFormat59);
            cellFormats1.Append(cellFormat60);
            cellFormats1.Append(cellFormat61);
            cellFormats1.Append(cellFormat62);
            cellFormats1.Append(cellFormat63);
            cellFormats1.Append(cellFormat64);
            cellFormats1.Append(cellFormat65);
            cellFormats1.Append(cellFormat66);
            cellFormats1.Append(cellFormat67);
            cellFormats1.Append(cellFormat68);
            cellFormats1.Append(cellFormat69);
            cellFormats1.Append(cellFormat70);
            cellFormats1.Append(cellFormat71);
            cellFormats1.Append(cellFormat72);
            cellFormats1.Append(cellFormat73);
            cellFormats1.Append(cellFormat74);
            cellFormats1.Append(cellFormat75);
            cellFormats1.Append(cellFormat76);
            cellFormats1.Append(cellFormat77);
            cellFormats1.Append(cellFormat78);
            cellFormats1.Append(cellFormat79);
            cellFormats1.Append(cellFormat80);
            cellFormats1.Append(cellFormat81);
            cellFormats1.Append(cellFormat82);
            cellFormats1.Append(cellFormat83);
            cellFormats1.Append(cellFormat84);
            cellFormats1.Append(cellFormat85);
            cellFormats1.Append(cellFormat86);
            cellFormats1.Append(cellFormat87);
            cellFormats1.Append(cellFormat88);
            cellFormats1.Append(cellFormat89);
            cellFormats1.Append(cellFormat90);
            cellFormats1.Append(cellFormat91);
            cellFormats1.Append(cellFormat92);
            cellFormats1.Append(cellFormat93);
            cellFormats1.Append(cellFormat94);
            cellFormats1.Append(cellFormat95);
            cellFormats1.Append(cellFormat96);
            cellFormats1.Append(cellFormat97);
            cellFormats1.Append(cellFormat98);
            cellFormats1.Append(cellFormat99);
            cellFormats1.Append(cellFormat100);
            cellFormats1.Append(cellFormat101);
            cellFormats1.Append(cellFormat102);
            cellFormats1.Append(cellFormat103);
            cellFormats1.Append(cellFormat104);
            cellFormats1.Append(cellFormat105);
            cellFormats1.Append(cellFormat106);
            cellFormats1.Append(cellFormat107);
            cellFormats1.Append(cellFormat108);
            cellFormats1.Append(cellFormat109);
            cellFormats1.Append(cellFormat110);
            cellFormats1.Append(cellFormat111);
            cellFormats1.Append(cellFormat112);
            cellFormats1.Append(cellFormat113);
            cellFormats1.Append(cellFormat114);
            cellFormats1.Append(cellFormat115);
            cellFormats1.Append(cellFormat116);
            cellFormats1.Append(cellFormat117);
            cellFormats1.Append(cellFormat118);
            cellFormats1.Append(cellFormat119);
            cellFormats1.Append(cellFormat120);
            cellFormats1.Append(cellFormat121);
            cellFormats1.Append(cellFormat122);
            cellFormats1.Append(cellFormat123);
            cellFormats1.Append(cellFormat124);
            cellFormats1.Append(cellFormat125);
            cellFormats1.Append(cellFormat126);
            cellFormats1.Append(cellFormat127);
            cellFormats1.Append(cellFormat128);
            cellFormats1.Append(cellFormat129);
            cellFormats1.Append(cellFormat130);
            cellFormats1.Append(cellFormat131);
            cellFormats1.Append(cellFormat132);
            cellFormats1.Append(cellFormat133);
            cellFormats1.Append(cellFormat134);
            cellFormats1.Append(cellFormat135);
            cellFormats1.Append(cellFormat136);
            cellFormats1.Append(cellFormat137);
            cellFormats1.Append(cellFormat138);
            cellFormats1.Append(cellFormat139);
            cellFormats1.Append(cellFormat140);
            cellFormats1.Append(cellFormat141);
            cellFormats1.Append(cellFormat142);
            cellFormats1.Append(cellFormat143);
            cellFormats1.Append(cellFormat144);
            cellFormats1.Append(cellFormat145);
            cellFormats1.Append(cellFormat146);
            cellFormats1.Append(cellFormat147);
            cellFormats1.Append(cellFormat148);
            cellFormats1.Append(cellFormat149);
            cellFormats1.Append(cellFormat150);
            cellFormats1.Append(cellFormat151);
            cellFormats1.Append(cellFormat152);
            cellFormats1.Append(cellFormat153);
            cellFormats1.Append(cellFormat154);
            cellFormats1.Append(cellFormat155);
            cellFormats1.Append(cellFormat156);
            cellFormats1.Append(cellFormat157);
            cellFormats1.Append(cellFormat158);
            cellFormats1.Append(cellFormat159);
            cellFormats1.Append(cellFormat160);
            cellFormats1.Append(cellFormat161);
            cellFormats1.Append(cellFormat162);
            cellFormats1.Append(cellFormat163);
            cellFormats1.Append(cellFormat17a);
            cellFormats1.Append(cellFormat164);
            cellFormats1.Append(cellFormat165);
            cellFormats1.Append(cellFormat166);
            cellFormats1.Append(cellFormat167);
            cellFormats1.Append(cellFormat168);
            cellFormats1.Append(cellFormat169);
            cellFormats1.Append(cellFormat170);
            cellFormats1.Append(cellFormat171);
            cellFormats1.Append(cellFormat172);
            cellFormats1.Append(cellFormat173);
            cellFormats1.Append(cellFormat174);
            cellFormats1.Append(cellFormat175);
            cellFormats1.Append(cellFormat176);
            cellFormats1.Append(cellFormat177);
            cellFormats1.Append(cellFormat178);
            cellFormats1.Append(cellFormat179);
            cellFormats1.Append(cellFormat180);

            cellFormats1.Append(cellFormat181);
            cellFormats1.Append(cellFormat182);
            cellFormats1.Append(cellFormat183);
            cellFormats1.Append(cellFormat184);
            cellFormats1.Append(cellFormat185);
            cellFormats1.Append(cellFormat186);
            cellFormats1.Append(cellFormat187);
            cellFormats1.Append(cellFormat188);
            cellFormats1.Append(cellFormat189);
            cellFormats1.Append(cellFormat190);
            cellFormats1.Append(cellFormat191);
            cellFormats1.Append(cellFormat192);
            cellFormats1.Append(cellFormat193);
            cellFormats1.Append(cellFormat194);
            cellFormats1.Append(cellFormat195);
            cellFormats1.Append(cellFormat196);
            cellFormats1.Append(cellFormat197);
            cellFormats1.Append(cellFormat198);
            cellFormats1.Append(cellFormat199);
            cellFormats1.Append(cellFormat200);
            cellFormats1.Append(cellFormat201);
            cellFormats1.Append(cellFormat202);
            cellFormats1.Append(cellFormat203);
            cellFormats1.Append(cellFormat204);
            cellFormats1.Append(cellFormat205);
            cellFormats1.Append(cellFormat206);
            cellFormats1.Append(cellFormat207);
            cellFormats1.Append(cellFormat208);
            cellFormats1.Append(cellFormat209);
            cellFormats1.Append(cellFormat210);
            cellFormats1.Append(cellFormat211);
            cellFormats1.Append(cellFormat212);
            cellFormats1.Append(cellFormat213);
            cellFormats1.Append(cellFormat214);
            cellFormats1.Append(cellFormat215);
            cellFormats1.Append(cellFormat216);
            cellFormats1.Append(cellFormat217);
            cellFormats1.Append(cellFormat218);
            cellFormats1.Append(cellFormat219);
            cellFormats1.Append(cellFormat220);
            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)2U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };
            CellStyle cellStyle2 = new CellStyle() { Name = "Normal 2", FormatId = (UInt32Value)1U };

            cellStyles1.Append(cellStyle1);
            cellStyles1.Append(cellStyle2);

            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)20U };

            DifferentialFormat differentialFormat1 = new DifferentialFormat();

            Font font25a = new Font();
            Bold bold15 = new Bold();
            Italic italic4 = new Italic() { Val = false };
            Color color114 = new Color() { Rgb = "FF009900" };

            font25a.Append(bold15);
            font25a.Append(italic4);
            font25a.Append(color114);

            differentialFormat1.Append(font25a);

            DifferentialFormat differentialFormat2 = new DifferentialFormat();

            Font font26a = new Font();
            Color color115 = new Color() { Rgb = "FFFF9900" };

            font26a.Append(color115);

            differentialFormat2.Append(font26a);

            DifferentialFormat differentialFormat3 = new DifferentialFormat();

            Font font279 = new Font();
            Bold bold16 = new Bold();
            Italic italic5 = new Italic() { Val = false };
            Color color116 = new Color() { Rgb = "FFFF0000" };

            font279.Append(bold16);
            font279.Append(italic5);
            font279.Append(color116);

            differentialFormat3.Append(font279);

            DifferentialFormat differentialFormat4 = new DifferentialFormat();

            Font font28a = new Font();
            Bold bold17 = new Bold();
            Italic italic6 = new Italic() { Val = false };
            Color color117 = new Color() { Rgb = "FFFF0000" };

            font28a.Append(bold17);
            font28a.Append(italic6);
            font28a.Append(color117);

            differentialFormat4.Append(font28a);

            DifferentialFormat differentialFormat5 = new DifferentialFormat();

            Font font29 = new Font();
            Bold bold18 = new Bold();
            Italic italic7 = new Italic() { Val = false };
            Color color118 = new Color() { Rgb = "FFC00000" };

            font29.Append(bold18);
            font29.Append(italic7);
            font29.Append(color118);

            differentialFormat5.Append(font29);

            DifferentialFormat differentialFormat6 = new DifferentialFormat();

            Font font30 = new Font();
            Bold bold19 = new Bold();
            Italic italic8 = new Italic() { Val = false };
            Color color119 = new Color() { Rgb = "FF009900" };

            font30.Append(bold19);
            font30.Append(italic8);
            font30.Append(color119);

            differentialFormat6.Append(font30);

            DifferentialFormat differentialFormat7 = new DifferentialFormat();

            Font font31 = new Font();
            Color color120 = new Color() { Rgb = "FFFF9900" };

            font31.Append(color120);

            differentialFormat7.Append(font31);

            DifferentialFormat differentialFormat8 = new DifferentialFormat();

            Font font32 = new Font();
            Bold bold20 = new Bold();
            Italic italic9 = new Italic() { Val = false };
            Color color121 = new Color() { Rgb = "FFFF0000" };

            font32.Append(bold20);
            font32.Append(italic9);
            font32.Append(color121);

            differentialFormat8.Append(font32);

            DifferentialFormat differentialFormat9 = new DifferentialFormat();

            Font font33 = new Font();
            Bold bold21 = new Bold();
            Italic italic10 = new Italic() { Val = false };
            Color color122 = new Color() { Rgb = "FFFF0000" };

            font33.Append(bold21);
            font33.Append(italic10);
            font33.Append(color122);

            differentialFormat9.Append(font33);

            DifferentialFormat differentialFormat10 = new DifferentialFormat();

            Font font34 = new Font();
            Bold bold22 = new Bold();
            Italic italic11 = new Italic() { Val = false };
            Color color123 = new Color() { Rgb = "FFC00000" };

            font34.Append(bold22);
            font34.Append(italic11);
            font34.Append(color123);

            differentialFormat10.Append(font34);

            DifferentialFormat differentialFormat11 = new DifferentialFormat();

            Font font35 = new Font();
            Bold bold23 = new Bold();
            Italic italic12 = new Italic() { Val = false };
            Color color124 = new Color() { Rgb = "FF009900" };

            font35.Append(bold23);
            font35.Append(italic12);
            font35.Append(color124);

            differentialFormat11.Append(font35);

            DifferentialFormat differentialFormat12 = new DifferentialFormat();

            Font font36 = new Font();
            Color color125 = new Color() { Rgb = "FFFF9900" };

            font36.Append(color125);

            differentialFormat12.Append(font36);

            DifferentialFormat differentialFormat13 = new DifferentialFormat();

            Font font37 = new Font();
            Bold bold24 = new Bold();
            Italic italic13 = new Italic() { Val = false };
            Color color126 = new Color() { Rgb = "FFFF0000" };

            font37.Append(bold24);
            font37.Append(italic13);
            font37.Append(color126);

            differentialFormat13.Append(font37);

            DifferentialFormat differentialFormat14 = new DifferentialFormat();

            Font font38 = new Font();
            Bold bold25 = new Bold();
            Italic italic14 = new Italic() { Val = false };
            Color color127 = new Color() { Rgb = "FFFF0000" };

            font38.Append(bold25);
            font38.Append(italic14);
            font38.Append(color127);

            differentialFormat14.Append(font38);

            DifferentialFormat differentialFormat15 = new DifferentialFormat();

            Font font39 = new Font();
            Bold bold26 = new Bold();
            Italic italic15 = new Italic() { Val = false };
            Color color128 = new Color() { Rgb = "FFC00000" };

            font39.Append(bold26);
            font39.Append(italic15);
            font39.Append(color128);

            differentialFormat15.Append(font39);

            DifferentialFormat differentialFormat16 = new DifferentialFormat();

            Font font40 = new Font();
            Bold bold27 = new Bold();
            Italic italic16 = new Italic() { Val = false };
            Color color129 = new Color() { Rgb = "FF009900" };

            font40.Append(bold27);
            font40.Append(italic16);
            font40.Append(color129);

            differentialFormat16.Append(font40);

            DifferentialFormat differentialFormat17 = new DifferentialFormat();

            Font font41 = new Font();
            Color color130 = new Color() { Rgb = "FFFF9900" };

            font41.Append(color130);

            differentialFormat17.Append(font41);

            DifferentialFormat differentialFormat18 = new DifferentialFormat();

            Font font42 = new Font();
            Bold bold28 = new Bold();
            Italic italic17 = new Italic() { Val = false };
            Color color131 = new Color() { Rgb = "FFFF0000" };

            font42.Append(bold28);
            font42.Append(italic17);
            font42.Append(color131);

            differentialFormat18.Append(font42);

            DifferentialFormat differentialFormat19 = new DifferentialFormat();

            Font font43 = new Font();
            Bold bold29 = new Bold();
            Italic italic18 = new Italic() { Val = false };
            Color color132 = new Color() { Rgb = "FFFF0000" };

            font43.Append(bold29);
            font43.Append(italic18);
            font43.Append(color132);

            differentialFormat19.Append(font43);

            DifferentialFormat differentialFormat20 = new DifferentialFormat();

            Font font44 = new Font();
            Bold bold30 = new Bold();
            Italic italic19 = new Italic() { Val = false };
            Color color133 = new Color() { Rgb = "FFC00000" };

            font44.Append(bold30);
            font44.Append(italic19);
            font44.Append(color133);

            differentialFormat20.Append(font44);

            differentialFormats1.Append(differentialFormat1);
            differentialFormats1.Append(differentialFormat2);
            differentialFormats1.Append(differentialFormat3);
            differentialFormats1.Append(differentialFormat4);
            differentialFormats1.Append(differentialFormat5);
            differentialFormats1.Append(differentialFormat6);
            differentialFormats1.Append(differentialFormat7);
            differentialFormats1.Append(differentialFormat8);
            differentialFormats1.Append(differentialFormat9);
            differentialFormats1.Append(differentialFormat10);
            differentialFormats1.Append(differentialFormat11);
            differentialFormats1.Append(differentialFormat12);
            differentialFormats1.Append(differentialFormat13);
            differentialFormats1.Append(differentialFormat14);
            differentialFormats1.Append(differentialFormat15);
            differentialFormats1.Append(differentialFormat16);
            differentialFormats1.Append(differentialFormat17);
            differentialFormats1.Append(differentialFormat18);
            differentialFormats1.Append(differentialFormat19);
            differentialFormats1.Append(differentialFormat20);
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            //X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            //stylesheetExtension1.Append(slicerStyles1);

            stylesheetExtensionList1.Append(stylesheetExtension1);

            stylesheet1.Append(numberingFormats1);
            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);

            workbookStylesPart1.Stylesheet = stylesheet1;
        }
		// Generates content of worksheetPart1.
     
        // Generates content of worksheetPart7.
        private void GenerateWorksheetPart7Content(WorksheetPart worksheetPart7)
        {
            Worksheet worksheet7 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet7.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet7.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet7.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetProperties sheetProperties7 = new SheetProperties() { CodeName = "Sheet2" };
            SheetDimension sheetDimension7 = new SheetDimension() { Reference = "A1:B25" };

            SheetViews sheetViews7 = new SheetViews();

            SheetView sheetView7 = new SheetView() { TabSelected = true, ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U };
            Selection selection7 = new Selection() { ActiveCell = "B2", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "B2" } };

            sheetView7.Append(selection7);

            sheetViews7.Append(sheetView7);
            SheetFormatProperties sheetFormatProperties7 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            Columns columns7 = new Columns();
            Column column62 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 33D, Style = (UInt32Value)13U, CustomWidth = true };
            Column column63 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 48.85546875D, Style = (UInt32Value)16U, CustomWidth = true };
            Column column64 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)16384U, Width = 9.140625D, Style = (UInt32Value)1U };

            columns7.Append(column62);
            columns7.Append(column63);
            columns7.Append(column64);

            SheetData sheetData7 = new SheetData();

            Row row2748 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, Height = 15.75D, DyDescent = 0.25D };

            Cell cell57030 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)23U, DataType = CellValues.SharedString };
            CellValue cellValue736 = new CellValue();
            cellValue736.Text = "36";

            cell57030.Append(cellValue736);

            row2748.Append(cell57030);

            Row row2749 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.25D };

            Cell cell57031 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)15U, DataType = CellValues.SharedString };
            CellValue cellValue737 = new CellValue();
            cellValue737.Text = "34";

            cell57031.Append(cellValue737);

            Cell cell57032 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)269U, DataType = CellValues.SharedString };
            CellValue cellValue738 = new CellValue();
            cellValue738.Text = "82";

            cell57032.Append(cellValue738);

            row2749.Append(cell57031);
            row2749.Append(cell57032);

            Row row2750 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.25D };

            Cell cell57033 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)15U, DataType = CellValues.SharedString };
            CellValue cellValue739 = new CellValue();
            cellValue739.Text = "39";

            cell57033.Append(cellValue739);

            Cell cell57034 = new Cell() { CellReference = "B3", StyleIndex = (UInt32Value)269U, DataType = CellValues.SharedString };
            CellValue cellValue740 = new CellValue();
            cellValue740.Text = "82";

            cell57034.Append(cellValue740);

            row2750.Append(cell57033);
            row2750.Append(cell57034);

            Row row2751 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.25D };

            Cell cell57035 = new Cell() { CellReference = "A4", StyleIndex = (UInt32Value)15U, DataType = CellValues.SharedString };
            CellValue cellValue741 = new CellValue();
            cellValue741.Text = "35";

            cell57035.Append(cellValue741);

            Cell cell57036 = new Cell() { CellReference = "B4", StyleIndex = (UInt32Value)269U, DataType = CellValues.SharedString };
            CellValue cellValue742 = new CellValue();
            cellValue742.Text = "82";

            cell57036.Append(cellValue742);

            row2751.Append(cell57035);
            row2751.Append(cell57036);

            Row row2752 = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.25D };
            Cell cell57037 = new Cell() { CellReference = "A6", StyleIndex = (UInt32Value)15U };

            row2752.Append(cell57037);

            Row row2753 = new Row() { RowIndex = (UInt32Value)8U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, Height = 15.75D, ThickBot = true, DyDescent = 0.3D };

            Cell cell57038 = new Cell() { CellReference = "A8", StyleIndex = (UInt32Value)58U, DataType = CellValues.SharedString };
            CellValue cellValue743 = new CellValue();
            cellValue743.Text = "88";

            cell57038.Append(cellValue743);

            row2753.Append(cell57038);

            Row row2754 = new Row() { RowIndex = (UInt32Value)9U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, Height = 22.5D, CustomHeight = true, ThickBot = true, DyDescent = 0.3D };

            Cell cell57039 = new Cell() { CellReference = "A9", StyleIndex = (UInt32Value)267U, DataType = CellValues.SharedString };
            CellValue cellValue744 = new CellValue();
            cellValue744.Text = "0";

            cell57039.Append(cellValue744);

            Cell cell57040 = new Cell() { CellReference = "B9", StyleIndex = (UInt32Value)268U, DataType = CellValues.SharedString };
            CellValue cellValue745 = new CellValue();
            cellValue745.Text = "89";

            cell57040.Append(cellValue745);

            row2754.Append(cell57039);
            row2754.Append(cell57040);

            Row row2755 = new Row() { RowIndex = (UInt32Value)10U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.25D };
            Cell cell57041 = new Cell() { CellReference = "A10", StyleIndex = (UInt32Value)270U };
            Cell cell57042 = new Cell() { CellReference = "B10", StyleIndex = (UInt32Value)271U };

            row2755.Append(cell57041);
            row2755.Append(cell57042);

            Row row2756 = new Row() { RowIndex = (UInt32Value)11U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.25D };
            Cell cell57043 = new Cell() { CellReference = "A11", StyleIndex = (UInt32Value)137U };
            Cell cell57044 = new Cell() { CellReference = "B11", StyleIndex = (UInt32Value)59U };

            row2756.Append(cell57043);
            row2756.Append(cell57044);

            Row row2757 = new Row() { RowIndex = (UInt32Value)12U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.25D };
            Cell cell57045 = new Cell() { CellReference = "A12", StyleIndex = (UInt32Value)272U };
            Cell cell57046 = new Cell() { CellReference = "B12", StyleIndex = (UInt32Value)273U };

            row2757.Append(cell57045);
            row2757.Append(cell57046);

            Row row2758 = new Row() { RowIndex = (UInt32Value)13U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.25D };
            Cell cell57047 = new Cell() { CellReference = "A13", StyleIndex = (UInt32Value)137U };
            Cell cell57048 = new Cell() { CellReference = "B13", StyleIndex = (UInt32Value)59U };

            row2758.Append(cell57047);
            row2758.Append(cell57048);

            Row row2759 = new Row() { RowIndex = (UInt32Value)14U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.25D };
            Cell cell57049 = new Cell() { CellReference = "A14", StyleIndex = (UInt32Value)272U };
            Cell cell57050 = new Cell() { CellReference = "B14", StyleIndex = (UInt32Value)273U };

            row2759.Append(cell57049);
            row2759.Append(cell57050);

            Row row2760 = new Row() { RowIndex = (UInt32Value)15U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.25D };
            Cell cell57051 = new Cell() { CellReference = "A15", StyleIndex = (UInt32Value)137U };
            Cell cell57052 = new Cell() { CellReference = "B15", StyleIndex = (UInt32Value)59U };

            row2760.Append(cell57051);
            row2760.Append(cell57052);

            Row row2761 = new Row() { RowIndex = (UInt32Value)16U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.25D };
            Cell cell57053 = new Cell() { CellReference = "A16", StyleIndex = (UInt32Value)272U };
            Cell cell57054 = new Cell() { CellReference = "B16", StyleIndex = (UInt32Value)273U };

            row2761.Append(cell57053);
            row2761.Append(cell57054);

            Row row2762 = new Row() { RowIndex = (UInt32Value)17U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.25D };
            Cell cell57055 = new Cell() { CellReference = "A17", StyleIndex = (UInt32Value)137U };
            Cell cell57056 = new Cell() { CellReference = "B17", StyleIndex = (UInt32Value)59U };

            row2762.Append(cell57055);
            row2762.Append(cell57056);

            Row row2763 = new Row() { RowIndex = (UInt32Value)18U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.25D };
            Cell cell57057 = new Cell() { CellReference = "A18", StyleIndex = (UInt32Value)272U };
            Cell cell57058 = new Cell() { CellReference = "B18", StyleIndex = (UInt32Value)273U };

            row2763.Append(cell57057);
            row2763.Append(cell57058);

            Row row2764 = new Row() { RowIndex = (UInt32Value)19U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.25D };
            Cell cell57059 = new Cell() { CellReference = "A19", StyleIndex = (UInt32Value)137U };
            Cell cell57060 = new Cell() { CellReference = "B19", StyleIndex = (UInt32Value)59U };

            row2764.Append(cell57059);
            row2764.Append(cell57060);

            Row row2765 = new Row() { RowIndex = (UInt32Value)20U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.25D };
            Cell cell57061 = new Cell() { CellReference = "A20", StyleIndex = (UInt32Value)272U };
            Cell cell57062 = new Cell() { CellReference = "B20", StyleIndex = (UInt32Value)273U };

            row2765.Append(cell57061);
            row2765.Append(cell57062);

            Row row2766 = new Row() { RowIndex = (UInt32Value)21U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.25D };
            Cell cell57063 = new Cell() { CellReference = "A21", StyleIndex = (UInt32Value)138U };
            Cell cell57064 = new Cell() { CellReference = "B21", StyleIndex = (UInt32Value)60U };

            row2766.Append(cell57063);
            row2766.Append(cell57064);

            Row row2767 = new Row() { RowIndex = (UInt32Value)22U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.25D };
            Cell cell57065 = new Cell() { CellReference = "A22", StyleIndex = (UInt32Value)272U };
            Cell cell57066 = new Cell() { CellReference = "B22", StyleIndex = (UInt32Value)273U };

            row2767.Append(cell57065);
            row2767.Append(cell57066);

            Row row2768 = new Row() { RowIndex = (UInt32Value)23U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.25D };
            Cell cell57067 = new Cell() { CellReference = "A23", StyleIndex = (UInt32Value)137U };
            Cell cell57068 = new Cell() { CellReference = "B23", StyleIndex = (UInt32Value)59U };

            row2768.Append(cell57067);
            row2768.Append(cell57068);

            Row row2769 = new Row() { RowIndex = (UInt32Value)24U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, DyDescent = 0.25D };
            Cell cell57069 = new Cell() { CellReference = "A24", StyleIndex = (UInt32Value)272U };
            Cell cell57070 = new Cell() { CellReference = "B24", StyleIndex = (UInt32Value)273U };

            row2769.Append(cell57069);
            row2769.Append(cell57070);

            Row row2770 = new Row() { RowIndex = (UInt32Value)25U, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, Height = 15.75D, ThickBot = true, DyDescent = 0.3D };
            Cell cell57071 = new Cell() { CellReference = "A25", StyleIndex = (UInt32Value)274U };
            Cell cell57072 = new Cell() { CellReference = "B25", StyleIndex = (UInt32Value)275U };

            row2770.Append(cell57071);
            row2770.Append(cell57072);

            sheetData7.Append(row2748);
            sheetData7.Append(row2749);
            sheetData7.Append(row2750);
            sheetData7.Append(row2751);
            sheetData7.Append(row2752);
            sheetData7.Append(row2753);
            sheetData7.Append(row2754);
            sheetData7.Append(row2755);
            sheetData7.Append(row2756);
            sheetData7.Append(row2757);
            sheetData7.Append(row2758);
            sheetData7.Append(row2759);
            sheetData7.Append(row2760);
            sheetData7.Append(row2761);
            sheetData7.Append(row2762);
            sheetData7.Append(row2763);
            sheetData7.Append(row2764);
            sheetData7.Append(row2765);
            sheetData7.Append(row2766);
            sheetData7.Append(row2767);
            sheetData7.Append(row2768);
            sheetData7.Append(row2769);
            sheetData7.Append(row2770);
            PageMargins pageMargins19 = new PageMargins() { Left = 0.25D, Right = 0.25D, Top = 1.2916666666666667D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup19 = new PageSetup() { Orientation = OrientationValues.Landscape, Id = "rId1" };

            HeaderFooter headerFooter19 = new HeaderFooter();
            OddHeader oddHeader7 = new OddHeader();
            oddHeader7.Text = "&C&G";
            OddFooter oddFooter7 = new OddFooter();
            oddFooter7.Text = "&C&G";

            headerFooter19.Append(oddHeader7);
            headerFooter19.Append(oddFooter7);
            LegacyDrawingHeaderFooter legacyDrawingHeaderFooter7 = new LegacyDrawingHeaderFooter() { Id = "rId2" };

            worksheet7.Append(sheetProperties7);
            worksheet7.Append(sheetDimension7);
            worksheet7.Append(sheetViews7);
            worksheet7.Append(sheetFormatProperties7);
            worksheet7.Append(columns7);
            worksheet7.Append(sheetData7);
            worksheet7.Append(pageMargins19);
            worksheet7.Append(pageSetup19);
            worksheet7.Append(headerFooter19);
            worksheet7.Append(legacyDrawingHeaderFooter7);

            worksheetPart7.Worksheet = worksheet7;
        }
        // Generates content of vmlDrawingPart7.
        private void GenerateVmlDrawingPart7Content(VmlDrawingPart vmlDrawingPart7)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(vmlDrawingPart7.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\r\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\r\n <o:shapelayout v:ext=\"edit\">\r\n  <o:idmap v:ext=\"edit\" data=\"13\"/>\r\n </o:shapelayout><v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\"\r\n  o:preferrelative=\"t\" path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\">\r\n  <v:stroke joinstyle=\"miter\"/>\r\n  <v:formulas>\r\n   <v:f eqn=\"if lineDrawn pixelLineWidth 0\"/>\r\n   <v:f eqn=\"sum @0 1 0\"/>\r\n   <v:f eqn=\"sum 0 0 @1\"/>\r\n   <v:f eqn=\"prod @2 1 2\"/>\r\n   <v:f eqn=\"prod @3 21600 pixelWidth\"/>\r\n   <v:f eqn=\"prod @3 21600 pixelHeight\"/>\r\n   <v:f eqn=\"sum @0 0 1\"/>\r\n   <v:f eqn=\"prod @6 1 2\"/>\r\n   <v:f eqn=\"prod @7 21600 pixelWidth\"/>\r\n   <v:f eqn=\"sum @8 21600 0\"/>\r\n   <v:f eqn=\"prod @7 21600 pixelHeight\"/>\r\n   <v:f eqn=\"sum @10 21600 0\"/>\r\n  </v:formulas>\r\n  <v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/>\r\n  <o:lock v:ext=\"edit\" aspectratio=\"t\"/>\r\n </v:shapetype><v:shape id=\"CH\" o:spid=\"_x0000_s13313\" type=\"#_x0000_t75\"\r\n  style=\'position:absolute;margin-left:0;margin-top:0;width:180pt;height:64.5pt;\r\n  z-index:1\'>\r\n  <v:imagedata o:relid=\"rId1\" o:title=\"FPR-VISION\"/>\r\n  <o:lock v:ext=\"edit\" rotation=\"t\"/>\r\n </v:shape><v:shape id=\"CF\" o:spid=\"_x0000_s13314\" type=\"#_x0000_t75\" style=\'position:absolute;\r\n  margin-left:0;margin-top:0;width:10in;height:36pt;z-index:2\'>\r\n  <v:imagedata o:relid=\"rId2\" o:title=\"FPR Letterhead Footer-Corp LeftBlock\"/>\r\n  <o:lock v:ext=\"edit\" rotation=\"t\"/>\r\n </v:shape></xml>");
            writer.Flush();
            writer.Close();
        }
        //Volume Trend
      // Volume Trend Report
        private void GenerateWorksheetVolumeTrend(WorksheetPart worksheetPart1)
        {
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheet1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetProperties sheetProperties1 = new SheetProperties() { CodeName = "Sheet14" ,  };
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1:X30"  };
            
           
            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView() { ShowGridLines = false, ZoomScaleNormal = (UInt32Value)66U, WorkbookViewId = (UInt32Value)0U,View = SheetViewValues.PageLayout };
            Selection selection1 = new Selection() { ActiveCell = "A4", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A4" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            Columns columns1 = new Columns();
            Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 5.422223D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column2 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 7.140625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column3 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 11.140625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column4 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 7.5703125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column5 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 8.28515625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column6 = new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)6U, Width = 14.42578125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column7 = new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 9D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column8 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 6.42578125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column9 = new Column() { Min = (UInt32Value)9U, Max = (UInt32Value)9U, Width = 7.5703125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column10 = new Column() { Min = (UInt32Value)10U, Max = (UInt32Value)10U, Width = 6.42578125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column11 = new Column() { Min = (UInt32Value)11U, Max = (UInt32Value)11U, Width = 5.42578125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column12 = new Column() { Min = (UInt32Value)12U, Max = (UInt32Value)12U, Width = 6.140625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column13 = new Column() { Min = (UInt32Value)13U, Max = (UInt32Value)13U, Width = 8.42578125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column14 = new Column() { Min = (UInt32Value)14U, Max = (UInt32Value)14U, Width = 6.28515625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column15 = new Column() { Min = (UInt32Value)15U, Max = (UInt32Value)15U, Width = 8.42578125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column16 = new Column() { Min = (UInt32Value)16U, Max = (UInt32Value)16U, Width = 6.42578125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column17 = new Column() { Min = (UInt32Value)17U, Max = (UInt32Value)17U, Width = 8.28515625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column18 = new Column() { Min = (UInt32Value)18U, Max = (UInt32Value)18U, Width = 7.85546875D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column19 = new Column() { Min = (UInt32Value)19U, Max = (UInt32Value)19U, Width = 7.85546875D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column20 = new Column() { Min = (UInt32Value)20U, Max = (UInt32Value)20U, Width = 6.7109375D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column21 = new Column() { Min = (UInt32Value)21U, Max = (UInt32Value)21U, Width = 8.42578125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column22 = new Column() { Min = (UInt32Value)22U, Max = (UInt32Value)22U, Width = 15.5703125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column23 = new Column() { Min = (UInt32Value)23U, Max = (UInt32Value)23U, Width = 1.7109375D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column24 = new Column() { Min = (UInt32Value)24U, Max = (UInt32Value)24U, Width = 27.5703125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column25 = new Column() { Min = (UInt32Value)25U, Max = (UInt32Value)25U, Width = 0.11D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column26 = new Column() { Min = (UInt32Value)26U, Max = (UInt32Value)26U, Width = 0.11D, CustomWidth = true };

            columns1.Append(column1);
            columns1.Append(column2);
            columns1.Append(column3);
            columns1.Append(column4);
            columns1.Append(column5);
            columns1.Append(column6);
            columns1.Append(column7);
            columns1.Append(column8);
            columns1.Append(column9);
            columns1.Append(column10);
            columns1.Append(column11);
            columns1.Append(column12);
            columns1.Append(column13);
            columns1.Append(column14);
            columns1.Append(column15);
            columns1.Append(column16);
            columns1.Append(column17);
            columns1.Append(column18);
            columns1.Append(column19);
            columns1.Append(column20);
            columns1.Append(column21);
            columns1.Append(column22);
            columns1.Append(column23);
            columns1.Append(column24);
            columns1.Append(column25);
            columns1.Append(column26);

            SheetData sheetData1 = new SheetData();

            Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:25" }, StyleIndex = (UInt32Value)1U,  Height = 15.75D, DyDescent = 0.25D };

            Cell cell1 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)125U, DataType = CellValues.String };
            CellFormula cellFormula1 = new CellFormula();
            cellFormula1.Text = "SETUP!$B$2";
            CellValue cellValue1 = new CellValue();
            cellValue1.Text = "?";
            cellFormula1.CalculateCell = true;
            cell1.Append(cellFormula1);
            cell1.Append(cellValue1);
            Cell cell2 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)126U };
            Cell cell3 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)126U };
            Cell cell4 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)126U };
            Cell cell5 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)126U };
            Cell cell6 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)126U };
            Cell cell7 = new Cell() { CellReference = "G1", StyleIndex = (UInt32Value)126U };
            Cell cell8 = new Cell() { CellReference = "H1", StyleIndex = (UInt32Value)126U };
            Cell cell9 = new Cell() { CellReference = "I1", StyleIndex = (UInt32Value)126U };
            Cell cell10 = new Cell() { CellReference = "J1", StyleIndex = (UInt32Value)126U };
            Cell cell11 = new Cell() { CellReference = "K1", StyleIndex = (UInt32Value)126U };
            Cell cell12 = new Cell() { CellReference = "L1", StyleIndex = (UInt32Value)126U };
            Cell cell13 = new Cell() { CellReference = "M1", StyleIndex = (UInt32Value)126U };
            Cell cell14 = new Cell() { CellReference = "N1", StyleIndex = (UInt32Value)126U };
            Cell cell15 = new Cell() { CellReference = "O1", StyleIndex = (UInt32Value)126U };
            Cell cell16 = new Cell() { CellReference = "P1", StyleIndex = (UInt32Value)126U };
            Cell cell17 = new Cell() { CellReference = "Q1", StyleIndex = (UInt32Value)126U };
            Cell cell18 = new Cell() { CellReference = "R1", StyleIndex = (UInt32Value)126U };
            Cell cell19 = new Cell() { CellReference = "S1", StyleIndex = (UInt32Value)126U };
            Cell cell20 = new Cell() { CellReference = "T1", StyleIndex = (UInt32Value)126U };
            Cell cell21 = new Cell() { CellReference = "U1", StyleIndex = (UInt32Value)126U };
            Cell cell22 = new Cell() { CellReference = "V1", StyleIndex = (UInt32Value)126U };
            Cell cell23 = new Cell() { CellReference = "W1", StyleIndex = (UInt32Value)126U };
            Cell cell24 = new Cell() { CellReference = "X1", StyleIndex = (UInt32Value)129U };

            row1.Append(cell1);
            row1.Append(cell2);
            row1.Append(cell3);
            row1.Append(cell4);
            row1.Append(cell5);
            row1.Append(cell6);
            row1.Append(cell7);
            row1.Append(cell8);
            row1.Append(cell9);
            row1.Append(cell10);
            row1.Append(cell11);
            row1.Append(cell12);
            row1.Append(cell13);
            row1.Append(cell14);
            row1.Append(cell15);
            row1.Append(cell16);
            row1.Append(cell17);
            row1.Append(cell18);
            row1.Append(cell19);
            row1.Append(cell20);
            row1.Append(cell21);
            row1.Append(cell22);
            row1.Append(cell23);
            row1.Append(cell24);

            Row row2 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:25" }, StyleIndex = (UInt32Value)1U, DyDescent = 0.25D };

            Cell cell25 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)123U, DataType = CellValues.String };
            CellFormula cellFormula2 = new CellFormula();
            cellFormula2.Text = "\"Volume Trend Summary for Period Ending \"&TEXT(SETUP!B4,\"MMMMMMMMM DD, YYYY\")&\"\"";
            CellValue cellValue2 = new CellValue();
            cellValue2.Text = "Volume Trend Summary for Period Ending ?";
            cellFormula2.CalculateCell = true;
            cell25.Append(cellFormula2);
            cell25.Append(cellValue2);
            Cell cell26 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)124U };
            Cell cell27 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)124U };
            Cell cell28 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)124U };
            Cell cell29 = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)124U };
            Cell cell30 = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)124U };
            Cell cell31 = new Cell() { CellReference = "G2", StyleIndex = (UInt32Value)124U };
            Cell cell32 = new Cell() { CellReference = "H2", StyleIndex = (UInt32Value)124U };
            Cell cell33 = new Cell() { CellReference = "I2", StyleIndex = (UInt32Value)124U };
            Cell cell34 = new Cell() { CellReference = "J2", StyleIndex = (UInt32Value)124U };
            Cell cell35 = new Cell() { CellReference = "K2", StyleIndex = (UInt32Value)124U };
            Cell cell36 = new Cell() { CellReference = "L2", StyleIndex = (UInt32Value)124U };
            Cell cell37 = new Cell() { CellReference = "M2", StyleIndex = (UInt32Value)124U };
            Cell cell38 = new Cell() { CellReference = "N2", StyleIndex = (UInt32Value)124U };
            Cell cell39 = new Cell() { CellReference = "O2", StyleIndex = (UInt32Value)124U };
            Cell cell40 = new Cell() { CellReference = "P2", StyleIndex = (UInt32Value)124U };
            Cell cell41 = new Cell() { CellReference = "Q2", StyleIndex = (UInt32Value)124U };
            Cell cell42 = new Cell() { CellReference = "R2", StyleIndex = (UInt32Value)124U };
            Cell cell43 = new Cell() { CellReference = "S2", StyleIndex = (UInt32Value)124U };
            Cell cell44 = new Cell() { CellReference = "T2", StyleIndex = (UInt32Value)124U };
            Cell cell45 = new Cell() { CellReference = "U2", StyleIndex = (UInt32Value)124U };
            Cell cell46 = new Cell() { CellReference = "V2", StyleIndex = (UInt32Value)124U };
            Cell cell47 = new Cell() { CellReference = "W2", StyleIndex = (UInt32Value)124U };
            Cell cell48 = new Cell() { CellReference = "X2", StyleIndex = (UInt32Value)129U };

            row2.Append(cell25);
            row2.Append(cell26);
            row2.Append(cell27);
            row2.Append(cell28);
            row2.Append(cell29);
            row2.Append(cell30);
            row2.Append(cell31);
            row2.Append(cell32);
            row2.Append(cell33);
            row2.Append(cell34);
            row2.Append(cell35);
            row2.Append(cell36);
            row2.Append(cell37);
            row2.Append(cell38);
            row2.Append(cell39);
            row2.Append(cell40);
            row2.Append(cell41);
            row2.Append(cell42);
            row2.Append(cell43);
            row2.Append(cell44);
            row2.Append(cell45);
            row2.Append(cell46);
            row2.Append(cell47);
            row2.Append(cell48);
 
           
            Row row4 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:25" }, DyDescent = 0.25D };

            Cell cell73 = new Cell() { CellReference = "A4", StyleIndex = (UInt32Value)124U };
            Cell cell74 = new Cell() { CellReference = "C4", StyleIndex = (UInt32Value)124U };
            Cell cell75 = new Cell() { CellReference = "D4", StyleIndex = (UInt32Value)124U };
            Cell cell76 = new Cell() { CellReference = "E4", StyleIndex = (UInt32Value)124U };
            Cell cell77 = new Cell() { CellReference = "F4", StyleIndex = (UInt32Value)124U };
            Cell cell78 = new Cell() { CellReference = "G4", StyleIndex = (UInt32Value)124U };
            Cell cell79 = new Cell() { CellReference = "H4", StyleIndex = (UInt32Value)124U };
            Cell cell80 = new Cell() { CellReference = "I4", StyleIndex = (UInt32Value)124U };
            Cell cell81 = new Cell() { CellReference = "J4", StyleIndex = (UInt32Value)124U };
            Cell cell82 = new Cell() { CellReference = "K4", StyleIndex = (UInt32Value)124U };
            Cell cell83 = new Cell() { CellReference = "L4", StyleIndex = (UInt32Value)124U };
            Cell cell84 = new Cell() { CellReference = "M4", StyleIndex = (UInt32Value)124U };
            Cell cell85 = new Cell() { CellReference = "N4", StyleIndex = (UInt32Value)124U };
            Cell cell86 = new Cell() { CellReference = "O4", StyleIndex = (UInt32Value)124U };
            Cell cell87 = new Cell() { CellReference = "P4", StyleIndex = (UInt32Value)124U };
            Cell cell88 = new Cell() { CellReference = "Q4", StyleIndex = (UInt32Value)124U };
            Cell cell89 = new Cell() { CellReference = "R4", StyleIndex = (UInt32Value)124U };
            Cell cell90 = new Cell() { CellReference = "S4", StyleIndex = (UInt32Value)124U };
            Cell cell91 = new Cell() { CellReference = "T4", StyleIndex = (UInt32Value)124U };
            Cell cell92 = new Cell() { CellReference = "U4", StyleIndex = (UInt32Value)124U };
            Cell cell93 = new Cell() { CellReference = "V4", StyleIndex = (UInt32Value)124U };
            Cell cell94 = new Cell() { CellReference = "W4", StyleIndex = (UInt32Value)124U };
            Cell cell95 = new Cell() { CellReference = "AA4", StyleIndex = (UInt32Value)187U, DataType = CellValues.Number };
            CellValue cellValue195 = new CellValue();
            cellValue195.Text = ".35";
            cell95.Append(cellValue195);

            row4.Append(cell73);
            row4.Append(cell74);
            row4.Append(cell75);
            row4.Append(cell76);
            row4.Append(cell77);
            row4.Append(cell78);
            row4.Append(cell79);
            row4.Append(cell80);
            row4.Append(cell81);
            row4.Append(cell82);
            row4.Append(cell83);
            row4.Append(cell84);
            row4.Append(cell85);
            row4.Append(cell86);
            row4.Append(cell87);
            row4.Append(cell88);
            row4.Append(cell89);
            row4.Append(cell90);
            row4.Append(cell91);
            row4.Append(cell92);
            row4.Append(cell93);
            row4.Append(cell94);
            row4.Append(cell95);

            sheetData1.Append(row1);
            sheetData1.Append(row2);
            sheetData1.Append(row4);
//Start Loop here to include Headers.
            CoFreedomEntities db = new CoFreedomEntities();
            var contractinfo = (from cf in db.vw_csSCBillingContracts
                               where cf.ContractID == _contractID
                               orderby cf.InvoiceID descending
                               select cf).FirstOrDefault();
            var customer = (from cs in db.ARCustomers
                            where cs.CustomerID == _customerID
                            select cs.CustomerNumber).FirstOrDefault();
            //var ToDate = Convert.ToDateTime(contractinfo.OverageToDate);
            //if (_period != contractinfo.OverageToDate)
            //{
            //    ToDate = Convert.ToDateTime(_period);
            //}
            //var FromDate = Convert.ToDateTime(_startDate).AddDays(-1);
            //if (_overrideDate != null)
            //{
            //    FromDate = Convert.ToDateTime(_overrideDate).AddDays(-1);
            //}

            var ToDate = Convert.ToDateTime(_period).AddDays(1);
            var FromDate = Convert.ToDateTime(_startDate).AddDays(1);
            var vtrend = db.Database.SqlQuery<VolumeTrendModel>("exec csVolumeTrend @vd_FromDate, @vd_ToDate, @vs_Customer, @vs_CustomerNumber", new SqlParameter("@vd_FromDate", FromDate), new SqlParameter("@vd_ToDate", ToDate), new SqlParameter("@vs_Customer", "") , new SqlParameter("@vs_CustomerNumber", customer) ).ToList();

           var query = (from cu in vtrend
                         orderby cu.MeterGroup ascending, cu.PeriodVolume descending
                         select cu).Distinct();
             String MeterGroup = String.Empty;
             int _even = 0;
             UInt32Value _rowIndex = 2;
             UInt32Value StartRow = 4;
            int _firstheader = 1;
             foreach (var row in query)
             {
              
                 _rowIndex++;
                 if (MeterGroup != row.MeterGroup)
                 {
                     Row row4b = new Row() { RowIndex = (UInt32Value)_rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 15.5D, CustomHeight = true, DyDescent = 0.25D };
                     Cell cell74a = new Cell() { CellReference = "A" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
                     Cell cell74aa = new Cell() { CellReference = "B" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
                     Cell cell74aaa = new Cell() { CellReference = "C" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
                     Cell cell75a = new Cell() { CellReference = "D" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
                     Cell cell76a = new Cell() { CellReference = "E" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
                     Cell cell77a = new Cell() { CellReference = "F" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
                     Cell cell78a = new Cell() { CellReference = "G" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
                     Cell cell79a = new Cell() { CellReference = "H" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
                     Cell cell80a = new Cell() { CellReference = "I" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
                     Cell cell81a = new Cell() { CellReference = "J" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
                     Cell cell82a = new Cell() { CellReference = "K" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
                     Cell cell83a = new Cell() { CellReference = "L" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
                     Cell cell84a = new Cell() { CellReference = "M" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
                     Cell cell85a = new Cell() { CellReference = "N" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
                     Cell cell86a = new Cell() { CellReference = "O" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
                     Cell cell87a = new Cell() { CellReference = "P" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
                     Cell cell88a = new Cell() { CellReference = "Q" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
                     Cell cell89aa = new Cell() { CellReference = "R" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
                     
                     UInt32Value EndRow = _rowIndex - 1;
                     if (StartRow <= 7) { StartRow = StartRow - 1; } else { StartRow =StartRow  + 1; }
                     Cell cell90aa = new Cell() { CellReference = "S" + _rowIndex.ToString(),StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)205U, DataType = CellValues.Number };
                     CellFormula LastPeriodTotal = new CellFormula();
                     LastPeriodTotal.Text = _rowIndex < 6U ? "" : "SUM(S" + StartRow.Value.ToString() + ":S" + EndRow.Value.ToString() + ")";
                     LastPeriodTotal.CalculateCell = true;
                     cell90aa.Append(LastPeriodTotal);

                     Cell cell91aa = new Cell() { CellReference = "T" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
                     Cell cell92a = new Cell() { CellReference = "U" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)205U , DataType = CellValues.Number };
                     CellFormula PeriodVolumeTotal = new CellFormula();
                     PeriodVolumeTotal.Text = _rowIndex < 6U ? "" : "SUM(U" + StartRow.Value.ToString() + ":U" + EndRow.Value.ToString() + ")";
                     PeriodVolumeTotal.CalculateCell = true;
                     cell92a.Append(PeriodVolumeTotal);
 
                     Cell cell93a = new Cell() { CellReference = "V" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)205U, DataType = CellValues.Number };
                     CellFormula AvgVolumeTotal = new CellFormula();
                    //   AvgVolumeTotal.Text = _rowIndex < 6U ? "" : "SUM(V" + StartRow.Value.ToString() + ":V" + EndRow.Value.ToString() + ")";
                    AvgVolumeTotal.Text = _rowIndex < 6U ? "" : "SUM(V" + StartRow.Value.ToString() + ":V" + EndRow.Value.ToString() + ")";
                    AvgVolumeTotal.CalculateCell = true;
                     cell93a.Append(AvgVolumeTotal);
                     StartRow = MeterGroup == "" ? 7 : _rowIndex + 2;
                     Cell cell94a = new Cell() { CellReference = "W" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
                     Cell cell95a = new Cell() { CellReference = "X" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };

                     row4b.Append(cell74a);
                     row4b.Append(cell74aa);
                     row4b.Append(cell74aaa);
                     row4b.Append(cell75a);
                     row4b.Append(cell76a);
                     row4b.Append(cell77a);
                     row4b.Append(cell78a);
                     row4b.Append(cell79a);
                     row4b.Append(cell80a);
                     row4b.Append(cell81a);
                     row4b.Append(cell82a);
                     row4b.Append(cell83a);
                     row4b.Append(cell84a);
                     row4b.Append(cell85a);
                     row4b.Append(cell86a);
                     row4b.Append(cell87a);
                     row4b.Append(cell88a);
                     row4b.Append(cell89aa);
                     row4b.Append(cell90aa);
                     row4b.Append(cell91aa);
                     row4b.Append(cell92a);
                     row4b.Append(cell93a);
                     row4b.Append(cell94a);
                     row4b.Append(cell95a);
                     sheetData1.Append(row4b);
                     _rowIndex++;
                     Row row4a = new Row() { RowIndex = (UInt32Value)_rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:25" }, DyDescent = 0.25D, Height =  _firstheader == 1 ? 15.5D : 38.5D , CustomHeight= true };

                     Cell cell73a = new Cell() { CellReference = "A" + _rowIndex.ToString()};
                     cell73a.StyleIndex = _firstheader == 1 ? (UInt32Value)194U : (UInt32Value)194U;
                   
                     CellValue CellValueMeterGroup = new CellValue();
                     CellValueMeterGroup.Text = row.MeterGroup;
                     cell73a.Append(CellValueMeterGroup);
                      
                     row4a.Append(cell73a);
                     
                     sheetData1.Append(row4a);
                     _firstheader = 2;   
                     _rowIndex++;
                     Row row5 = new Row() { RowIndex = (UInt32Value)_rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:25" }, StyleIndex = (UInt32Value)12U, CustomFormat = true, Height = 37.35D, CustomHeight = true, DyDescent = 0.25D };

                     Cell cell96 = new Cell() { CellReference = "A" + _rowIndex.ToString(), StyleIndex = (UInt32Value)47U, DataType = CellValues.String };
                     CellValue cellValue4 = new CellValue();
                     cellValue4.Text = "Line #";

                     cell96.Append(cellValue4);

                     Cell cell97 = new Cell() { CellReference = "B" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue5 = new CellValue();
                     cellValue5.Text = "Device ID";

                     cell97.Append(cellValue5);

                     Cell cell98 = new Cell() { CellReference = "C" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue6 = new CellValue();
                     cellValue6.Text = "Model ";

                     cell98.Append(cellValue6);

                     Cell cell99 = new Cell() { CellReference = "D" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue7 = new CellValue();
                     cellValue7.Text = "Meter Type";

                     cell99.Append(cellValue7);

                     Cell cell100 = new Cell() { CellReference = "E" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue8 = new CellValue();
                     cellValue8.Text = "Serial Number";

                     cell100.Append(cellValue8);

                     Cell cell101 = new Cell() { CellReference = "F" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue9 = new CellValue();
                     cellValue9.Text = "Device Status";

                     cell101.Append(cellValue9);

                     Cell cell102 = new Cell() { CellReference = "G" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue10 = new CellValue();
                     cellValue10.Text = "Location";

                     cell102.Append(cellValue10);

                     Cell cell103 = new Cell() { CellReference = "H" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue11 = new CellValue();
                     cellValue11.Text = "Bldg";

                     cell103.Append(cellValue11);

                     Cell cell104 = new Cell() { CellReference = "I" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue12 = new CellValue();
                     cellValue12.Text = "Dept.";

                     cell104.Append(cellValue12);

                     Cell cell105 = new Cell() { CellReference = "J" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue13 = new CellValue();
                     cellValue13.Text = "Floor";

                     cell105.Append(cellValue13);

                     Cell cell106 = new Cell() { CellReference = "K" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue14 = new CellValue();
                     cellValue14.Text = "User";

                     cell106.Append(cellValue14);

                     Cell cell107 = new Cell() { CellReference = "L" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue15 = new CellValue();
                     cellValue15.Text = "Cost Center";

                     cell107.Append(cellValue15);

                     Cell cell108 = new Cell() { CellReference = "M" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue16 = new CellValue();
                     cellValue16.Text = "Start Date";

                     cell108.Append(cellValue16);

                     Cell cell109 = new Cell() { CellReference = "N" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue17 = new CellValue();
                     cellValue17.Text = "Start Meter";

                     cell109.Append(cellValue17);

                     Cell cell110 = new Cell() { CellReference = "O" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue18 = new CellValue();
                     cellValue18.Text = "End Date";

                     cell110.Append(cellValue18);

                     Cell cell111 = new Cell() { CellReference = "P" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue19 = new CellValue();
                     cellValue19.Text = "End Meter";

                     cell111.Append(cellValue19);

                     Cell cell112 = new Cell() { CellReference = "Q" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue20 = new CellValue();
                     cellValue20.Text = "% of Total Usage";

                     cell112.Append(cellValue20);

                     Cell cell113 = new Cell() { CellReference = "R" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue21 = new CellValue();
                     cellValue21.Text = "% Usage By Meter Group";

                     cell113.Append(cellValue21);

                     Cell cell114 = new Cell() { CellReference = "S" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue22 = new CellValue();
                     cellValue22.Text = "Volume Last Period";

                     cell114.Append(cellValue22);

                     Cell cell115 = new Cell() { CellReference = "T" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue23 = new CellValue();
                     cellValue23.Text = "% Change";

                     cell115.Append(cellValue23);

                     Cell cell116 = new Cell() { CellReference = "U" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue24 = new CellValue();
                     cellValue24.Text = "Volume Current Period";

                     cell116.Append(cellValue24);

                     Cell cell117 = new Cell() { CellReference = "V" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                     CellValue cellValue25 = new CellValue();
                     cellValue25.Text = "Avg. Monthly Vol.";

                     cell117.Append(cellValue25);
                     Cell cell118 = new Cell() { CellReference = "W" + _rowIndex.ToString(), StyleIndex = (UInt32Value)45U };

                     Cell cell119 = new Cell() { CellReference = "X" + _rowIndex.ToString(), StyleIndex = (UInt32Value)46U, DataType = CellValues.String };
                     CellValue cellValue26 = new CellValue();
                     cellValue26.Text = "Comments";

                     cell119.Append(cellValue26);

                     row5.Append(cell96);
                     row5.Append(cell97);
                     row5.Append(cell98);
                     row5.Append(cell99);
                     row5.Append(cell100);
                     row5.Append(cell101);
                     row5.Append(cell102);
                     row5.Append(cell103);
                     row5.Append(cell104);
                     row5.Append(cell105);
                     row5.Append(cell106);
                     row5.Append(cell107);
                     row5.Append(cell108);
                     row5.Append(cell109);
                     row5.Append(cell110);
                     row5.Append(cell111);
                     row5.Append(cell112);
                     row5.Append(cell113);
                     row5.Append(cell114);
                     row5.Append(cell115);
                     row5.Append(cell116);
                     row5.Append(cell117);
                     row5.Append(cell118);
                     row5.Append(cell119);
                     sheetData1.Append(row5);
                     _rowIndex++;

                     Row row6 = new Row() { RowIndex = (UInt32Value)_rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:25" }, StyleIndex = (UInt32Value)12U, DyDescent = 0.25D };

                     Cell cell96x = new Cell() { CellReference = "A" + _rowIndex.ToString(), StyleIndex = (UInt32Value)90U, DataType = CellValues.String };
                     CellValue cellValue4x = new CellValue();
                     cellValue4x.Text = row.LineID.ToString();

                     cell96x.Append(cellValue4x);

                     Cell cell97x = new Cell() { CellReference = "B" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                     CellValue cellValue5x = new CellValue();
                     cellValue5x.Text = row.DeviceID;

                     cell97x.Append(cellValue5x);

                     Cell cell98x = new Cell() { CellReference = "C" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                     CellValue cellValue6x = new CellValue();
                     cellValue6x.Text = row.Model;

                     cell98x.Append(cellValue6x);

                     Cell cell99x = new Cell() { CellReference = "D" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                     CellValue cellValue7x = new CellValue();
                     cellValue7x.Text = row.MeterGroup;

                     cell99x.Append(cellValue7x);

                     Cell cell100x = new Cell() { CellReference = "E" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                     CellValue cellValue8x = new CellValue();
                     cellValue8x.Text = row.SerialNumber;

                     cell100x.Append(cellValue8x);

                     Cell cell101x = new Cell() { CellReference = "F" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                     CellValue cellValue9x = new CellValue();
                     cellValue9x.Text = row.DeviceStatus;

                     cell101x.Append(cellValue9x);

                     Cell cell102x = new Cell() { CellReference = "G" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                     CellValue cellValue10x = new CellValue();
                     cellValue10x.Text = row.Location;

                     cell102x.Append(cellValue10x);

                     Cell cell103x = new Cell() { CellReference = "H" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                     CellValue cellValue11x = new CellValue();
                     cellValue11x.Text = row.Building;

                     cell103x.Append(cellValue11x);

                     Cell cell104x = new Cell() { CellReference = "I" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                     CellValue cellValue12x = new CellValue();
                     cellValue12x.Text = row.Dept;

                     cell104x.Append(cellValue12x);

                     Cell cell105x = new Cell() { CellReference = "J" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                     CellValue cellValue13x = new CellValue();
                     cellValue13x.Text = row.Floor;

                     cell105x.Append(cellValue13x);

                     Cell cell106x = new Cell() { CellReference = "K" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                     CellValue cellValue14x = new CellValue();
                     cellValue14x.Text = row.User;

                     cell106x.Append(cellValue14x);

                     Cell cell107x = new Cell() { CellReference = "L" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                     CellValue cellValue15x = new CellValue();
                     cellValue15x.Text = row.CostCenter;

                     cell107x.Append(cellValue15x);

                     Cell cell108x = new Cell() { CellReference = "M" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                     CellValue cellValue16x = new CellValue();
                     cellValue16x.Text = row.StartDate ==  null ? "" : row.StartDate.Value.ToShortDateString();

                     cell108x.Append(cellValue16x);

                     Cell cell109x = new Cell() { CellReference = "N" + _rowIndex.ToString(), StyleIndex = (UInt32Value)189U, DataType = CellValues.Number };
                     CellValue cellValue17x = new CellValue();
                     cellValue17x.Text = row.StartMeter == null ? "" : row.StartMeter.Value.ToString();

                     cell109x.Append(cellValue17x);

                     Cell cell110x = new Cell() { CellReference = "O" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                     CellValue cellValue18x = new CellValue();
                     cellValue18x.Text = row.EndDate == null ? "" : row.EndDate.Value.ToShortDateString();

                     cell110x.Append(cellValue18x);

                     Cell cell111x = new Cell() { CellReference = "P" + _rowIndex.ToString(), StyleIndex = (UInt32Value)189U, DataType = CellValues.Number };
                     CellValue cellValue19x = new CellValue();
                     cellValue19x.Text = row.EndMeter == null ? "" : row.EndMeter.Value.ToString();

                     cell111x.Append(cellValue19x);

                     Cell cell112x = new Cell() { CellReference = "Q" + _rowIndex.ToString(), StyleIndex = (UInt32Value)206U, DataType = CellValues.Number };
                     CellValue cellValue20x = new CellValue();
                     cellValue20x.Text = row.PercOfTotal == null ? "" : (row.PercOfTotal.Value /100).ToString();

                     cell112x.Append(cellValue20x);

                     Cell cell113x = new Cell() { CellReference = "R" + _rowIndex.ToString(), StyleIndex = (UInt32Value)206U, DataType = CellValues.Number };
                     CellValue cellValue21x = new CellValue();
                     cellValue21x.Text = row.PercOfMeterGroup == null ? "" : (row.PercOfMeterGroup.Value /100).ToString(); ;

                     cell113x.Append(cellValue21x);

                     Cell cell114x = new Cell() { CellReference = "S" + _rowIndex.ToString(), StyleIndex = (UInt32Value)189U, DataType = CellValues.Number };
                     CellValue cellValue22x = new CellValue();
                     cellValue22x.Text = row.LastPeriodVolume == null ? "" : row.LastPeriodVolume.ToString();

                     cell114x.Append(cellValue22x);

                     Cell cell115x = new Cell() { CellReference = "T" + _rowIndex.ToString(), StyleIndex = (UInt32Value)187U, DataType = CellValues.Number };
                     CellValue cellValue23x = new CellValue();
                     CellFormula cellFormula115x = new CellFormula();
                     CellFormula cellFormulaTCell = new CellFormula();
                     //cellFormulaTCell.Text = "=IF(AND(U" + _rowIndex.ToString() + "=0,S" + _rowIndex.ToString() + ">0),-100%,IF(OR(S" + _rowIndex.ToString() + "=0,S" + _rowIndex.ToString() + "=\"-\"),\"-\",IF(AND(S" + _rowIndex.ToString() + "=0,U" + _rowIndex.ToString() + "=0),0%,IF(S" + _rowIndex.ToString() + "<U" + _rowIndex.ToString() + ",SUM(U" + _rowIndex.ToString() + "/S" + _rowIndex.ToString() + ")/U" + _rowIndex.ToString() + ",IF(S" + _rowIndex.ToString() + ">U" + _rowIndex.ToString() + ",SUM(-1*(S" + _rowIndex.ToString() + "-U" + _rowIndex.ToString() + ")/S" + _rowIndex.ToString() + "),\"\")))))";
                     //cellFormulaTCell.Text = "=IF(AND(U" + _rowIndex.ToString() + "=0,S" + _rowIndex.ToString() + ">0),-100%,IF(OR(S" + _rowIndex.ToString() + "=0,S" + _rowIndex.ToString() + "=\"-\"),\"-\",IF(AND(S" + _rowIndex.ToString() + "=0,U" + _rowIndex.ToString() + "=0),0%,IF(S" + _rowIndex.ToString() + "<U" + _rowIndex.ToString() + ",SUM(U" + _rowIndex.ToString() + "/S" + _rowIndex.ToString() + "),IF(S" + _rowIndex.ToString() + ">U" + _rowIndex.ToString() + ",SUM(-(S" + _rowIndex.ToString() + "-U" + _rowIndex.ToString() + ")/U" + _rowIndex.ToString() + "),\"\")))))";
                     cellFormulaTCell.Text = "=IF(AND(U" + _rowIndex.ToString() + "=0,S" + _rowIndex.ToString() + ">0),-100%,IF(OR(S" + _rowIndex.ToString() + "=0,S" + _rowIndex.ToString() + "=\"-\"),\"-\",IF(AND(S" + _rowIndex.ToString() + "=0,U" + _rowIndex.ToString() + "=0),0%,IF(S" + _rowIndex.ToString() + "<U" + _rowIndex.ToString() + ",SUM(U" + _rowIndex.ToString() + "-S" + _rowIndex.ToString() + ")/S" + _rowIndex.ToString() + ",IF(S" + _rowIndex.ToString() + ">U" + _rowIndex.ToString() + ",SUM(-1*(S" + _rowIndex.ToString() + "-U" + _rowIndex.ToString() + ")/S" + _rowIndex.ToString() + "),\"\")))))";

                     //
                    //                                                                                                                                                                                                                                                                                         =IF(AND(U6=0,S6>0),-100%,IF(OR(S6=0,S6="-"),"-",IF(AND(S6=0,U6=0),0%,IF(S6<U6,SUM(U6/S6),IF(S6>U6,SUM(-(S6-U6)/U6),"")))))                      
                     cellValue23x.Text = "";
                     cellFormulaTCell.CalculateCell = true;
                     cell115x.Append(cellFormulaTCell);
                     cell115x.Append(cellValue23x);

                     Cell cell116x = new Cell() { CellReference = "U" + _rowIndex.ToString(), StyleIndex = (UInt32Value)189U, DataType = CellValues.Number };
                     CellValue cellValue24x = new CellValue();
                     cellValue24x.Text = row.PeriodVolume == null ? "" : row.PeriodVolume.Value.ToString();

                     cell116x.Append(cellValue24x);

                     Cell cell117x = new Cell() { CellReference = "V" + _rowIndex.ToString(), StyleIndex = (UInt32Value)189U, DataType = CellValues.Number };
                     CellValue cellValue199 = new CellValue();
                     CellFormula cellFormula199 = new CellFormula();
                     cellFormula199.Text = "=SUM(U" + _rowIndex.ToString() + "/ ((O" + _rowIndex.ToString() + "- M" + _rowIndex.ToString() + ")/30.42))";
                     cellValue199.Text = "";
                     cellFormula199.CalculateCell = true;
                     cell117x.Append(cellFormula199);

                     cell117x.Append(cellValue199);

                     Cell cell118x = new Cell() { CellReference = "W" + _rowIndex.ToString(), StyleIndex = (UInt32Value)195U };
                     CellValue cellValue200 = new CellValue();
                     CellFormula cellFormula200 = new CellFormula();
                     cellFormula200.Text = "=IF(T" + _rowIndex.ToString() + "=\"\",\"\",IF(T" + _rowIndex.ToString() + "=\"-\",\"\",IF(U" + _rowIndex.ToString() + "=0,\"\",IF(OR(T" + _rowIndex.ToString() + "> AA4 ,T" + _rowIndex.ToString() + "<-AA4),\"*\",\"\"))))";
                     cellValue200.Text = "";
                     cellFormula200.CalculateCell = true;
                     cell118x.Append(cellFormula200);

                     cell118x.Append(cellValue200);
                     

                     Cell cell119x = new Cell() { CellReference = "X" + _rowIndex.ToString(), StyleIndex = (UInt32Value)91U, DataType = CellValues.String };
                     CellValue cellValue26x = new CellValue();
                     cellValue26x.Text = "";

                     cell119x.Append(cellValue26x);

                     row6.Append(cell96x);
                     row6.Append(cell97x);
                     row6.Append(cell98x);
                     row6.Append(cell99x);
                     row6.Append(cell100x);
                     row6.Append(cell101x);
                     row6.Append(cell102x);
                     row6.Append(cell103x);
                     row6.Append(cell104x);
                     row6.Append(cell105x);
                     row6.Append(cell106x);
                     row6.Append(cell107x);
                     row6.Append(cell108x);
                     row6.Append(cell109x);
                     row6.Append(cell110x);
                     row6.Append(cell111x);
                     row6.Append(cell112x);
                     row6.Append(cell113x);
                     row6.Append(cell114x);
                     row6.Append(cell115x);
                     row6.Append(cell116x);
                     row6.Append(cell117x);
                     row6.Append(cell118x);
                     row6.Append(cell119x);
                     sheetData1.Append(row6);
                     
                    
                 }
                 else
                 {
                     if (_even == 0)
                     {
                         Row row5 = new Row() { RowIndex = (UInt32Value)_rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:25" }, StyleIndex = (UInt32Value)12U, DyDescent = 0.25D };

                         Cell cell96 = new Cell() { CellReference = "A" + _rowIndex.ToString(), StyleIndex = (UInt32Value)96U, DataType = CellValues.String };
                         CellValue cellValue4 = new CellValue();
                         cellValue4.Text = row.LineID.ToString();

                         cell96.Append(cellValue4);

                         Cell cell97 = new Cell() { CellReference = "B" + _rowIndex.ToString(), StyleIndex = (UInt32Value)192U, DataType = CellValues.String };
                         CellValue cellValue5 = new CellValue();
                         cellValue5.Text = row.DeviceID;

                         cell97.Append(cellValue5);

                         Cell cell98 = new Cell() { CellReference = "C" + _rowIndex.ToString(), StyleIndex = (UInt32Value)192U, DataType = CellValues.String };
                         CellValue cellValue6 = new CellValue();
                         cellValue6.Text = row.Model;

                         cell98.Append(cellValue6);

                         Cell cell99 = new Cell() { CellReference = "D" + _rowIndex.ToString(), StyleIndex = (UInt32Value)192U, DataType = CellValues.String };
                         CellValue cellValue7 = new CellValue();
                         cellValue7.Text = row.MeterGroup;

                         cell99.Append(cellValue7);

                         Cell cell100 = new Cell() { CellReference = "E" + _rowIndex.ToString(), StyleIndex = (UInt32Value)192U, DataType = CellValues.String };
                         CellValue cellValue8 = new CellValue();
                         cellValue8.Text = row.SerialNumber;

                         cell100.Append(cellValue8);

                         Cell cell101 = new Cell() { CellReference = "F" + _rowIndex.ToString(), StyleIndex = (UInt32Value)192U, DataType = CellValues.String };
                         CellValue cellValue9 = new CellValue();
                         cellValue9.Text = row.DeviceStatus;

                         cell101.Append(cellValue9);

                         Cell cell102 = new Cell() { CellReference = "G" + _rowIndex.ToString(), StyleIndex = (UInt32Value)192U, DataType = CellValues.String };
                         CellValue cellValue10 = new CellValue();
                         cellValue10.Text = row.Location;

                         cell102.Append(cellValue10);

                         Cell cell103 = new Cell() { CellReference = "H" + _rowIndex.ToString(), StyleIndex = (UInt32Value)192U, DataType = CellValues.String };
                         CellValue cellValue11 = new CellValue();
                         cellValue11.Text = row.Building;

                         cell103.Append(cellValue11);

                         Cell cell104 = new Cell() { CellReference = "I" + _rowIndex.ToString(), StyleIndex = (UInt32Value)192U, DataType = CellValues.String };
                         CellValue cellValue12 = new CellValue();
                         cellValue12.Text = row.Dept;

                         cell104.Append(cellValue12);

                         Cell cell105 = new Cell() { CellReference = "J" + _rowIndex.ToString(), StyleIndex = (UInt32Value)192U, DataType = CellValues.String };
                         CellValue cellValue13 = new CellValue();
                         cellValue13.Text = row.Floor;

                         cell105.Append(cellValue13);

                         Cell cell106 = new Cell() { CellReference = "K" + _rowIndex.ToString(), StyleIndex = (UInt32Value)192U, DataType = CellValues.String };
                         CellValue cellValue14 = new CellValue();
                         cellValue14.Text = row.User;

                         cell106.Append(cellValue14);

                         Cell cell107 = new Cell() { CellReference = "L" + _rowIndex.ToString(), StyleIndex = (UInt32Value)192U, DataType = CellValues.String };
                         CellValue cellValue15 = new CellValue();
                         cellValue15.Text = row.CostCenter;

                         cell107.Append(cellValue15);

                         Cell cell108 = new Cell() { CellReference = "M" + _rowIndex.ToString(), StyleIndex = (UInt32Value)192U, DataType = CellValues.String };
                         CellValue cellValue16 = new CellValue();
                         cellValue16.Text = row.StartDate == null ? "" : row.StartDate.Value.ToShortDateString();

                         cell108.Append(cellValue16);

                         Cell cell109 = new Cell() { CellReference = "N" + _rowIndex.ToString(), StyleIndex = (UInt32Value)183U, DataType = CellValues.Number };
                         CellValue cellValue17 = new CellValue();
                         cellValue17.Text = row.StartMeter == null ? "" : row.StartMeter.Value.ToString();

                         cell109.Append(cellValue17);

                         Cell cell110 = new Cell() { CellReference = "O" + _rowIndex.ToString(), StyleIndex = (UInt32Value)192U, DataType = CellValues.String };
                         CellValue cellValue18 = new CellValue();
                         cellValue18.Text = row.EndDate == null ? "" : row.EndDate.Value.ToShortDateString();

                         cell110.Append(cellValue18);

                         Cell cell111 = new Cell() { CellReference = "P" + _rowIndex.ToString(), StyleIndex = (UInt32Value)183U, DataType = CellValues.Number };
                         CellValue cellValue19 = new CellValue();
                         cellValue19.Text = row.EndMeter == null ? "" : row.EndMeter.Value.ToString();

                         cell111.Append(cellValue19);

                         Cell cell112 = new Cell() { CellReference = "Q" + _rowIndex.ToString(), StyleIndex = (UInt32Value)207U, DataType = CellValues.Number };
                         CellValue cellValue20 = new CellValue();
                         cellValue20.Text = row.PercOfTotal == null ? "" : (row.PercOfTotal.Value / 100).ToString();

                         cell112.Append(cellValue20);

                         Cell cell113 = new Cell() { CellReference = "R" + _rowIndex.ToString(), StyleIndex = (UInt32Value)207U, DataType = CellValues.Number };
                         CellValue cellValue21 = new CellValue();
                         cellValue21.Text = row.PercOfMeterGroup == null ? "" : (row.PercOfMeterGroup.Value / 100).ToString();

                         cell113.Append(cellValue21);

                         Cell cell114 = new Cell() { CellReference = "S" + _rowIndex.ToString(), StyleIndex = (UInt32Value)183U, DataType = CellValues.Number };
                         CellValue cellValue22 = new CellValue();
                         cellValue22.Text = row.LastPeriodVolume == null ? "" : row.LastPeriodVolume.ToString();

                         cell114.Append(cellValue22);

                         Cell cell115 = new Cell() { CellReference = "T" + _rowIndex.ToString(), StyleIndex = (UInt32Value)182U, DataType = CellValues.Number };
                         CellFormula cellFormula115 = new CellFormula();
                         CellValue cellValue23 = new CellValue();
                         CellFormula cellFormulaTCell = new CellFormula();
                         //cellFormulaTCell.Text = "=IF(AND(U" + _rowIndex.ToString() + "=0,S" + _rowIndex.ToString() + ">0),-100%,IF(OR(S" + _rowIndex.ToString() + "=0,S" + _rowIndex.ToString() + "=\"-\"),\"-\",IF(AND(S" + _rowIndex.ToString() + "=0,U" + _rowIndex.ToString() + "=0),0%,IF(S" + _rowIndex.ToString() + "<U" + _rowIndex.ToString() + ",SUM(U" + _rowIndex.ToString() + "-S" + _rowIndex.ToString() + ")/U" + _rowIndex.ToString() + ",IF(S" + _rowIndex.ToString() + ">U" + _rowIndex.ToString() + ",SUM(-1*(S" + _rowIndex.ToString() + "-U" + _rowIndex.ToString() + ")/S" + _rowIndex.ToString() + "),\"\")))))";
                         //cellFormulaTCell.Text =   "=IF(AND(U" + _rowIndex.ToString() + "=0,S" + _rowIndex.ToString() + ">0),-100%,IF(OR(S" + _rowIndex.ToString() + "=0,S" + _rowIndex.ToString() + "=\"-\"),\"-\",IF(AND(S" + _rowIndex.ToString() + "=0,U" + _rowIndex.ToString() + "=0),0%,IF(S" + _rowIndex.ToString() + "<U" + _rowIndex.ToString() + ",SUM(U" + _rowIndex.ToString() + "-S" + _rowIndex.ToString() + ")/S" + _rowIndex.ToString() + ",IF(S" + _rowIndex.ToString() + ">U" + _rowIndex.ToString() + ",SUM(-1*(S" + _rowIndex.ToString() + "-U" + _rowIndex.ToString() + ")/S" + _rowIndex.ToString() + "),\"\")))))";
                         //cellFormulaTCell.Text = "=IF(AND(U" + _rowIndex.ToString() + "=0,S" + _rowIndex.ToString() + ">0),-100%,IF(OR(S" + _rowIndex.ToString() + "=0,S" + _rowIndex.ToString() + "=\"-\"),\"-\",IF(AND(S" + _rowIndex.ToString() + "=0,U" + _rowIndex.ToString() + "=0),0%,IF(S" + _rowIndex.ToString() + "<U" + _rowIndex.ToString() + ",SUM(U" + _rowIndex.ToString() + "-S" + _rowIndex.ToString() + ")/S" + _rowIndex.ToString() + ",IF(S" + _rowIndex.ToString() + ">U" + _rowIndex.ToString() + ",SUM(-1*(S" + _rowIndex.ToString() + "-U" + _rowIndex.ToString() + ")/S" + _rowIndex.ToString() + "),\"\")))))";
                         cellFormulaTCell.Text = "=IF(AND(U" + _rowIndex.ToString() + "=0,S" + _rowIndex.ToString() + ">0),-100%,IF(OR(S" + _rowIndex.ToString() + "=0,S" + _rowIndex.ToString() + "=\"-\"),\"-\",IF(AND(S" + _rowIndex.ToString() + "=0,U" + _rowIndex.ToString() + "=0),0%,IF(S" + _rowIndex.ToString() + "<U" + _rowIndex.ToString() + ",SUM(U" + _rowIndex.ToString() + "-S" + _rowIndex.ToString() + ")/S" + _rowIndex.ToString() + ",IF(S" + _rowIndex.ToString() + ">U" + _rowIndex.ToString() + ",SUM(-1*(S" + _rowIndex.ToString() + "-U" + _rowIndex.ToString() + ")/S" + _rowIndex.ToString() + "),\"\")))))";

                         cellValue23.Text = "";
                         cellFormulaTCell.CalculateCell = true;
                         cell115.Append(cellFormulaTCell);
                         cell115.Append(cellValue23);

                         Cell cell116 = new Cell() { CellReference = "U" + _rowIndex.ToString(), StyleIndex = (UInt32Value)183U, DataType = CellValues.Number };
                         CellValue cellValue24 = new CellValue();
                         cellValue24.Text = row.PeriodVolume == null ? "" : row.PeriodVolume.Value.ToString();

                         cell116.Append(cellValue24);

                         Cell cell117 = new Cell() { CellReference = "V" + _rowIndex.ToString(), StyleIndex = (UInt32Value)183U, DataType = CellValues.Number };
                         CellValue cellValue199 = new CellValue();
                         CellFormula cellFormula199 = new CellFormula();
                         cellFormula199.Text = "=SUM(U" + _rowIndex.ToString() + "/ ((O" + _rowIndex.ToString() + "- M" + _rowIndex.ToString() + ")/30.42))";
                         cellValue199.Text = "";
                         cellFormula199.CalculateCell = true;
                         cell117.Append(cellFormula199);
                         cell117.Append(cellValue199);

                         Cell cell118 = new Cell() { CellReference = "W" + _rowIndex.ToString(), StyleIndex = (UInt32Value)196U };
                         CellValue cellValue200 = new CellValue();
                         CellFormula cellFormula200 = new CellFormula();
                         cellFormula200.Text = "=IF(T" + _rowIndex.ToString() + "=\"\",\"\",IF(T" + _rowIndex.ToString() + "=\"-\",\"\",IF(U" + _rowIndex.ToString() + "=0,\"\",IF(OR(T" + _rowIndex.ToString() + ">AA4,T" + _rowIndex.ToString() + "<-AA4),\"*\",\"\"))))";
                         cellValue200.Text = "";
                         cellFormula200.CalculateCell = true;
                         cell118.Append(cellFormula200);

                         cell118.Append(cellValue200);
                         Cell cell119 = new Cell() { CellReference = "X" + _rowIndex.ToString(), StyleIndex = (UInt32Value)98U, DataType = CellValues.String };
                         CellValue cellValue26 = new CellValue();
                         cellValue26.Text = row.Comments;

                         cell119.Append(cellValue26);

                         row5.Append(cell96);
                         row5.Append(cell97);
                         row5.Append(cell98);
                         row5.Append(cell99);
                         row5.Append(cell100);
                         row5.Append(cell101);
                         row5.Append(cell102);
                         row5.Append(cell103);
                         row5.Append(cell104);
                         row5.Append(cell105);
                         row5.Append(cell106);
                         row5.Append(cell107);
                         row5.Append(cell108);
                         row5.Append(cell109);
                         row5.Append(cell110);
                         row5.Append(cell111);
                         row5.Append(cell112);
                         row5.Append(cell113);
                         row5.Append(cell114);
                         row5.Append(cell115);
                         row5.Append(cell116);
                         row5.Append(cell117);
                         row5.Append(cell118);
                         row5.Append(cell119);
                         sheetData1.Append(row5);
                         _even = 1;
                     }
                     else
                     {
                         Row row5 = new Row() { RowIndex = (UInt32Value)_rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:25" }, StyleIndex = (UInt32Value)12U, DyDescent = 0.25D };

                         Cell cell96 = new Cell() { CellReference = "A" + _rowIndex.ToString(), StyleIndex = (UInt32Value)90U, DataType = CellValues.String };
                         CellValue cellValue4 = new CellValue();
                         cellValue4.Text = row.LineID.ToString();

                         cell96.Append(cellValue4);

                         Cell cell97 = new Cell() { CellReference = "B" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                         CellValue cellValue5 = new CellValue();
                         cellValue5.Text = row.DeviceID;

                         cell97.Append(cellValue5);

                         Cell cell98 = new Cell() { CellReference = "C" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                         CellValue cellValue6 = new CellValue();
                         cellValue6.Text = row.Model;

                         cell98.Append(cellValue6);

                         Cell cell99 = new Cell() { CellReference = "D" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                         CellValue cellValue7 = new CellValue();
                         cellValue7.Text = row.MeterGroup;

                         cell99.Append(cellValue7);

                         Cell cell100 = new Cell() { CellReference = "E" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                         CellValue cellValue8 = new CellValue();
                         cellValue8.Text = row.SerialNumber;

                         cell100.Append(cellValue8);

                         Cell cell101 = new Cell() { CellReference = "F" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                         CellValue cellValue9 = new CellValue();
                         cellValue9.Text = row.DeviceStatus;

                         cell101.Append(cellValue9);

                         Cell cell102 = new Cell() { CellReference = "G" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                         CellValue cellValue10 = new CellValue();
                         cellValue10.Text = row.Location;

                         cell102.Append(cellValue10);

                         Cell cell103 = new Cell() { CellReference = "H" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                         CellValue cellValue11 = new CellValue();
                         cellValue11.Text = row.Building;

                         cell103.Append(cellValue11);

                         Cell cell104 = new Cell() { CellReference = "I" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                         CellValue cellValue12 = new CellValue();
                         cellValue12.Text = row.Dept;

                         cell104.Append(cellValue12);

                         Cell cell105 = new Cell() { CellReference = "J" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                         CellValue cellValue13 = new CellValue();
                         cellValue13.Text = row.Floor;

                         cell105.Append(cellValue13);

                         Cell cell106 = new Cell() { CellReference = "K" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                         CellValue cellValue14 = new CellValue();
                         cellValue14.Text = row.User;

                         cell106.Append(cellValue14);

                         Cell cell107 = new Cell() { CellReference = "L" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                         CellValue cellValue15 = new CellValue();
                         cellValue15.Text = row.CostCenter;

                         cell107.Append(cellValue15);

                         Cell cell108 = new Cell() { CellReference = "M" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                         CellValue cellValue16 = new CellValue();
                         cellValue16.Text = row.StartDate == null ? "" : row.StartDate.Value.ToShortDateString();

                         cell108.Append(cellValue16);

                         Cell cell109 = new Cell() { CellReference = "N" + _rowIndex.ToString(), StyleIndex = (UInt32Value)189U, DataType = CellValues.Number };
                         CellValue cellValue17 = new CellValue();
                         cellValue17.Text = row.StartMeter == null ? "" : row.StartMeter.Value.ToString();

                         cell109.Append(cellValue17);

                         Cell cell110 = new Cell() { CellReference = "O" + _rowIndex.ToString(), StyleIndex = (UInt32Value)191U, DataType = CellValues.String };
                         CellValue cellValue18 = new CellValue();
                         cellValue18.Text = row.EndDate == null ? "" : row.EndDate.Value.ToShortDateString();

                         cell110.Append(cellValue18);

                         Cell cell111 = new Cell() { CellReference = "P" + _rowIndex.ToString(), StyleIndex = (UInt32Value)189U, DataType = CellValues.Number };
                         CellValue cellValue19 = new CellValue();
                         cellValue19.Text = row.EndMeter == null ? "" : row.EndMeter.Value.ToString();

                         cell111.Append(cellValue19);

                         Cell cell112 = new Cell() { CellReference = "Q" + _rowIndex.ToString(), StyleIndex = (UInt32Value)206U, DataType = CellValues.Number };
                         CellValue cellValue20 = new CellValue();
                         cellValue20.Text = row.PercOfTotal == null ? "" : (row.PercOfTotal.Value / 100).ToString();

                         cell112.Append(cellValue20);

                         Cell cell113 = new Cell() { CellReference = "R" + _rowIndex.ToString(), StyleIndex = (UInt32Value)206U, DataType = CellValues.Number };
                         CellValue cellValue21 = new CellValue();
                         cellValue21.Text = row.PercOfMeterGroup == null ? "" : (row.PercOfMeterGroup.Value / 100).ToString();

                         cell113.Append(cellValue21);

                         Cell cell114 = new Cell() { CellReference = "S" + _rowIndex.ToString(), StyleIndex = (UInt32Value)189U, DataType = CellValues.Number };
                         CellValue cellValue22 = new CellValue();
                         cellValue22.Text = row.LastPeriodVolume == null ? "" : row.LastPeriodVolume.ToString();

                         cell114.Append(cellValue22);

                         Cell cell115 = new Cell() { CellReference = "T" + _rowIndex.ToString(), StyleIndex = (UInt32Value)187U, DataType = CellValues.Number };
                         CellFormula cellFormula115 = new CellFormula();
                         CellValue cellValue23 = new CellValue();
                         //cellFormula115.Text = "=IF(AND(U" + _rowIndex.ToString() + "=0,S" + _rowIndex.ToString() + ">0),-100%,IF(OR(S" + _rowIndex.ToString() + "=0,S" + _rowIndex.ToString() + "=\"-\"),\"-\",IF(AND(S" + _rowIndex.ToString() + "=0,U" + _rowIndex.ToString() + "=0),0%,IF(S" + _rowIndex.ToString() + "<U" + _rowIndex.ToString() + ",SUM(U" + _rowIndex.ToString() + "-S" + _rowIndex.ToString() + ")/U" + _rowIndex.ToString() + ",IF(S" + _rowIndex.ToString() + ">U" + _rowIndex.ToString() + ",SUM(-1*(S" + _rowIndex.ToString() + "-U" + _rowIndex.ToString() + ")/S" + _rowIndex.ToString() + "),\"\")))))";
                         cellFormula115.Text = "=IF(AND(U" + _rowIndex.ToString() + "=0,S" + _rowIndex.ToString() + ">0),-100%,IF(OR(S" + _rowIndex.ToString() + "=0,S" + _rowIndex.ToString() + "=\"-\"),\"-\",IF(AND(S" + _rowIndex.ToString() + "=0,U" + _rowIndex.ToString() + "=0),0%,IF(S" + _rowIndex.ToString() + "<U" + _rowIndex.ToString() + ",SUM(U" + _rowIndex.ToString() + "-S" + _rowIndex.ToString() + ")/S" + _rowIndex.ToString() + ",IF(S" + _rowIndex.ToString() + ">U" + _rowIndex.ToString() + ",SUM(-1*(S" + _rowIndex.ToString() + "-U" + _rowIndex.ToString() + ")/S" + _rowIndex.ToString() + "),\"\")))))";
                         cellValue23.Text = "";
                         cellFormula115.CalculateCell = true;
                         cell115.Append(cellFormula115);
                         cell115.Append(cellValue23);

                         Cell cell116 = new Cell() { CellReference = "U" + _rowIndex.ToString(), StyleIndex = (UInt32Value)189U, DataType = CellValues.Number };
                         CellValue cellValue24 = new CellValue();
                         cellValue24.Text = row.PeriodVolume == null ? "" : row.PeriodVolume.Value.ToString();

                         cell116.Append(cellValue24);

                         Cell cell117 = new Cell() { CellReference = "V" + _rowIndex.ToString(), StyleIndex = (UInt32Value)189U, DataType = CellValues.Number };
                         CellValue cellValue199 = new CellValue();
                         CellFormula cellFormula199 = new CellFormula();
                         cellFormula199.Text = "=SUM(U" + _rowIndex.ToString() + "/ ((O" + _rowIndex.ToString() + "- M" + _rowIndex.ToString() + ")/30.42))";
                         cellValue199.Text = "";
                         cellFormula199.CalculateCell = true;
                         cell117.Append(cellFormula199);
                         cell117.Append(cellValue199);

                         Cell cell118 = new Cell() { CellReference = "W" + _rowIndex.ToString(), StyleIndex = (UInt32Value)195U };
                         CellValue cellValue200 = new CellValue();
                         CellFormula cellFormula200 = new CellFormula();
                         cellFormula200.Text = "=IF(T" + _rowIndex.ToString() + "=\"\",\"\",IF(T" + _rowIndex.ToString() + "=\"-\",\"\",IF(U" + _rowIndex.ToString() + "=0,\"\",IF(OR(T" + _rowIndex.ToString() + ">AA4,T" + _rowIndex.ToString() + "<-AA4),\"*\",\"\"))))";
                         cellValue200.Text = "";
                         cellFormula200.CalculateCell = true;
                         cell118.Append(cellFormula200);

                         cell118.Append(cellValue200);
                         Cell cell119 = new Cell() { CellReference = "X" + _rowIndex.ToString(), StyleIndex = (UInt32Value)91U, DataType = CellValues.String };
                         CellValue cellValue26 = new CellValue();
                         cellValue26.Text = row.Comments;

                         cell119.Append(cellValue26);

                         row5.Append(cell96);
                         row5.Append(cell97);
                         row5.Append(cell98);
                         row5.Append(cell99);
                         row5.Append(cell100);
                         row5.Append(cell101);
                         row5.Append(cell102);
                         row5.Append(cell103);
                         row5.Append(cell104);
                         row5.Append(cell105);
                         row5.Append(cell106);
                         row5.Append(cell107);
                         row5.Append(cell108);
                         row5.Append(cell109);
                         row5.Append(cell110);
                         row5.Append(cell111);
                         row5.Append(cell112);
                         row5.Append(cell113);
                         row5.Append(cell114);
                         row5.Append(cell115);
                         row5.Append(cell116);
                         row5.Append(cell117);
                         row5.Append(cell118);
                         row5.Append(cell119);
                         sheetData1.Append(row5);
                         _even = 0;
                     }

                 }
                
                 MeterGroup = row.MeterGroup;
             }
             _rowIndex++;
            Row row6x = new Row() { RowIndex = (UInt32Value)_rowIndex , Spans = new ListValue<StringValue>() { InnerText = "1:25" },  DyDescent = 0.25D };
            Cell cell120 = new Cell() { CellReference = "A" + _rowIndex .ToString() , StyleIndex = (UInt32Value)174U };
            Cell cell121 = new Cell() { CellReference = "B" + _rowIndex.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell122 = new Cell() { CellReference = "C" + _rowIndex.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell123 = new Cell() { CellReference = "D" + _rowIndex.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell124 = new Cell() { CellReference = "E" + _rowIndex.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell125 = new Cell() { CellReference = "F" + _rowIndex.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell126 = new Cell() { CellReference = "G" + _rowIndex.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell127 = new Cell() { CellReference = "H" + _rowIndex.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell128 = new Cell() { CellReference = "I" + _rowIndex.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell129 = new Cell() { CellReference = "J" + _rowIndex.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell130 = new Cell() { CellReference = "K" + _rowIndex.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell131 = new Cell() { CellReference = "L" + _rowIndex.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell132 = new Cell() { CellReference = "M" + _rowIndex.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell133 = new Cell() { CellReference = "N" + _rowIndex.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell134 = new Cell() { CellReference = "O" + _rowIndex.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell135 = new Cell() { CellReference = "P" + _rowIndex.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell136 = new Cell() { CellReference = "Q" + _rowIndex.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell137 = new Cell() { CellReference = "R" + _rowIndex.ToString(), StyleIndex = (UInt32Value)174U };
            UInt32Value EndRow2 = _rowIndex - 1;

            Cell cell138 = new Cell() { CellReference = "S" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)205U, DataType = CellValues.Number };
            CellFormula LastPeriodTotal2 = new CellFormula();
            LastPeriodTotal2.Text = "=SUM(S" + StartRow.Value.ToString() + ":S" + EndRow2.Value.ToString() + ")";
            LastPeriodTotal2.CalculateCell = true;
            cell138.Append(LastPeriodTotal2);
            Cell cell139 = new Cell() { CellReference = "T" + _rowIndex.ToString(), StyleIndex = (UInt32Value)193U };
           // Cell cell139 = new Cell() { CellReference = "T" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)174U };
            Cell cell140 = new Cell() { CellReference = "U" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)205U, DataType = CellValues.Number };
            CellFormula PeriodVolumeTotal2 = new CellFormula();
            PeriodVolumeTotal2.Text =  "=SUM(U" + StartRow.Value.ToString() + ":U" + EndRow2.Value.ToString() + ")";
            PeriodVolumeTotal2.CalculateCell = true;
            cell140.Append(PeriodVolumeTotal2);

            Cell cell141 = new Cell() { CellReference = "V" + _rowIndex.ToString(), StyleIndex = _firstheader == 1 ? (UInt32Value)19U : (UInt32Value)205U, DataType = CellValues.Number };
            CellFormula AvgVolumeTotal2 = new CellFormula();
            AvgVolumeTotal2.Text = "=SUM(V" + StartRow.Value.ToString() + ":V" + EndRow2.Value.ToString() + ")";
            AvgVolumeTotal2.CalculateCell = true;
            cell141.Append(AvgVolumeTotal2);
            StartRow = MeterGroup == "" ? 7 : _rowIndex + 2;
            Cell cell142 = new Cell() { CellReference = "W" + _rowIndex.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell143 = new Cell() { CellReference = "X" + _rowIndex.ToString(), StyleIndex = (UInt32Value)174U };

          
            

            row6x.Append(cell120);
            row6x.Append(cell121);
            row6x.Append(cell122);
            row6x.Append(cell123);
            row6x.Append(cell124);
            row6x.Append(cell125);
            row6x.Append(cell126);
            row6x.Append(cell127);
            row6x.Append(cell128);
            row6x.Append(cell129);
            row6x.Append(cell130);
            row6x.Append(cell131);
            row6x.Append(cell132);
            row6x.Append(cell133);
            row6x.Append(cell134);
            row6x.Append(cell135);
            row6x.Append(cell136);
            row6x.Append(cell137);
            row6x.Append(cell138);
            row6x.Append(cell139);
            row6x.Append(cell140);
            row6x.Append(cell141);
            row6x.Append(cell142);
            row6x.Append(cell143);
            _rowIndex++;
            Row row7 = new Row() { RowIndex = (UInt32Value)_rowIndex, Spans = new ListValue<StringValue>() { InnerText = "1:25" },  DyDescent = 0.25D };            
            Cell cell144z = new Cell() { CellReference = "A" + _rowIndex.ToString(), StyleIndex = (UInt32Value)195U, DataType = CellValues.String };
//            CellValue MessageText = new CellValue();
           CellFormula MessageFormula = new CellFormula();         
            MessageFormula.Text = "=\"* All devices with a net volume change of \" & SUM(AA4 * 100) & \"% are indicated with a red asterisk (*).\"";
//           MessageText.Text = "";
//           cell144.Append(MessageText);
          MessageFormula.CalculateCell = true;
           cell144z.Append(MessageFormula);

            Cell cell145 = new Cell() { CellReference = "B" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell146 = new Cell() { CellReference = "C" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell147 = new Cell() { CellReference = "D" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell148 = new Cell() { CellReference = "E" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell149 = new Cell() { CellReference = "F" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell150 = new Cell() { CellReference = "G" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell151 = new Cell() { CellReference = "H" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell152 = new Cell() { CellReference = "I" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell153 = new Cell() { CellReference = "J" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell154 = new Cell() { CellReference = "K" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell155 = new Cell() { CellReference = "L" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell156 = new Cell() { CellReference = "M" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell157 = new Cell() { CellReference = "N" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell158 = new Cell() { CellReference = "O" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell159 = new Cell() { CellReference = "P" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell160 = new Cell() { CellReference = "Q" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell161 = new Cell() { CellReference = "R" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell162 = new Cell() { CellReference = "S" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell163 = new Cell() { CellReference = "T" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell164 = new Cell() { CellReference = "U" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell165 = new Cell() { CellReference = "V" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell166 = new Cell() { CellReference = "W" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell167 = new Cell() { CellReference = "X" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell168 = new Cell() { CellReference = "Y" + _rowIndex.ToString(), StyleIndex = (UInt32Value)19U };

            row7.Append(cell144z);
            row7.Append(cell145);
            row7.Append(cell146);
            row7.Append(cell147);
            row7.Append(cell148);
            row7.Append(cell149);
            row7.Append(cell150);
            row7.Append(cell151);
            row7.Append(cell152);
            row7.Append(cell153);
            row7.Append(cell154);
            row7.Append(cell155);
            row7.Append(cell156);
            row7.Append(cell157);
            row7.Append(cell158);
            row7.Append(cell159);
            row7.Append(cell160);
            row7.Append(cell161);
            row7.Append(cell162);
            row7.Append(cell163);
            row7.Append(cell164);
            row7.Append(cell165);
            row7.Append(cell166);
            row7.Append(cell167);
            row7.Append(cell168);

         
            
           
            sheetData1.Append(row6x);
          
            sheetData1.Append(row7);
          
            MergeCells mergeCells22 = new MergeCells() { Count = (UInt32Value)3U };
            MergeCell mergeCell122 = new MergeCell() { Reference = "A1:X1" };
            MergeCell mergeCell2 = new MergeCell() { Reference = "A2:X2" };
            MergeCell mergeCell3 = new MergeCell() { Reference = "V5:W5" };

            mergeCells22.Append(mergeCell122);
            mergeCells22.Append(mergeCell2);
            mergeCells22.Append(mergeCell3);


            DataValidations dataValidations1 = new DataValidations() { Count = (UInt32Value)1U };
            DataValidation dataValidation1 = new DataValidation() { AllowBlank = true, ShowInputMessage = true, SequenceOfReferences = new ListValue<StringValue>() { InnerText = "C7:D27" } };

            dataValidations1.Append(dataValidation1);
            PageMargins pageMargins1 = new PageMargins() { Left = 0.25D, Right = 0.25D, Top = 0.85D, Bottom = 0.94791666666666663D, Header = 0.2D, Footer = 0.3D };
            PageSetup pageSetup1 = new PageSetup() { Orientation = OrientationValues.Landscape,FitToHeight = 0U, FitToWidth =1U,Scale = 64U, Id = "rId1" };

            HeaderFooter headerFooter1 = new HeaderFooter();
            OddHeader oddHeader1 = new OddHeader();
            oddHeader1.Text = "&C&G";
            OddFooter oddFooter1 = new OddFooter();
            oddFooter1.Text = "&C&G";

            headerFooter1.Append(oddHeader1);
            headerFooter1.Append(oddFooter1);
            LegacyDrawing legacyDrawing1 = new LegacyDrawing() { Id = "rId2" };
            LegacyDrawingHeaderFooter legacyDrawingHeaderFooter1 = new LegacyDrawingHeaderFooter() { Id = "rId3" };

            worksheet1.Append(sheetProperties1);
            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(columns1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(mergeCells22);
            worksheet1.Append(dataValidations1);
            worksheet1.Append(pageMargins1);
            worksheet1.Append(pageSetup1);
            worksheet1.Append(headerFooter1);
            worksheet1.Append(legacyDrawing1);
            worksheet1.Append(legacyDrawingHeaderFooter1);
           
            worksheetPart1.Worksheet = worksheet1;
        }
        //Quarterly History
        // Generates content of worksheetPart114.
        private void GenerateWorksheetPartQuarterlyHistory(WorksheetPart worksheetPart114)
        {
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheet1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            UInt32Value _rowindex = 3;
           
            int _even = 1;
            String PeriodChange = "";


            DateTime period = Convert.ToDateTime(_period);
            CoFreedomEntities db = new CoFreedomEntities();
            GlobalViewEntities db3 = new GlobalViewEntities();
            var overridedate = Convert.ToDateTime(_overrideDate);
            var query = (from r in db.vw_RevisionInvoiceHistory
                         where r.ContractID == _contractID  && r.OverageToDate >= overridedate
                         orderby r.OverageToDate descending, r.MeterGroup ascending 
                         select r).ToList();

            foreach(var row in query)
            {
                var rollovers = db3.RevisionDatas.Where(o => o.InvoiceID == row.InvoiceID && o.MeterGroupID == row.ContractMeterGroupID).Select(o => o.Rollover).FirstOrDefault();
                if(rollovers != null)
                row.Rollover = (decimal)rollovers;
            }

            SheetProperties sheetProperties1 = new SheetProperties() {  CodeName = "Sheet14" };
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1:J" + query.Count.ToString() };

            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView() { ShowGridLines= false, ZoomScaleNormal = (UInt32Value)100U,   WorkbookViewId = (UInt32Value)0U, View = SheetViewValues.PageLayout };
            Selection selection1 = new Selection() { ActiveCell = "B7", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "B7" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            Columns columns1 = new Columns();
            Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 1.140625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column2 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 26.6D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column3 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 12D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column4 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 12.5703125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column5 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)6U, Width = 18.57D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column6 = new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 12.140625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column7 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 12.5703125D, Style = (UInt32Value)16U, CustomWidth = true };
            Column column8 = new Column() { Min = (UInt32Value)9U, Max = (UInt32Value)9U, Width = 12.28515625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column9 = new Column() { Min = (UInt32Value)10U, Max = (UInt32Value)10U, Width = 13.42578125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column10 = new Column() { Min = (UInt32Value)11U, Max = (UInt32Value)16384U, Width = 12.140625D, Style = (UInt32Value)1U };

            columns1.Append(column1);
            columns1.Append(column2);
            columns1.Append(column3);
            columns1.Append(column4);
            columns1.Append(column5);
            columns1.Append(column6);
            columns1.Append(column7);
            columns1.Append(column8);
            columns1.Append(column9);
            columns1.Append(column10);


        

            SheetData sheetData1 = new SheetData();

            Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 15.75D, DyDescent = 0.25D };

            Cell cell1 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)125U, DataType = CellValues.String };
            CellFormula cellFormula1 = new CellFormula();
            cellFormula1.Text = "SETUP!$B$2";
            CellValue cellValue1 = new CellValue();
            cellValue1.Text = "?";
            cellFormula1.CalculateCell = true;
            cell1.Append(cellFormula1);
            cell1.Append(cellValue1);
            Cell cell2 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)126U };
            Cell cell3 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)126U };
            Cell cell4 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)126U };
            Cell cell5 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)126U };
            Cell cell6 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)126U };
            Cell cell7 = new Cell() { CellReference = "G1", StyleIndex = (UInt32Value)126U };
            Cell cell8 = new Cell() { CellReference = "H1", StyleIndex = (UInt32Value)126U };
            Cell cell9 = new Cell() { CellReference = "I1", StyleIndex = (UInt32Value)126U };
            Cell cell10 = new Cell() { CellReference = "J1", StyleIndex = (UInt32Value)129U };

            row1.Append(cell1);
            row1.Append(cell2);
            row1.Append(cell3);
            row1.Append(cell4);
            row1.Append(cell5);
            row1.Append(cell6);
            row1.Append(cell7);
            row1.Append(cell8);
            row1.Append(cell9);
            row1.Append(cell10);

            Row row2 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 15.75D, DyDescent = 0.25D };

            Cell cell1a = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)123U, DataType = CellValues.String };
            CellFormula cellFormula1a = new CellFormula();
            cellFormula1a.Text = "=\"Quarterly History Through Period Ending \"&TEXT(SETUP!$B$4,\"MMMMMMMMM DD, YYYY\")&\"\"";
            CellValue cellValue1a = new CellValue();
            cellValue1a.Text = "?";
            cellFormula1a.CalculateCell = true;
            cell1a.Append(cellFormula1a);
            cell1a.Append(cellValue1a);
            Cell cell2a = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)126U };
            Cell cell3a = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)126U };
            Cell cell4a = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)126U };
            Cell cell5a = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)126U };
            Cell cell6a = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)126U };
            Cell cell7a = new Cell() { CellReference = "G2", StyleIndex = (UInt32Value)126U };
            Cell cell8a = new Cell() { CellReference = "H2", StyleIndex = (UInt32Value)126U };
            Cell cell9a = new Cell() { CellReference = "I2", StyleIndex = (UInt32Value)126U };
            Cell cell10a = new Cell() { CellReference = "J2", StyleIndex = (UInt32Value)129U };

            row2.Append(cell1a);
            row2.Append(cell2a);
            row2.Append(cell3a);
            row2.Append(cell4a);
            row2.Append(cell5a);
            row2.Append(cell6a);
            row2.Append(cell7a);
            row2.Append(cell8a);
            row2.Append(cell9a);
            row2.Append(cell10a);

          
            Row row3 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 30.75D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell11 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)199U, DataType = CellValues.String };
            CellFormula cellFormula2 = new CellFormula();
                cellFormula2.Text = "=\"Below, you will find quarterly volume trend summaries for \" &SETUP!$B$2&\" (by Meter Group). The value of this data, allows you to review historical trends by meter group.\"";
            CellValue cellValue2 = new CellValue();
            cellValue2.Text = "Below is the historical quarterly volume trend summaries for  (by Meter Group). The value of this data, allows you to review historical trends by meter group.";
            cellFormula2.CalculateCell = true;
            cell11.Append(cellFormula2);
            cell11.Append(cellValue2);
            Cell cell12 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)124U };
            Cell cell13 = new Cell() { CellReference = "D3", StyleIndex = (UInt32Value)124U };
            Cell cell14 = new Cell() { CellReference = "E3", StyleIndex = (UInt32Value)124U };
            Cell cell15 = new Cell() { CellReference = "F3", StyleIndex = (UInt32Value)124U };
            Cell cell16 = new Cell() { CellReference = "G3", StyleIndex = (UInt32Value)124U };
            Cell cell17 = new Cell() { CellReference = "H3", StyleIndex = (UInt32Value)124U };
            Cell cell18 = new Cell() { CellReference = "I3", StyleIndex = (UInt32Value)124U };
            Cell cell19 = new Cell() { CellReference = "J3", StyleIndex = (UInt32Value)124U };
           

            row3.Append(cell11);
            row3.Append(cell12);
            row3.Append(cell13);
            row3.Append(cell14);
            row3.Append(cell15);
            row3.Append(cell16);
            row3.Append(cell17);
            row3.Append(cell18);
            row3.Append(cell19);
            sheetData1.Append(row1);
            sheetData1.Append(row2);
            sheetData1.Append(row3);
            UInt32Value EndRow = 0;
            UInt32Value StartRow = 4;
            foreach (var item in query)
            {

                _rowindex++;
                if (PeriodChange  != item.OverageToDate.Value.ToShortDateString())
                {
                    Row row4a = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 15.5D, CustomHeight = true, DyDescent = 0.25D };
                    Cell cell31a = new Cell() { CellReference = "A" + _rowindex.ToString(), StyleIndex = (UInt32Value)19U };
                    Cell cell32a = new Cell() { CellReference = "B" + _rowindex.ToString() };
                    cell32a.StyleIndex = PeriodChange == "" ? 19U : 209U;
                    Cell cell33a = new Cell() { CellReference = "C" + _rowindex.ToString() };
                    cell33a.StyleIndex = PeriodChange == "" ? 19U : 209U;
                    Cell cell34a = new Cell() { CellReference = "D" + _rowindex.ToString() };
                    cell34a.StyleIndex = PeriodChange == "" ? 19U : 209U;
                    Cell cell35a = new Cell() { CellReference = "E" + _rowindex.ToString() };
                    cell35a.StyleIndex = PeriodChange == "" ? 19U : 209U;
                    Cell cell36a = new Cell() { CellReference = "F" + _rowindex.ToString() };
                    cell36a.StyleIndex = PeriodChange == "" ? 19U : 209U;
                    Cell cell37a = new Cell() { CellReference = "G" + _rowindex.ToString() };
                    cell37a.StyleIndex = PeriodChange == "" ? 19U : 209U;
                    Cell cell38a = new Cell() { CellReference = "H" + _rowindex.ToString() };
                    cell38a.StyleIndex = PeriodChange == "" ? 19U : 209U;
                    CellValue OverageExpenseLabel = new CellValue();
                    OverageExpenseLabel.Text = PeriodChange == "" ? "" : "Total Overage Expense:";
                     cell38a.Append(OverageExpenseLabel);
                    Cell cell39a = new Cell() { CellReference = "I" + _rowindex.ToString()};
                   cell39a.StyleIndex = PeriodChange == "" ? 19U : 209U;
                    CellFormula OverageExpense = new CellFormula();
                    EndRow = _rowindex.Value - 1;
                    OverageExpense.Text = PeriodChange == "" ? "" : "=SUM(I" + StartRow.ToString() + ":I" + EndRow.ToString() + ")";
                    OverageExpense.CalculateCell = true;
                    cell39a.Append(OverageExpense);
                    row4a.Append(cell31a);
                    row4a.Append(cell32a);
                    row4a.Append(cell33a);
                    row4a.Append(cell34a);
                    row4a.Append(cell35a);
                    row4a.Append(cell36a);
                    row4a.Append(cell37a);
                    row4a.Append(cell38a);
                    row4a.Append(cell39a);
                   
                    _rowindex++;

                    Row row4 = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 25.5D, CustomHeight = true, DyDescent = 0.25D };
                    Cell cell31 = new Cell() { CellReference = "A" + _rowindex.ToString(), StyleIndex = (UInt32Value)19U };
                    

                    Cell cell32 = new Cell() { CellReference = "B" + _rowindex.ToString()  };
                    cell32.StyleIndex = PeriodChange == "" ? 14U : 14U;
                    CellValue cellValuePeriod = new CellValue();
                    cellValuePeriod.Text = item.OverageFromDate.Value.ToString("MMMM yyyy") + " - " + item.OverageToDate.Value.ToString("MMMM yyyy");
                    cell32.Append(cellValuePeriod);
                    Cell cell33 = new Cell() { CellReference = "C" + _rowindex.ToString()};
                    cell33.StyleIndex = PeriodChange == "" ? 19U : 19U;
                    Cell cell34 = new Cell() { CellReference = "D" + _rowindex.ToString() };
                    cell34.StyleIndex = PeriodChange == "" ? 19U : 19U;
                    Cell cell35 = new Cell() { CellReference = "E" + _rowindex.ToString()  };
                    cell35.StyleIndex = PeriodChange == "" ? 19U : 19U;
                    Cell cell36 = new Cell() { CellReference = "F" + _rowindex.ToString()  };
                    cell36.StyleIndex = PeriodChange == "" ? 19U : 19U;
                    Cell cell37 = new Cell() { CellReference = "G" + _rowindex.ToString() };
                    cell37.StyleIndex = PeriodChange == "" ? 19U : 19U;
                    Cell cell38 = new Cell() { CellReference = "H" + _rowindex.ToString() };
                    cell38.StyleIndex = PeriodChange == "" ? 19U : 19U;
                    Cell cell39 = new Cell() { CellReference = "I" + _rowindex.ToString()  };
                    cell39.StyleIndex = PeriodChange == "" ? 19U : 19U;

                    row4.Append(cell31);
                    row4.Append(cell32);
                    row4.Append(cell33);
                    row4.Append(cell34);
                    row4.Append(cell35);
                    row4.Append(cell36);
                    row4.Append(cell37);
                    row4.Append(cell38);
                    row4.Append(cell39);

                    _rowindex++;
                    
                    Row row6 = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 35.75D, CustomHeight = true, DyDescent = 0.25D };


                    Cell cell42 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)47U, DataType = CellValues.String };
                    CellValue cellValue6 = new CellValue();
                    cellValue6.Text = "Meter Group";

                    cell42.Append(cellValue6);

                    Cell cell43 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                    CellValue cellValue7 = new CellValue();
                    cellValue7.Text = "Quarterly Contracted Volume";

                    cell43.Append(cellValue7);

                    Cell cell44 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                    CellValue cellValue8 = new CellValue();
                    cellValue8.Text = "Actual Volume For Period";

                    cell44.Append(cellValue8);

                    Cell cell45 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                    CellValue cellValue9 = new CellValue();
                    cellValue9.Text = "Overage/(Underage) For Period";

                    cell45.Append(cellValue9);

                    Cell cell46 = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                    CellValue cellValue10 = new CellValue();
                    cellValue10.Text = "Accrued Rollover Pages";

                    cell46.Append(cellValue10);

                    Cell cell47 = new Cell() { CellReference = "G" + _rowindex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                    CellValue cellValue11 = new CellValue();
                    cellValue11.Text = "Net Overage (Underage)";

                    cell47.Append(cellValue11);

                    Cell cell48 = new Cell() { CellReference = "H" + _rowindex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                    CellValue cellValue12 = new CellValue();
                    cellValue12.Text = "Excess CPP";

                    cell48.Append(cellValue12);

                    Cell cell48a = new Cell() { CellReference = "I" + _rowindex.ToString(), StyleIndex = (UInt32Value)46U, DataType = CellValues.String };
                    CellValue cellValue12a = new CellValue();
                    cellValue12a.Text = "Overage Expense";

                    cell48a.Append(cellValue12a);

                    row6.Append(cell42);
                    row6.Append(cell43);
                    row6.Append(cell44);
                    row6.Append(cell45);
                    row6.Append(cell46);
                    row6.Append(cell47);
                    row6.Append(cell48);
                    row6.Append(cell48a);

                    sheetData1.Append(row4a);
                    sheetData1.Append(row4);
                    sheetData1.Append(row6);

                   PeriodChange =  item.OverageToDate.Value.ToShortDateString();
                    _rowindex++;
                    StartRow = _rowindex;
                }
               
                if (_even <= 2)
                {
                    Row row7 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };
                    row7.RowIndex = _rowindex;
                    Cell cell49 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)60U, DataType = CellValues.String };
                    CellValue cellValue13 = new CellValue();
                    cellValue13.Text = item.MeterGroup;

                    cell49.Append(cellValue13);
                    Cell cell50 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)73U, DataType = CellValues.String };
                    CellValue SavingsTypeValue = new CellValue();
                    SavingsTypeValue.Text = item.ContractVolume.Value.ToString("#,##0");
                    cell50.Append(SavingsTypeValue);

                    Cell cell51 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)73U, DataType = CellValues.String };
                    CellValue MonthsValues = new CellValue();
                    MonthsValues.Text = item.ActualVolume.Value.ToString("#,##0");
                    cell51.Append(MonthsValues);

                    Cell cell52 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)177U, DataType = CellValues.Number };
                    CellFormula Overage = new CellFormula();
                   // Overage.Text = "=IF(SUM(C" + _rowindex.ToString() + " - D" + _rowindex.ToString() + ") < 0 ,0,SUM(C" + _rowindex.ToString() + " - D" + _rowindex.ToString() + "))";
                    Overage.Text = "=SUM(D" + _rowindex.ToString() + " - C" + _rowindex.ToString() + ")";
                    Overage.CalculateCell = true;
                    cell52.Append(Overage);

                    Cell cell53 = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)73U, DataType = CellValues.Number };
                    CellValue EndDateValue = new CellValue();
                    EndDateValue.Text = item.Rollover.ToString();
                    cell53.Append(EndDateValue);

                    Cell cell54 = new Cell() { CellReference = "G" + _rowindex.ToString(), StyleIndex = (UInt32Value)177U, DataType = CellValues.Number };
                    CellFormula NetOverage = new CellFormula();
                    //NetOverage.Text = "=IF(SUM(D" + _rowindex.ToString() + " - F" + _rowindex.ToString() + ") < 0 ,0,SUM(D" + _rowindex.ToString() + " - F" + _rowindex.ToString() + "))";
                    NetOverage.Text = "=SUM(E" + _rowindex.ToString() + " - F" + _rowindex.ToString() + ")";
                    NetOverage.CalculateCell = true;
                    cell54.Append(NetOverage);

                    Cell cell55 = new Cell() { CellReference = "H" + _rowindex.ToString(), StyleIndex = (UInt32Value)200U, DataType = CellValues.Number };
                    CellValue CostSavingsAmountValue = new CellValue();
                    CostSavingsAmountValue.Text = item.CPP.Value.ToString();
                    cell55.Append(CostSavingsAmountValue);

                    Cell cell56 = new Cell() { CellReference = "I" + _rowindex.ToString(), StyleIndex = (UInt32Value)213U, DataType = CellValues.Number };
                    CellFormula OverageCharge = new CellFormula();
                    OverageCharge.Text = "=IF(SUM(G" + _rowindex.ToString() + " * H" + _rowindex.ToString() + ") < 0, 0,SUM(G" + _rowindex.ToString() + " * H" + _rowindex.ToString() + "))";
                    OverageCharge.CalculateCell = true;
                    cell56.Append(OverageCharge);
                  

                    row7.Append(cell49);
                    row7.Append(cell50);
                    row7.Append(cell51);
                    row7.Append(cell52);
                    row7.Append(cell53);
                    row7.Append(cell54);
                    row7.Append(cell55);
                    row7.Append(cell56);
                    sheetData1.Append(row7);
                }

                if (_even > 2 && _even < 5)
                {

                    Row row7 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };
                    row7.RowIndex = _rowindex;
                    Cell cell49 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)58U, DataType = CellValues.String };
                    CellValue cellValue13 = new CellValue();
                    cellValue13.Text = item.MeterGroup;

                    cell49.Append(cellValue13);
                    Cell cell50 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)6U, DataType = CellValues.String };
                    CellValue SavingsTypeValue = new CellValue();
                    SavingsTypeValue.Text = item.ContractVolume.Value.ToString("#,##0");
                    cell50.Append(SavingsTypeValue);

                    Cell cell51 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)6U, DataType = CellValues.String };
                    CellValue MonthsValues = new CellValue();
                    MonthsValues.Text = item.ActualVolume.Value.ToString("#,##0");
                    cell51.Append(MonthsValues);

                    Cell cell52 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)176U, DataType = CellValues.Number };
                    CellFormula Overage = new CellFormula();
                    Overage.Text = "=SUM(D" + _rowindex.ToString() + " - C" + _rowindex.ToString() + ")";
                    Overage.CalculateCell = true;
                    cell52.Append(Overage);

                    Cell cell53 = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)6U, DataType = CellValues.Number };
                    CellValue EndDateValue = new CellValue();
                    EndDateValue.Text = item.Rollover.ToString();
                    cell53.Append(EndDateValue);

                    Cell cell54 = new Cell() { CellReference = "G" + _rowindex.ToString(), StyleIndex = (UInt32Value)176U, DataType = CellValues.Number };
                    CellFormula NetOverage = new CellFormula();
                    NetOverage.Text = "=SUM(E" + _rowindex.ToString() + " - F" + _rowindex.ToString() + ")";
                    NetOverage.CalculateCell = true;
                    cell54.Append(NetOverage);

                    Cell cell55 = new Cell() { CellReference = "H" + _rowindex.ToString(), StyleIndex = (UInt32Value)201U, DataType = CellValues.Number };
                    CellValue CostSavingsAmountValue = new CellValue();
                    CostSavingsAmountValue.Text = item.CPP.Value.ToString();
                    cell55.Append(CostSavingsAmountValue);

                    Cell cell56 = new Cell() { CellReference = "I" + _rowindex.ToString(), StyleIndex = (UInt32Value)212U, DataType = CellValues.String };
                    CellFormula OverageCharge = new CellFormula();
                    OverageCharge.Text = "=IF(SUM(G" + _rowindex.ToString() + " * H" + _rowindex.ToString() + ") < 0, 0,SUM(G" + _rowindex.ToString() + " * H" + _rowindex.ToString() + "))";
                    OverageCharge.CalculateCell = true;
                    cell56.Append(OverageCharge);

                    row7.Append(cell49);
                    row7.Append(cell50);
                    row7.Append(cell51);
                    row7.Append(cell52);
                    row7.Append(cell53);
                    row7.Append(cell54);
                    row7.Append(cell55);
                    row7.Append(cell56);
                    sheetData1.Append(row7);

                }
                if (_even == 4) _even = 1; else _even++;

            }
       
            _rowindex++;
            Row row127 = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };
            Cell cell837 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)92U };
            Cell cell838 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)119U };
            Cell cell839 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell840 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell841 = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell842 = new Cell() { CellReference = "G" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell842a = new Cell() { CellReference = "H" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell843 = new Cell() { CellReference = "I" + _rowindex.ToString(), StyleIndex = (UInt32Value)121U };

            row127.Append(cell837);
            row127.Append(cell838);
            row127.Append(cell839);
            row127.Append(cell840);
            row127.Append(cell841);
            row127.Append(cell842);
            row127.Append(cell842a);
            row127.Append(cell843);
            sheetData1.Append(row127);
            _rowindex++;
            Row row4b = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 15.5D, CustomHeight = true, DyDescent = 0.25D };
            Cell cell31b = new Cell() { CellReference = "A" + _rowindex.ToString(), StyleIndex = (UInt32Value)19U };
            Cell cell32b = new Cell() { CellReference = "B" + _rowindex.ToString() };
            cell32b.StyleIndex = 19U;
            Cell cell33b = new Cell() { CellReference = "C" + _rowindex.ToString() };
            cell33b.StyleIndex = 19U;
            Cell cell34b = new Cell() { CellReference = "D" + _rowindex.ToString() };
            cell34b.StyleIndex = 19U;
            Cell cell35b = new Cell() { CellReference = "E" + _rowindex.ToString() };
            cell35b.StyleIndex = 19U;
            Cell cell36b = new Cell() { CellReference = "F" + _rowindex.ToString() };
            cell36b.StyleIndex = 19U;
            Cell cell37b = new Cell() { CellReference = "G" + _rowindex.ToString() };
            cell37b.StyleIndex = 19U;
            Cell cell38b = new Cell() { CellReference = "H" + _rowindex.ToString() };
            CellValue OverageLabel2 = new CellValue();
            OverageLabel2.Text = "Total Overage Expense:";
            cell38b.StyleIndex = 209U;
            cell38b.Append(OverageLabel2);
            Cell cell39b = new Cell() { CellReference = "I" + _rowindex.ToString() };
            cell39b.StyleIndex = 209U;
            CellFormula OverageExpense2 = new CellFormula();
            EndRow = _rowindex.Value - 1;
            OverageExpense2.Text = "=SUM(I" + StartRow.ToString() + ":I" + EndRow.ToString() + ")";
            OverageExpense2.CalculateCell = true;
            cell39b.Append(OverageExpense2);
            row4b.Append(cell31b);
            row4b.Append(cell32b);
            row4b.Append(cell33b);
            row4b.Append(cell34b);
            row4b.Append(cell35b);
            row4b.Append(cell36b);
            row4b.Append(cell37b);
            row4b.Append(cell38b);
            row4b.Append(cell39b);
            sheetData1.Append(row4b);

            cellFormula1.CalculateCell = true;
            MergeCells mergeCells1 = new MergeCells() { Count = (UInt32Value)3U };
            MergeCell mergeCell1 = new MergeCell() { Reference = "A1:I1" };
            MergeCell mergeCell2 = new MergeCell() { Reference = "A2:I2" };
            MergeCell mergeCell3 = new MergeCell() { Reference = "A3:I3" };

            mergeCells1.Append(mergeCell1);
            mergeCells1.Append(mergeCell2);
            mergeCells1.Append(mergeCell3);

            DataValidations dataValidations1 = new DataValidations() { Count = (UInt32Value)1U };
            DataValidation dataValidation1 = new DataValidation() { AllowBlank = true, ShowInputMessage = true, SequenceOfReferences = new ListValue<StringValue>() { InnerText = "C7:D27" } };

            dataValidations1.Append(dataValidation1);
            PrintOptions printOptions3 = new PrintOptions() { HorizontalCentered = true };
            PageMargins pageMargins1 = new PageMargins() { Left = 0.333D, Right = 0.333D, Top = 1.3958333333333333D, Bottom = 0.94791666666666663D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup1 = new PageSetup() { Orientation = OrientationValues.Landscape, Id = "rId1" };

            HeaderFooter headerFooter1 = new HeaderFooter();
            OddHeader oddHeader1 = new OddHeader();
            oddHeader1.Text = "&C&G";
            OddFooter oddFooter1 = new OddFooter();
            oddFooter1.Text = "&C&G";

            headerFooter1.Append(oddHeader1);
            headerFooter1.Append(oddFooter1);
           // LegacyDrawing legacyDrawing1 = new LegacyDrawing() { Id = "rId2" };
            LegacyDrawingHeaderFooter legacyDrawingHeaderFooter1 = new LegacyDrawingHeaderFooter() { Id = "rId2" };

            worksheet1.Append(sheetProperties1);
            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(columns1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(mergeCells1);
            worksheet1.Append(dataValidations1);
            worksheet1.Append(printOptions3);
            worksheet1.Append(pageMargins1);
            worksheet1.Append(pageSetup1);
            worksheet1.Append(headerFooter1);
           // worksheet1.Append(legacyDrawing1);
            worksheet1.Append(legacyDrawingHeaderFooter1);

            worksheetPart114.Worksheet = worksheet1;
        }
        //Rollover History
        // Generates content of worksheetPart115.
        private void GenerateWorksheetPartRolloverHistory(WorksheetPart worksheetPart114)
        {
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheet1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            UInt32Value _rowindex = 3;

            int _even = 1;
          
            ExcelRevisionExport db3 = new ExcelRevisionExport();
            var overridedate = Convert.ToDateTime(_overrideDate);
            var perioddate = Convert.ToDateTime(_period);
            var periods = db3.GetRolloverHistory(_contractID).Where(o=> o.Period >= overridedate && o.Period <= perioddate).OrderByDescending(r => r.Period).ToList();
 
                _rolloverTotalString = periods.Sum(o => o.TotalSavings).ToString();

                String RefNum = periods.Count() == 0 ? "31" : periods.Count().ToString();
                SheetProperties sheetProperties1 = new SheetProperties() { CodeName = "Sheet14" };
                SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1:F" + RefNum };

                SheetViews sheetViews1 = new SheetViews();

                SheetView sheetView1 = new SheetView() { ShowGridLines = false, ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U, View = SheetViewValues.PageLayout };
                Selection selection1 = new Selection() { ActiveCell = "B7", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "B7" } };

                sheetView1.Append(selection1);

                sheetViews1.Append(sheetView1);
                SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

                Columns columns1 = new Columns();
                Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 14.1D, Style = (UInt32Value)1U, CustomWidth = true };
                Column column2 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 38.6D, Style = (UInt32Value)1U, CustomWidth = true };
                Column column3 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 14.5D, Style = (UInt32Value)1U, CustomWidth = true };
                Column column4 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 14.5D, Style = (UInt32Value)1U, CustomWidth = true };
                Column column5 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)6U, Width = 15.57D, Style = (UInt32Value)1U, CustomWidth = true };
                Column column6 = new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 12.140625D, Style = (UInt32Value)5U, CustomWidth = true };

                columns1.Append(column1);
                columns1.Append(column2);
                columns1.Append(column3);
                columns1.Append(column4);
                columns1.Append(column5);
                columns1.Append(column6);




                SheetData sheetData1 = new SheetData();

                Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 15.75D, DyDescent = 0.25D };

                Cell cell1 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)125U, DataType = CellValues.String };
                CellFormula cellFormula1 = new CellFormula();
                cellFormula1.Text = "SETUP!$B$2";
                CellValue cellValue1 = new CellValue();
                cellValue1.Text = "?";
                cellFormula1.CalculateCell = true;
                cell1.Append(cellFormula1);
                cell1.Append(cellValue1);
                Cell cell2 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)126U };
                Cell cell3 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)126U };
                Cell cell4 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)126U };
                Cell cell5 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)126U };

                row1.Append(cell1);
                row1.Append(cell2);
                row1.Append(cell3);
                row1.Append(cell4);
                row1.Append(cell5);
                sheetData1.Append(row1);

                Row row2 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };

                Cell cell11 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)123U, DataType = CellValues.String };
                CellFormula cellFormula2 = new CellFormula();
                cellFormula2.Text = "=\"Rollover Savings Through Period Ending \"&TEXT(SETUP!$B$4,\"MMMMMMMMM DD, YYYY\")&\"\"";
                CellValue cellValue2 = new CellValue();
                cellValue2.Text = "=\"Rollover Savings Through Period Ending \"&TEXT(SETUP!$B$4,\"MMMMMMMMM DD, YYYY\")&\"\"";
                cellFormula2.CalculateCell = true;
                cell11.Append(cellFormula2);
                cell11.Append(cellValue2);
                Cell cell12 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)124U };
                Cell cell13 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)124U };
                Cell cell14 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)124U };
                Cell cell15 = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)124U };


                row2.Append(cell11);
                row2.Append(cell12);
                row2.Append(cell13);
                row2.Append(cell14);
                row2.Append(cell15);

                sheetData1.Append(row2);

                Row row3 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 50.50D, CustomHeight = true, DyDescent = 0.25D };
                Cell cell21 = new Cell() { CellReference = "B3", StyleIndex = (UInt32Value)19U };

                Cell cell22 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)199U, DataType = CellValues.String };
                CellValue cellValue3 = new CellValue();
                cellValue3.Text = "The data below details the incremental cost which has been deferred in comparison with your prior operating environment to FPR.  By calculating the product of the Rollover Pages for the Period and your actual cost per page for that specific meter group, FPR is able to determine an overall Rollover Savings by period, below.  This is yet another way in which FPR separates itself from any other company in the marketplace.";

                cell22.Append(cellValue3);
                Cell cell23 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)131U };
                Cell cell24 = new Cell() { CellReference = "D3", StyleIndex = (UInt32Value)131U };
                Cell cell25 = new Cell() { CellReference = "E3", StyleIndex = (UInt32Value)131U };


                row3.Append(cell22);
                row3.Append(cell21);
                row3.Append(cell23);
                row3.Append(cell24);
                row3.Append(cell25);




                sheetData1.Append(row3);

                if (periods.Count() == 0)
                {
                    _rowindex = 5U;
                    Row row6 = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 27.75D, CustomHeight = true, DyDescent = 0.25D };

                    Cell cell41 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)47U, DataType = CellValues.String };
                    CellValue cellValue5 = new CellValue();
                    cellValue5.Text = "Meter Group";

                    cell41.Append(cellValue5);

                    Cell cell42 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                    CellValue cellValue6 = new CellValue();
                    cellValue6.Text = "Rollover Pages For Period";

                    cell42.Append(cellValue6);

                    Cell cell43 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                    CellValue cellValue7 = new CellValue();
                    cellValue7.Text = "Pre-FPR CPP";

                    cell43.Append(cellValue7);

                    Cell cell44 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)46U, DataType = CellValues.String };
                    CellValue cellValue8 = new CellValue();
                    cellValue8.Text = "Rollover Savings";

                    cell44.Append(cellValue8);
                    row6.Append(cell41);
                    row6.Append(cell42);
                    row6.Append(cell43);
                    row6.Append(cell44);

                    sheetData1.Append(row6);
                    _rowindex++;
                    for (var i = _rowindex; i < 31; i++)
                    {

                        if (_even < 3)
                        {
                            Row row37 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };
                            row37.RowIndex = i;
                            Cell cell1885 = new Cell() { CellReference = "B" + i.ToString(), StyleIndex = (UInt32Value)60U };
                            Cell cell251 = new Cell() { CellReference = "C" + i.ToString(), StyleIndex = (UInt32Value)54U };
                            Cell cell252 = new Cell() { CellReference = "D" + i.ToString(), StyleIndex = (UInt32Value)54U };
                            Cell cell257 = new Cell() { CellReference = "E" + i.ToString(), StyleIndex = (UInt32Value)61U };


                            row37.Append(cell1885);
                            row37.Append(cell251);
                            row37.Append(cell252);

                            row37.Append(cell257);
                            sheetData1.Append(row37);




                        }

                        if (_even >= 3 && _even < 5)
                        {
                            Row row39 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 14.25D, CustomHeight = true, DyDescent = 0.25D };
                            row39.RowIndex = i;
                            Cell cell263 = new Cell() { CellReference = "B" + i.ToString(), StyleIndex = (UInt32Value)58U };
                            Cell cell264 = new Cell() { CellReference = "C" + i.ToString(), StyleIndex = (UInt32Value)26U };
                            Cell cell265 = new Cell() { CellReference = "D" + i.ToString(), StyleIndex = (UInt32Value)21U };

                            Cell cell270 = new Cell() { CellReference = "E" + i.ToString(), StyleIndex = (UInt32Value)59U };

                            CellValue cellValue69 = new CellValue();
                            cellValue69.Text = "";


                            row39.Append(cell263);
                            row39.Append(cell264);
                            row39.Append(cell265);

                            row39.Append(cell270);
                            sheetData1.Append(row39);


                        }
                        if (_even == 4) _even = 1; else _even++;
                        _rowindex++;
                    }


                    Row row127 = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };
                    Cell cell837 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)92U };
                    Cell cell838 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)119U };
                    Cell cell839 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };

                    Cell cell843 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)121U };

                    row127.Append(cell837);
                    row127.Append(cell838);
                    row127.Append(cell839);

                    row127.Append(cell843);
                    //  sheetData1.Append(row127);
                    //  _rowindex++;

                    Row row127a = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };
                    Cell cell837a = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)14U };
                    Cell cell838a = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)14U };
                    Cell cell839a = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)14U };

                    Cell cell843a = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)167U };
                    CellFormula CostAvoidanceTotal = new CellFormula();
                    UInt32Value EndValue = _rowindex - 1;
                    CostAvoidanceTotal.Text = "SUM(I7:I" + EndValue + ")";
                    cell843a.Append(CostAvoidanceTotal);
                    row127a.Append(cell837a);
                    row127a.Append(cell838a);
                    row127a.Append(cell839a);

                    row127a.Append(cell843a);
                    sheetData1.Append(row127a);

                    Row row127b = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };
                    Cell cell837b = new Cell() { CellReference = "B5", StyleIndex = (UInt32Value)14U };
                    Cell cell838b = new Cell() { CellReference = "C5", StyleIndex = (UInt32Value)14U };
                    Cell cell839b = new Cell() { CellReference = "D5", StyleIndex = (UInt32Value)14U };

                    Cell cell843b = new Cell() { CellReference = "E5", StyleIndex = (UInt32Value)167U };
                    CellFormula RollOverTotal = new CellFormula();

                    RollOverTotal.Text = "SUM(I7:I" + EndValue + ")";
                    cell843b.Append(RollOverTotal);
                    row127b.Append(cell837b);
                    row127b.Append(cell838b);
                    row127b.Append(cell839b);
                    row127b.Append(cell843b);
                    sheetData1.Append(row127b);
                }
                else
                {
                    UInt32Value StartRow = 5;
                   
                _rowindex++;
                foreach (var period in periods)
                {
                  
                    Row row5 = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 15.5D, CustomHeight = true, DyDescent = 0.25D };
                    Cell cell31b = new Cell() { CellReference = "A" + _rowindex.ToString() };
                    Cell cell32b = new Cell() { CellReference = "B" + _rowindex.ToString(), DataType = CellValues.String };
                    cell32b.StyleIndex = 194U;
                    CellValue cellPeriodValue = new CellValue();
                    cellPeriodValue.Text = period.StartDate.Value.ToString("MMMM yyyy") + " - " + period.Period.Value.ToString("MMMM yyyy");
                    cell32b.Append(cellPeriodValue);
                    Cell cell33b = new Cell() { CellReference = "C" + _rowindex.ToString() };
                    cell33b.StyleIndex = _invoiceID == 0 ? 19U : 164U;
                    Cell cell34b = new Cell() { CellReference = "D" + _rowindex.ToString() };
                    cell34b.StyleIndex = _invoiceID == 0 ? 19U : 164U;
                    Cell cell35b = new Cell() { CellReference = "E" + _rowindex.ToString() };
                    cell35b.StyleIndex = _invoiceID == 0 ? 19U : 164U;


                    row5.Append(cell31b);
                    row5.Append(cell32b);
                    row5.Append(cell33b);
                    row5.Append(cell34b);
                    row5.Append(cell35b);

                    sheetData1.Append(row5);
                    _rowindex++;
                    Row row6 = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 27.75D, CustomHeight = true, DyDescent = 0.25D };

                    Cell cell41 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)47U, DataType = CellValues.String };
                    CellValue cellValue5 = new CellValue();
                    cellValue5.Text = "Meter Group";

                    cell41.Append(cellValue5);

                    Cell cell42 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                    CellValue cellValue6 = new CellValue();
                    cellValue6.Text = "Rollover Pages For Period";

                    cell42.Append(cellValue6);

                    Cell cell43 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                    CellValue cellValue7 = new CellValue();
                    cellValue7.Text = "Pre-FPR CPP";

                    cell43.Append(cellValue7);

                    Cell cell44 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)46U, DataType = CellValues.String };
                    CellValue cellValue8 = new CellValue();
                    cellValue8.Text = "Rollover Savings";

                    cell44.Append(cellValue8);
                    row6.Append(cell41);
                    row6.Append(cell42);
                    row6.Append(cell43);
                    row6.Append(cell44);

                    sheetData1.Append(row6);


                   
                    StartRow = _rowindex.Value;


                    var peroidData = period.Data.Where(o=> o.InvoiceID == period.InvoiceID);
                    foreach (var item in peroidData)
                    {


                        _rowindex++;
 
                        if (_even <= 2)
                        {
                            Row row7 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };
                            row7.RowIndex = _rowindex;
                            Cell cell49 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)60U, DataType = CellValues.String };
                            CellValue cellValue13 = new CellValue();
                            cellValue13.Text = item.ERPMeterGroupDesc;

                            cell49.Append(cellValue13);
                            Cell cell50 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)73U, DataType = CellValues.String };
                            CellValue SavingsTypeValue = new CellValue();
                            SavingsTypeValue.Text = item.Rollover.Value.ToString("#,##0");
                            cell50.Append(SavingsTypeValue);

                            Cell cell51 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)73U, DataType = CellValues.String };
                            CellValue MonthsValues = new CellValue();
                            MonthsValues.Text = item.CPP.ToString("$ .0000");
                            cell51.Append(MonthsValues);

                            Cell cell52 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)61U, DataType = CellValues.Number };
                            CellFormula StartDateValue = new CellFormula();
                            StartDateValue.Text = "=SUM(C" + _rowindex.ToString() + "* D" + _rowindex.ToString() + ")"; ;
                            cell52.Append(StartDateValue);



                            row7.Append(cell49);
                            row7.Append(cell50);
                            row7.Append(cell51);
                            row7.Append(cell52);

                            sheetData1.Append(row7);
                        }

                        if (_even > 2 && _even <= 4)
                        {

                            Row row7 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };
                            row7.RowIndex = _rowindex;
                            Cell cell49 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)58U, DataType = CellValues.String };
                            CellValue cellValue13 = new CellValue();
                            cellValue13.Text = item.ERPMeterGroupDesc;

                            cell49.Append(cellValue13);
                            Cell cell50 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)6U, DataType = CellValues.String };
                            CellValue SavingsTypeValue = new CellValue();
                            SavingsTypeValue.Text = item.Rollover.Value.ToString("#,##0");
                            cell50.Append(SavingsTypeValue);

                            Cell cell51 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)6U, DataType = CellValues.String };
                            CellValue MonthsValues = new CellValue();
                            MonthsValues.Text = item.CPP.ToString("$ .0000");
                            cell51.Append(MonthsValues);

                            Cell cell52 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)59U, DataType = CellValues.String };
                            CellFormula StartDateValue = new CellFormula();
                            StartDateValue.Text = "=SUM(C" + _rowindex.ToString() + "* D" + _rowindex.ToString() + ")"; ;
                            cell52.Append(StartDateValue);



                            row7.Append(cell49);
                            row7.Append(cell50);
                            row7.Append(cell51);
                            row7.Append(cell52);

                            sheetData1.Append(row7);

                        }
                        if (_even == 4) _even = 1; else _even++;


                    }
                    _rowindex++;
                    Row row4a = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };
                    Cell cell31a = new Cell() { CellReference = "A" + _rowindex.ToString() };
                    cell31a.StyleIndex = 19U;
                    Cell cell32a = new Cell() { CellReference = "B" + _rowindex.ToString() };
                    cell32a.StyleIndex = 174U;
                    Cell cell33a = new Cell() { CellReference = "C" + _rowindex.ToString() };
                    cell33a.StyleIndex = 174U;
                    Cell cell34a = new Cell() { CellReference = "D" + _rowindex.ToString() };
                    cell34a.StyleIndex =   174U; ;
                    CellValue TotalLabel2a = new CellValue();
                    TotalLabel2a.Text = "Rollover Savings: ";
                    cell34a.Append(TotalLabel2a);
                 
                    Cell cell35a = new Cell() { CellReference = "E" + _rowindex.ToString() };
                    cell35a.StyleIndex =  174U;
                    CellFormula RollOverTotal = new CellFormula();
                    UInt32Value EndRow = _rowindex - 1;
                    RollOverTotal.Text = "SUM(E" + StartRow.Value.ToString() + ":E" + EndRow.Value.ToString() + ")";
                    StartRow = _rowindex++;
                    RollOverTotal.CalculateCell = true;
                    cell35a.Append(RollOverTotal);
                     
                    row4a.Append(cell31a);
                    row4a.Append(cell32a);
                    row4a.Append(cell33a);
                    row4a.Append(cell34a);
                    row4a.Append(cell35a);
                    
                    sheetData1.Append(row4a);
                    _rowindex++;
                   

                }
                 
                    Row row4 = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };
                    Cell cell31 = new Cell() { CellReference = "A" + _rowindex.ToString() };
                    cell31.StyleIndex = 6U;
                    Cell cell32 = new Cell() { CellReference = "B" + _rowindex.ToString() };
                    cell32.StyleIndex = 6U;
                    Cell cell33 = new Cell() { CellReference = "C" + _rowindex.ToString() };
                    cell33.StyleIndex = 6U;
                    Cell cell34 = new Cell() { CellReference = "D" + _rowindex.ToString() };
                    cell34.StyleIndex = 174U; ;
                    CellValue TotalLabel2 = new CellValue();
                    TotalLabel2.Text = "Total Rollover Savings: ";
                    cell34.Append(TotalLabel2);
                    Cell cell35 = new Cell() { CellReference = "E" + _rowindex.ToString() };
                    cell35.StyleIndex = 174U;
                    CellValue RollOverTotal2 = new CellValue();
                    RollOverTotal2.Text = _rolloverTotalString;          
                    cell35.Append(RollOverTotal2);
                   
                    row4.Append(cell32);
                    row4.Append(cell33);
                    row4.Append(cell34);
                    row4.Append(cell35);

                    sheetData1.Append(row4);


                }

                cellFormula1.CalculateCell = true;
                MergeCells mergeCells1 = new MergeCells() { Count = (UInt32Value)3U };
                MergeCell mergeCell1 = new MergeCell() { Reference = "A1:G1" };
                MergeCell mergeCell2 = new MergeCell() { Reference = "A2:G2" };
                MergeCell mergeCell3 = new MergeCell() { Reference = "A3:G3" };

                mergeCells1.Append(mergeCell1);
                mergeCells1.Append(mergeCell2);
                mergeCells1.Append(mergeCell3);

                DataValidations dataValidations1 = new DataValidations() { Count = (UInt32Value)1U };
                DataValidation dataValidation1 = new DataValidation() { AllowBlank = true, ShowInputMessage = true, SequenceOfReferences = new ListValue<StringValue>() { InnerText = "C7:D27" } };

                dataValidations1.Append(dataValidation1);
                PrintOptions printOptions3 = new PrintOptions() { HorizontalCentered = true };
                PageMargins pageMargins1 = new PageMargins() { Left = 0.553D, Right = 0.333D, Top = 1.3958333333333333D, Bottom = 0.94791666666666663D, Header = 0.3D, Footer = 0.3D };
                PageSetup pageSetup1 = new PageSetup() { Orientation = OrientationValues.Landscape, Id = "rId1" };

                HeaderFooter headerFooter1 = new HeaderFooter();
                OddHeader oddHeader1 = new OddHeader();
                oddHeader1.Text = "&C&G";
                OddFooter oddFooter1 = new OddFooter();
                oddFooter1.Text = "&C&G";

                headerFooter1.Append(oddHeader1);
                headerFooter1.Append(oddFooter1);
                // LegacyDrawing legacyDrawing1 = new LegacyDrawing() { Id = "rId2" };
                LegacyDrawingHeaderFooter legacyDrawingHeaderFooter1 = new LegacyDrawingHeaderFooter() { Id = "rId2" };

                worksheet1.Append(sheetProperties1);
                worksheet1.Append(sheetDimension1);
                worksheet1.Append(sheetViews1);
                worksheet1.Append(sheetFormatProperties1);
                worksheet1.Append(columns1);
                worksheet1.Append(sheetData1);
                worksheet1.Append(mergeCells1);
                worksheet1.Append(dataValidations1);
                worksheet1.Append(printOptions3);
                worksheet1.Append(pageMargins1);
                worksheet1.Append(pageSetup1);
                worksheet1.Append(headerFooter1);
                //worksheet1.Append(legacyDrawing1);
                worksheet1.Append(legacyDrawingHeaderFooter1);

                worksheetPart114.Worksheet = worksheet1;
             
        }
        //Model Matrix
        // Generates content of worksheetPart115.
        private void GenerateWorksheetPartModelMatrix(WorksheetPart ModelMatrixPart)
        {
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheet1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
           

            DateTime period = Convert.ToDateTime(_period);
      
            //Start Loop here to include Headers.
            CoFreedomEntities db = new CoFreedomEntities();
            //var contractinfo = (from cf in db.vw_csSCBillingContracts
            //                   where cf.ContractID == _contractID
            //                   orderby cf.InvoiceID descending
            //                   select cf).FirstOrDefault();
             var ToDate = Convert.ToDateTime(period).AddDays(1);
             var FromDate = Convert.ToDateTime(_startDate);
            var customer = (from cs in db.ARCustomers
                            where cs.CustomerID == _customerID
                            select cs.CustomerNumber).FirstOrDefault();

            var query2 = (from cu in db.vw_ModelMatrix
                          where cu.CustomerID == _customerID && cu.EndMeterDate >= FromDate
                          select new
                          {
                              Model = cu.Model,
                              ModelCategory = cu.ModelCategory,
                              Volume = cu.Volume
                          }).ToList();

            int LineNum = 1;
            var query = (from row in query2
                         group row by new { row.Model, row.ModelCategory } into g
                         orderby g.Key.ModelCategory, g.Key.Model
                         select new
                         {
                             LineNumber = LineNum++,
                             DeviceType = g.Key.ModelCategory,
                             Model = g.Key.Model,
                             ModelCount = g.Count(),
                             TotalVolume = g.Sum(row => row.Volume)
                         }).ToList();


            String MeterGroup = String.Empty;
            
            

            SheetProperties sheetProperties1 = new SheetProperties() { CodeName = "Sheet16" };
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1:F" + query.Count().ToString() };

            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView() { ShowGridLines = false, ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U, View = SheetViewValues.PageLayout };
            Selection selection1 = new Selection() { ActiveCell = "B7", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "B7" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            Columns columns1 = new Columns();
            Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 14.5D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column2 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 14.6D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column3 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 24.5D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column4 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 24.5D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column5 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 15.57D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column6 = new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)7U, Width = 12.140625D, Style = (UInt32Value)5U, CustomWidth = true };

            columns1.Append(column1);
            columns1.Append(column2);
            columns1.Append(column3);
            columns1.Append(column4);
            columns1.Append(column5);
            columns1.Append(column6);




            SheetData sheetData1 = new SheetData();

            Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 15.75D, DyDescent = 0.25D };

            Cell cell1 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)125U, DataType = CellValues.String };
            CellFormula cellFormula1 = new CellFormula();
            cellFormula1.Text = "SETUP!$B$2";
            CellValue cellValue1 = new CellValue();
            cellValue1.Text = "?";
            cellFormula1.CalculateCell = true;
            cell1.Append(cellFormula1);
            cell1.Append(cellValue1);
            Cell cell2 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)126U };
            Cell cell3 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)126U };
            Cell cell4 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)126U };
            Cell cell5 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)126U };

            row1.Append(cell1);
            row1.Append(cell2);
            row1.Append(cell3);
            row1.Append(cell4);
            row1.Append(cell5);
            sheetData1.Append(row1);

            Row row2 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };

            Cell cell11 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)123U, DataType = CellValues.String };
            CellFormula cellFormula2 = new CellFormula();
            cellFormula2.Text = "=\"Model Matrix Period Ending \"&TEXT(SETUP!$B$4,\"MMMMMMMMM DD, YYYY\")&\"\"";
            CellValue cellValue2 = new CellValue();
            cellValue2.Text = "=\"Model Matrix Period Ending \"&TEXT(SETUP!$B$4,\"MMMMMMMMM DD, YYYY\")&\"\"";
            cellFormula2.CalculateCell = true;
            cell11.Append(cellFormula2);
            cell11.Append(cellValue2);
            Cell cell12 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)124U };
            Cell cell13 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)124U };
            Cell cell14 = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)124U };
            Cell cell15 = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)124U };


            row2.Append(cell11);
            row2.Append(cell12);
            row2.Append(cell13);
            row2.Append(cell14);
            row2.Append(cell15);

            sheetData1.Append(row2);

            Row row3 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:10" },  DyDescent = 0.25D };


            Cell cell22 = new Cell() { CellReference = "B3", StyleIndex = (UInt32Value)199U, DataType = CellValues.String };
            CellValue cellValue3 = new CellValue();
            cellValue3.Text = "The matrix below displays the quanitity and volume by Model in tabular format.";
            cell22.Append(cellValue3);

            Cell cell23 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)131U };
            Cell cell24 = new Cell() { CellReference = "D3", StyleIndex = (UInt32Value)131U };
            Cell cell25 = new Cell() { CellReference = "E3", StyleIndex = (UInt32Value)131U };

          
            row3.Append(cell22);
            row3.Append(cell23);
            row3.Append(cell24);
            row3.Append(cell25);
            sheetData1.Append(row3);
            Row row4 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:10" },  DyDescent = 0.25D };
            Cell cell21a = new Cell() { CellReference = "A4", StyleIndex = (UInt32Value)19U };

            Cell cell22a = new Cell() { CellReference = "B4", StyleIndex = (UInt32Value)127U, DataType = CellValues.String };
       
            Cell cell23a = new Cell() { CellReference = "C4", StyleIndex = (UInt32Value)131U };
            Cell cell24a = new Cell() { CellReference = "D4", StyleIndex = (UInt32Value)131U };
            Cell cell25a = new Cell() { CellReference = "E4", StyleIndex = (UInt32Value)131U };

            row4.Append(cell21a);
            row4.Append(cell22a);
            row4.Append(cell23a);
            row4.Append(cell24a);
            row4.Append(cell25a);

            sheetData1.Append(row4);
            int _even = 1;
            UInt32Value _rowindex = 5U;
           
            Row row6 = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 51.75D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell41 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)47U, DataType = CellValues.String };
            CellValue cellValue5 = new CellValue();
            cellValue5.Text = "Line #";

            cell41.Append(cellValue5);

            Cell cell42 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue6 = new CellValue();
            cellValue6.Text = "Device Type";

            cell42.Append(cellValue6);

            Cell cell43 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue7 = new CellValue();
            cellValue7.Text = "Model";

            cell43.Append(cellValue7);

            Cell cell44 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            UInt32Value EndRow = Convert.ToUInt32( query.Count()) + 6 ;
            CellFormula cellFormula52b = new CellFormula();
            cellFormula52b.Text = "=\"(\" & Text(Sum(E6:E" + EndRow + "),\"###,###\") & \") \r\n Quantity by Model \"";
            cellFormula52b.CalculateCell = true;
            cell44.Append(cellFormula52b);
           
            Cell cell45 = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)46U, DataType = CellValues.String };
            CellFormula cellValue9 = new CellFormula();
            cellValue9.Text = "=\"(\" & Text(Sum(F6:F" + EndRow +     "),\"###,###\") & \")\r\n Volume by Model \"";
            cellValue9.CalculateCell = true;
            cell45.Append(cellValue9);
            
            row6.Append(cell41);
            row6.Append(cell42);
            row6.Append(cell43);
            row6.Append(cell44);
            row6.Append(cell45);

            sheetData1.Append(row6);


            foreach (var item in query)
            {

                _rowindex++;
  
                if (_even <= 2)
                {
                    Row row7 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };
                    row7.RowIndex = _rowindex;
                    Cell cell49 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)60U, DataType = CellValues.String };
                    CellValue cellValue13 = new CellValue();
                    cellValue13.Text = item.LineNumber.ToString("#,###");

                    cell49.Append(cellValue13);
                    Cell cell50 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)73U, DataType = CellValues.String };
                    CellValue SavingsTypeValue = new CellValue();
                    SavingsTypeValue.Text = item.DeviceType.ToString();
                    cell50.Append(SavingsTypeValue);

                    Cell cell51 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)73U, DataType = CellValues.String };
                    CellValue MonthsValues = new CellValue();
                    MonthsValues.Text = item.Model.ToString();
                    cell51.Append(MonthsValues);

                    Cell cell52 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)73U, DataType = CellValues.Number };
                    CellValue QuantityModel = new CellValue();
                    QuantityModel.Text = item.ModelCount.ToString("#,###");
                    cell52.Append(QuantityModel);

                    Cell cell53 = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)98U, DataType = CellValues.Number };
                    CellValue VolumeModel = new CellValue();
                    VolumeModel.Text = item.TotalVolume.Value.ToString();
                    cell53.Append(VolumeModel);


                    row7.Append(cell49);
                    row7.Append(cell50);
                    row7.Append(cell51);
                    row7.Append(cell52);
                    row7.Append(cell53);
                    sheetData1.Append(row7);
                }

                if (_even > 2 && _even < 5)
                {

                    Row row7 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };
                    row7.RowIndex = _rowindex;
                    Cell cell49 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)58U, DataType = CellValues.String };
                    CellValue cellValue13 = new CellValue();
                    cellValue13.Text = item.LineNumber.ToString("#,###");

                    cell49.Append(cellValue13);
                    Cell cell50 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)6U, DataType = CellValues.String };
                    CellValue SavingsTypeValue = new CellValue();
                    SavingsTypeValue.Text = item.DeviceType;
                    cell50.Append(SavingsTypeValue);

                    Cell cell51 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)6U, DataType = CellValues.String };
                    CellValue MonthsValues = new CellValue();
                    MonthsValues.Text = item.Model;
                    cell51.Append(MonthsValues);

                    Cell cell52 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)6U, DataType = CellValues.Number };
                    CellValue StartDateValue = new CellValue();
                    StartDateValue.Text = item.ModelCount.ToString();
                    cell52.Append(StartDateValue);

                    Cell cell53 = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)91U, DataType = CellValues.Number };
                    CellValue VolumeModel = new CellValue();
                    VolumeModel.Text = item.TotalVolume.Value.ToString();
                    cell53.Append(VolumeModel);


                    row7.Append(cell49);
                    row7.Append(cell50);
                    row7.Append(cell51);
                    row7.Append(cell52);
                    row7.Append(cell53);
                    sheetData1.Append(row7);

                }
                if (_even == 4) _even = 1; else _even++;

            }
            _rowindex++;
            Row row127 = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };

            Cell cell838 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)92U };
            Cell cell839 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell840 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell841 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell846 = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)121U };

            row127.Append(cell838);
            row127.Append(cell839);
            row127.Append(cell840);
            row127.Append(cell841);
             
            row127.Append(cell846);
            sheetData1.Append(row127);


            
            


            cellFormula1.CalculateCell = true;
            MergeCells mergeCells1 = new MergeCells() { Count = (UInt32Value)4U };
            MergeCell mergeCell1 = new MergeCell() { Reference = "A1:H1" };
            MergeCell mergeCell2 = new MergeCell() { Reference = "A2:H2" };
            MergeCell mergeCell3 = new MergeCell() { Reference = "B3:F3" };
            MergeCell mergeCell4 = new MergeCell() { Reference = "A4:H4" };
            mergeCells1.Append(mergeCell1);
            mergeCells1.Append(mergeCell2);
            mergeCells1.Append(mergeCell3);
            mergeCells1.Append(mergeCell4);
            DataValidations dataValidations1 = new DataValidations() { Count = (UInt32Value)1U };
            DataValidation dataValidation1 = new DataValidation() { AllowBlank = true, ShowInputMessage = true, SequenceOfReferences = new ListValue<StringValue>() { InnerText = "C7:D27" } };

            dataValidations1.Append(dataValidation1);
            PrintOptions printOptions3 = new PrintOptions() { HorizontalCentered = true };
            PageMargins pageMargins1 = new PageMargins() { Left = 0.333D, Right = 0.333D, Top = 1.3958333333333333D, Bottom = 0.94791666666666663D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup1 = new PageSetup() { Orientation = OrientationValues.Landscape, Id = "rId1" };
            
            HeaderFooter headerFooter1 = new HeaderFooter();
            OddHeader oddHeader1 = new OddHeader();
            oddHeader1.Text = "&C&G";
            OddFooter oddFooter1 = new OddFooter();
            oddFooter1.Text = "&C&G";

            headerFooter1.Append(oddHeader1);
            headerFooter1.Append(oddFooter1);
           // LegacyDrawing legacyDrawing1 = new LegacyDrawing() { Id = "rId2" };
            LegacyDrawingHeaderFooter legacyDrawingHeaderFooter1 = new LegacyDrawingHeaderFooter() { Id = "rId2" };
            
            worksheet1.Append(sheetProperties1);
            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(columns1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(mergeCells1);
            worksheet1.Append(dataValidations1);
            worksheet1.Append(printOptions3);
           
            worksheet1.Append(pageMargins1);
            worksheet1.Append(pageSetup1);
            worksheet1.Append(headerFooter1);
           // worksheet1.Append(legacyDrawing1);
            worksheet1.Append(legacyDrawingHeaderFooter1);

            ModelMatrixPart.Worksheet = worksheet1;
        }
        //Cost Avoidance
        // Generates content of worksheetPart1.
        private void GenerateWorksheetPartCostAvoidance(WorksheetPart worksheetPart1)
        {
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheet1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetProperties sheetProperties1 = new SheetProperties() { CodeName = "Sheet8" };
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1:J30" };

            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView() { ShowGridLines = false, ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U, View = SheetViewValues.PageLayout };
            Selection selection1 = new Selection() { ActiveCell = "B7", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "B7" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            Columns columns1 = new Columns();
            Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 8.140625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column2 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 12.5703125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column3 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 16D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column4 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 15.5703125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column5 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)6U, Width = 9.5703125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column6 = new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 28.140625D, Style = (UInt32Value)5U, CustomWidth = true };
            Column column7 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 10.5703125D, Style = (UInt32Value)16U, CustomWidth = true };
            Column column8 = new Column() { Min = (UInt32Value)9U, Max = (UInt32Value)9U, Width = 16.28515625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column9 = new Column() { Min = (UInt32Value)10U, Max = (UInt32Value)10U, Width = 1.42578125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column10 = new Column() { Min = (UInt32Value)11U, Max = (UInt32Value)16384U, Width = 9.140625D, Style = (UInt32Value)1U };

            columns1.Append(column1);
            columns1.Append(column2);
            columns1.Append(column3);
            columns1.Append(column4);
            columns1.Append(column5);
            columns1.Append(column6);
            columns1.Append(column7);
            columns1.Append(column8);
            columns1.Append(column9);
            columns1.Append(column10);
            DateTime overrideDate = Convert.ToDateTime(_overrideDate);
            CustomerPortalEntities db = new CustomerPortalEntities();
            var ToDate = Convert.ToDateTime(_period).AddDays(1);
            var FromDate = Convert.ToDateTime(_startDate).AddDays(1); 
            var query = (from r in db.CostAvoidances
                         where r.CustomerID == _customerID && (r.SavingsDate >= FromDate && r.SavingsDate <= ToDate)
                         orderby r.SavingsDate descending
                         select r).ToList();
           
            SheetData sheetData1 = new SheetData();

            Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 15.75D, DyDescent = 0.25D };

            Cell cell1 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)125U, DataType = CellValues.String };
            CellFormula cellFormula1 = new CellFormula();
            cellFormula1.Text = "SETUP!$B$2";
            CellValue cellValue1 = new CellValue();
            cellValue1.Text = "?";
            cellFormula1.CalculateCell = true;
            cell1.Append(cellFormula1);
            cell1.Append(cellValue1);
            Cell cell2 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)126U };
            Cell cell3 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)126U };
            Cell cell4 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)126U };
            Cell cell5 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)126U };
            Cell cell6 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)126U };
            Cell cell7 = new Cell() { CellReference = "G1", StyleIndex = (UInt32Value)126U };
            Cell cell8 = new Cell() { CellReference = "H1", StyleIndex = (UInt32Value)126U };
            Cell cell9 = new Cell() { CellReference = "I1", StyleIndex = (UInt32Value)126U };
            Cell cell10 = new Cell() { CellReference = "J1", StyleIndex = (UInt32Value)129U };

            row1.Append(cell1);
            row1.Append(cell2);
            row1.Append(cell3);
            row1.Append(cell4);
            row1.Append(cell5);
            row1.Append(cell6);
            row1.Append(cell7);
            row1.Append(cell8);
            row1.Append(cell9);
            row1.Append(cell10);
            sheetData1.Append(row1);
            
                Row row2 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };

                Cell cell11 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)123U, DataType = CellValues.String };
                CellFormula cellFormula2 = new CellFormula();
                cellFormula2.Text = "=IF(B7=\"\",\"\",IF(B8=\"\",\"Cost Avoidance Summary for Period Ending \"&(TEXT(SETUP!$B$4,\"MMMMMMMMM DD, YYYY\"))&\"\",IF(B7=\"\",\"\",\"Cost Avoidance Summary from \"&(TEXT(MIN($B$7:$B$498),\"MMMMMMMMM DD, YYYY\")&\" to \"&(TEXT(MAX($B$7:$B$498),\"MMMMMMMMM DD, YYYY\")&\"\")))))";
                CellValue cellValue2 = new CellValue();
                cellValue2.Text =  "?";
                cellFormula2.CalculateCell = true;
                cell11.Append(cellFormula2);
                cell11.Append(cellValue2);
                Cell cell12 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)124U };
                Cell cell13 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)124U };
                Cell cell14 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)124U };
                Cell cell15 = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)124U };
                Cell cell16 = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)124U };
                Cell cell17 = new Cell() { CellReference = "G2", StyleIndex = (UInt32Value)124U };
                Cell cell18 = new Cell() { CellReference = "H2", StyleIndex = (UInt32Value)124U };
                Cell cell19 = new Cell() { CellReference = "I2", StyleIndex = (UInt32Value)124U };
                Cell cell20 = new Cell() { CellReference = "J2", StyleIndex = (UInt32Value)129U };

                row2.Append(cell11);
                row2.Append(cell12);
                row2.Append(cell13);
                row2.Append(cell14);
                row2.Append(cell15);
                row2.Append(cell16);
                row2.Append(cell17);
                row2.Append(cell18);
                row2.Append(cell19);
                row2.Append(cell20);
                sheetData1.Append(row2);
           
            Row row3 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 30.75D, CustomHeight = true, DyDescent = 0.25D };
            Cell cell21 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)19U };

            Cell cell22 = new Cell() { CellReference = "B3", StyleIndex = (UInt32Value)199U, DataType = CellValues.String };
            CellValue cellValue3 = new CellValue();
            cellValue3.Text = "Below, you will find a summary of the different areas where FPR was able to deliver solutions resulting in cost-avoidance.  Some examples of cost-avoidance solutions include color reduction/elimination strategies, vendor re-negotiation opportunities, and many others.";

            cell22.Append(cellValue3);
            Cell cell23 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)131U };
            Cell cell24 = new Cell() { CellReference = "D3", StyleIndex = (UInt32Value)131U };
            Cell cell25 = new Cell() { CellReference = "E3", StyleIndex = (UInt32Value)131U };
            Cell cell26 = new Cell() { CellReference = "F3", StyleIndex = (UInt32Value)131U };
            Cell cell27 = new Cell() { CellReference = "G3", StyleIndex = (UInt32Value)131U };
            Cell cell28 = new Cell() { CellReference = "H3", StyleIndex = (UInt32Value)131U };
            Cell cell29 = new Cell() { CellReference = "I3", StyleIndex = (UInt32Value)131U };
            Cell cell30 = new Cell() { CellReference = "J3", StyleIndex = (UInt32Value)131U };

            row3.Append(cell21);
            row3.Append(cell22);
            row3.Append(cell23);
            row3.Append(cell24);
            row3.Append(cell25);
            row3.Append(cell26);
            row3.Append(cell27);
            row3.Append(cell28);
            row3.Append(cell29);
            row3.Append(cell30);

            Row row4 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 3D, CustomHeight = true, DyDescent = 0.25D };
            Cell cell31 = new Cell() { CellReference = "A4", StyleIndex = (UInt32Value)19U };
            Cell cell32 = new Cell() { CellReference = "B4", StyleIndex = (UInt32Value)19U };
            Cell cell33 = new Cell() { CellReference = "C4", StyleIndex = (UInt32Value)19U };
            Cell cell34 = new Cell() { CellReference = "D4", StyleIndex = (UInt32Value)19U };
            Cell cell35 = new Cell() { CellReference = "E4", StyleIndex = (UInt32Value)19U };
            Cell cell36 = new Cell() { CellReference = "F4", StyleIndex = (UInt32Value)19U };
            Cell cell37 = new Cell() { CellReference = "G4", StyleIndex = (UInt32Value)19U };
            Cell cell38 = new Cell() { CellReference = "H4", StyleIndex = (UInt32Value)19U };
            Cell cell39 = new Cell() { CellReference = "I4", StyleIndex = (UInt32Value)19U };

            row4.Append(cell31);
            row4.Append(cell32);
            row4.Append(cell33);
            row4.Append(cell34);
            row4.Append(cell35);
            row4.Append(cell36);
            row4.Append(cell37);
            row4.Append(cell38);
            row4.Append(cell39);

            Row row5 = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 3D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell40 = new Cell() { CellReference = "I5", StyleIndex = (UInt32Value)10U };
            CellFormula cellFormula3 = new CellFormula();
            cellFormula3.Text = "SUM(I7:I27)";
            CellValue cellValue4 = new CellValue();
            cellValue4.Text = "0";
            cellFormula3.CalculateCell = true;
            cell40.Append(cellFormula3);
            cell40.Append(cellValue4);

            row5.Append(cell40);

            Row row6 = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 27.75D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell41 = new Cell() { CellReference = "B6", StyleIndex = (UInt32Value)47U, DataType = CellValues.SharedString };
            CellValue cellValue5 = new CellValue();
            cellValue5.Text = "3";

            cell41.Append(cellValue5);

            Cell cell42 = new Cell() { CellReference = "C6", StyleIndex = (UInt32Value)45U, DataType = CellValues.SharedString };
            CellValue cellValue6 = new CellValue();
            cellValue6.Text = "11";

            cell42.Append(cellValue6);

            Cell cell43 = new Cell() { CellReference = "D6", StyleIndex = (UInt32Value)45U, DataType = CellValues.SharedString };
            CellValue cellValue7 = new CellValue();
            cellValue7.Text = "10";

            cell43.Append(cellValue7);

            Cell cell44 = new Cell() { CellReference = "E6", StyleIndex = (UInt32Value)45U, DataType = CellValues.SharedString };
            CellValue cellValue8 = new CellValue();
            cellValue8.Text = "1";

            cell44.Append(cellValue8);

            Cell cell45 = new Cell() { CellReference = "F6", StyleIndex = (UInt32Value)45U, DataType = CellValues.SharedString };
            CellValue cellValue9 = new CellValue();
            cellValue9.Text = "2";

            cell45.Append(cellValue9);

            Cell cell46 = new Cell() { CellReference = "G6", StyleIndex = (UInt32Value)48U, DataType = CellValues.SharedString };
            CellValue cellValue10 = new CellValue();
            cellValue10.Text = "5";

            cell46.Append(cellValue10);

            Cell cell47 = new Cell() { CellReference = "H6", StyleIndex = (UInt32Value)67U, DataType = CellValues.SharedString };
            CellValue cellValue11 = new CellValue();
            cellValue11.Text = "24";

            cell47.Append(cellValue11);

            Cell cell48 = new Cell() { CellReference = "I6", StyleIndex = (UInt32Value)46U, DataType = CellValues.SharedString };
            CellValue cellValue12 = new CellValue();
            cellValue12.Text = "12";

            cell48.Append(cellValue12);

            row6.Append(cell41);
            row6.Append(cell42);
            row6.Append(cell43);
            row6.Append(cell44);
            row6.Append(cell45);
            row6.Append(cell46);
            row6.Append(cell47);
            row6.Append(cell48);

          

            sheetData1.Append(row3);
            sheetData1.Append(row4);
            sheetData1.Append(row5);
            sheetData1.Append(row6);
           
             UInt32Value _rowindex = 6;
          
             int _even = 1;

             if (query.Count() == 0)
             {
                 _rowindex++;
                 Row row7 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };
                 row7.RowIndex = _rowindex;
                 Cell cell49 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)106U, DataType = CellValues.String };
                 CellValue cellValue13 = new CellValue();
                 cellValue13.Text = "There are no Cost Avoidance to report.";
                 cell49.Append(cellValue13);
                 
                 Cell cell50 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)73U, DataType = CellValues.String };
                 CellValue SavingsTypeValue = new CellValue();
                 SavingsTypeValue.Text = "";
                 cell50.Append(SavingsTypeValue);

                 Cell cell51 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)73U, DataType = CellValues.String };
                 CellValue MonthsValues = new CellValue();
                 MonthsValues.Text = "";
                 cell51.Append(MonthsValues);

                 Cell cell52 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)74U, DataType = CellValues.Date };
                 CellValue StartDateValue = new CellValue();
                 StartDateValue.Text = "";
                 cell52.Append(StartDateValue);

                 Cell cell53 = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)74U, DataType = CellValues.Date };
                 CellValue EndDateValue = new CellValue();
                 EndDateValue.Text = "";
                 cell53.Append(EndDateValue);

                 Cell cell54 = new Cell() { CellReference = "G" + _rowindex.ToString(), StyleIndex = (UInt32Value)51U, DataType = CellValues.String };
                 CellValue CommentValue = new CellValue();
                 CommentValue.Text = "";
                 cell54.Append(CommentValue);
                 Cell cell55 = new Cell() { CellReference = "H" + _rowindex.ToString(), StyleIndex = (UInt32Value)75U, DataType = CellValues.Number };
                 CellValue CostSavingsAmountValue = new CellValue();
                 CostSavingsAmountValue.Text = "";
                 cell55.Append(CostSavingsAmountValue);
                 Cell cell56 = new Cell() { CellReference = "I" + _rowindex.ToString(), StyleIndex = (UInt32Value)61U, DataType = CellValues.String };
                 CellValue cellValue14 = new CellValue();
                 cellValue14.Text = "";
       
                 cell56.Append(cellValue14);

                 row7.Append(cell49);
                 row7.Append(cell50);
                 row7.Append(cell51);
                 row7.Append(cell52);
                 row7.Append(cell53);
                 row7.Append(cell54);
                 row7.Append(cell55);
                 row7.Append(cell56);



                 sheetData1.Append(row7);

                
              
             }

             foreach (var item in query)
             {
                 _rowindex++;
                 if (_even <= 2)
                 {
                     Row row7 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };
                     row7.RowIndex = _rowindex;
                     Cell cell49 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)60U, DataType = CellValues.Date };
                     CellValue cellValue13 = new CellValue();
                     cellValue13.Text = item.SavingsDate.Value.ToShortDateString();

                     cell49.Append(cellValue13);
                     Cell cell50 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)73U, DataType = CellValues.String };
                     CellValue SavingsTypeValue = new CellValue();
                     SavingsTypeValue.Text = item.SavingsType.Value == 1 ? "One-Time Savings" : "Monthly";
                     cell50.Append(SavingsTypeValue);

                     Cell cell51 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)73U, DataType = CellValues.String };
                     CellValue MonthsValues = new CellValue();
                     MonthsValues.Text = item.Months.Value.ToString();
                     cell51.Append(MonthsValues);

                     Cell cell52 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)74U, DataType = CellValues.Date };
                     CellValue StartDateValue = new CellValue();
                     StartDateValue.Text = item.SavingsDate.Value.ToShortDateString();
                     cell52.Append(StartDateValue);

                     Cell cell53 = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)74U, DataType = CellValues.Date };
                     CellValue EndDateValue = new CellValue();
                     EndDateValue.Text = item.EndDate.Value.ToShortDateString();
                     cell53.Append(EndDateValue);

                     Cell cell54 = new Cell() { CellReference = "G" + _rowindex.ToString(), StyleIndex = (UInt32Value)51U, DataType = CellValues.String };
                     CellValue CommentValue = new CellValue();
                     CommentValue.Text = item.Comments;
                     cell54.Append(CommentValue);
                     Cell cell55 = new Cell() { CellReference = "H" + _rowindex.ToString(), StyleIndex = (UInt32Value)75U, DataType = CellValues.Number };
                     CellValue CostSavingsAmountValue = new CellValue();
                     CostSavingsAmountValue.Text = item.SavingsCost.Value.ToString();
                     cell55.Append(CostSavingsAmountValue);
                     Cell cell56 = new Cell() { CellReference = "I" + _rowindex.ToString(), StyleIndex = (UInt32Value)61U, DataType = CellValues.String };
                     CellFormula cellFormula4 = new CellFormula();
                     cellFormula4.Text = "IF(C" + _rowindex.ToString() + "=\"\",\"\",IF(C" + _rowindex.ToString() + "=\"One-Time Savings\",H" + _rowindex.ToString() + ",SUM(H" + _rowindex.ToString() + "*((F" + _rowindex.ToString() + "-E" + _rowindex.ToString() + ")/30.42))))";
                     CellValue cellValue14 = new CellValue();
                     cellValue14.Text = "";
                     cellFormula4.CalculateCell = true;
                     cell56.Append(cellFormula4);
                     cell56.Append(cellValue14);

                     row7.Append(cell49);
                     row7.Append(cell50);
                     row7.Append(cell51);
                     row7.Append(cell52);
                     row7.Append(cell53);
                     row7.Append(cell54);
                     row7.Append(cell55);
                     row7.Append(cell56);
                     sheetData1.Append(row7);
                 }
                  
                 if (_even > 2 && _even < 5)
                 {
          
                     Row row7 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };
                     row7.RowIndex = _rowindex;
                     Cell cell49 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)58U, DataType = CellValues.Date };
                     CellValue cellValue13 = new CellValue();
                     cellValue13.Text = item.SavingsDate.Value.ToShortDateString();

                     cell49.Append(cellValue13);
                     Cell cell50 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)6U, DataType = CellValues.String };
                     CellValue SavingsTypeValue = new CellValue();
                     SavingsTypeValue.Text = item.SavingsType.Value == 1 ? "One-Time Savings" : "Monthly";
                     cell50.Append(SavingsTypeValue);

                     Cell cell51 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)6U, DataType = CellValues.String };
                     CellValue MonthsValues = new CellValue();
                     MonthsValues.Text = item.Months.Value.ToString();
                     cell51.Append(MonthsValues);

                     Cell cell52 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)9U, DataType = CellValues.Date };
                     CellValue StartDateValue = new CellValue();
                     StartDateValue.Text = item.SavingsDate.Value.ToShortDateString();
                     cell52.Append(StartDateValue);

                     Cell cell53 = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)9U, DataType = CellValues.Date };
                     CellValue EndDateValue = new CellValue();
                     EndDateValue.Text = item.EndDate.Value.ToShortDateString();
                     cell53.Append(EndDateValue);

                     Cell cell54 = new Cell() { CellReference = "G" + _rowindex.ToString(), StyleIndex = (UInt32Value)8U, DataType = CellValues.String };
                     CellValue CommentValue = new CellValue();
                     CommentValue.Text = item.Comments;
                     cell54.Append(CommentValue);
                     Cell cell55 = new Cell() { CellReference = "H" + _rowindex.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.Number };
                     CellValue CostSavingsAmountValue = new CellValue();
                     CostSavingsAmountValue.Text = item.SavingsCost.Value.ToString();
                     cell55.Append(CostSavingsAmountValue);
                     Cell cell56 = new Cell() { CellReference = "I" + _rowindex.ToString(), StyleIndex = (UInt32Value)59U, DataType = CellValues.String };
                     CellFormula cellFormula4 = new CellFormula();
                     cellFormula4.Text = "IF(C" + _rowindex.ToString() + "=\"\",\"\",IF(C" + _rowindex.ToString() + "=\"One-Time Savings\",H" + _rowindex.ToString() + ",SUM(H" + _rowindex.ToString() + "*((F" + _rowindex.ToString() + "-E" + _rowindex.ToString() + ")/30.42))))";
                     CellValue cellValue14 = new CellValue();
                     cellValue14.Text = "";
                     cellFormula4.CalculateCell = true;
                     cell56.Append(cellFormula4);
                     cell56.Append(cellValue14);

                     row7.Append(cell49);
                     row7.Append(cell50);
                     row7.Append(cell51);
                     row7.Append(cell52);
                     row7.Append(cell53);
                     row7.Append(cell54);
                     row7.Append(cell55);
                     row7.Append(cell56);
                     sheetData1.Append(row7);
                    
                 }
                 if (_even == 4) _even = 1; else _even++;

             }
             _rowindex++;
             _even = 1;
             if (_rowindex < 31)
             {
                 for (var i = _rowindex; i < 31; i++)
                 {

                     if (_even < 3)
                     {
                         Row row37 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };
                         row37.RowIndex = i;
                         Cell cell1885 = new Cell() { CellReference = "B" + i.ToString(), StyleIndex = (UInt32Value)60U };                       
                         Cell cell251 = new Cell() { CellReference = "C" + i.ToString(), StyleIndex = (UInt32Value)54U };
                         Cell cell252 = new Cell() { CellReference = "D" + i.ToString(), StyleIndex = (UInt32Value)54U };
                         Cell cell253 = new Cell() { CellReference = "E" + i.ToString(), StyleIndex = (UInt32Value)54U };
                         Cell cell254 = new Cell() { CellReference = "F" + i.ToString(), StyleIndex = (UInt32Value)54U };
                         Cell cell255 = new Cell() { CellReference = "G" + i.ToString(), StyleIndex = (UInt32Value)54U };
                         Cell cell256 = new Cell() { CellReference = "H" + i.ToString(), StyleIndex = (UInt32Value)54U };
                         Cell cell257 = new Cell() { CellReference = "I" + i.ToString(), StyleIndex = (UInt32Value)61U };


                         row37.Append(cell1885);
                         row37.Append(cell251);
                         row37.Append(cell252);
                         row37.Append(cell253);
                         row37.Append(cell254);
                         row37.Append(cell255);
                         row37.Append(cell256);
                         row37.Append(cell257);
                         sheetData1.Append(row37);




                     }

                     if (_even >= 3 && _even < 5)
                     {
                         Row row39 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 14.25D, CustomHeight = true, DyDescent = 0.25D };
                         row39.RowIndex = i;
                         Cell cell263 = new Cell() { CellReference = "B" + i.ToString(), StyleIndex = (UInt32Value)58U };
                         Cell cell264 = new Cell() { CellReference = "C" + i.ToString(), StyleIndex = (UInt32Value)26U };
                         Cell cell265 = new Cell() { CellReference = "D" + i.ToString(), StyleIndex = (UInt32Value)21U };
                         Cell cell266 = new Cell() { CellReference = "E" + i.ToString(), StyleIndex = (UInt32Value)21U };
                         Cell cell267 = new Cell() { CellReference = "F" + i.ToString(), StyleIndex = (UInt32Value)8U };
                         Cell cell268 = new Cell() { CellReference = "G" + i.ToString(), StyleIndex = (UInt32Value)8U };
                         Cell cell269 = new Cell() { CellReference = "H" + i.ToString(), StyleIndex = (UInt32Value)8U };
                         Cell cell270 = new Cell() { CellReference = "I" + i.ToString(), StyleIndex = (UInt32Value)59U };
                        
                         CellValue cellValue69 = new CellValue();
                         cellValue69.Text = "";

                        
                         cell268.Append(cellValue69);

                         row39.Append(cell263);
                         row39.Append(cell264);
                         row39.Append(cell265);
                         row39.Append(cell266);
                         row39.Append(cell267);
                         row39.Append(cell268);
                         row39.Append(cell269);
                         row39.Append(cell270);
                         sheetData1.Append(row39);


                     }
                     if (_even == 4) _even = 1; else _even++;
                     _rowindex++;
                 }
             }
           
             Row row127 = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };
             Cell cell837 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)92U };
             Cell cell838 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)119U };
             Cell cell839 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
             Cell cell840 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
             Cell cell841 = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
             Cell cell842 = new Cell() { CellReference = "G" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
             Cell cell842a = new Cell() { CellReference = "H" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
             Cell cell843 = new Cell() { CellReference = "I" + _rowindex.ToString(), StyleIndex = (UInt32Value)121U };

             row127.Append(cell837);
             row127.Append(cell838);
             row127.Append(cell839);
             row127.Append(cell840);
             row127.Append(cell841);
             row127.Append(cell842);
             row127.Append(cell842a);
             row127.Append(cell843);
             sheetData1.Append(row127);
             _rowindex++;
             Row row127a = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };
             Cell cell837a = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)14U };
             Cell cell838a = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)14U };
             Cell cell839a = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)14U };
             Cell cell840a = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)14U };
             Cell cell841a = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)14U };
             Cell cell842b = new Cell() { CellReference = "G" + _rowindex.ToString(), StyleIndex = (UInt32Value)14U };
             Cell cell842c = new Cell() { CellReference = "H" + _rowindex.ToString(), StyleIndex = (UInt32Value)164U };
             CellValue TotalLabel = new CellValue();
             TotalLabel.Text = "Total Cost Avoidance:";
             cell842c.Append(TotalLabel);
             Cell cell843a = new Cell() { CellReference = "I" + _rowindex.ToString(), StyleIndex = (UInt32Value)167U };
             CellFormula CostAvoidanceTotal = new CellFormula();
             UInt32Value EndValue = _rowindex - 1;
             CostAvoidanceTotal.Text = "SUM(I7:I" + EndValue + ")";
             cell843a.Append(CostAvoidanceTotal);
             row127a.Append(cell837a);
             row127a.Append(cell838a);
             row127a.Append(cell839a);
             row127a.Append(cell840a);
             row127a.Append(cell841a);
             row127a.Append(cell842b);
             row127a.Append(cell842c);
             row127a.Append(cell843a);
             sheetData1.Append(row127a);
             _costavoidanceTotalString = "='Cost Avoidance'!$I$" + _rowindex;
             cellFormula1.CalculateCell = true;
            MergeCells mergeCells1 = new MergeCells() { Count = (UInt32Value)3U };
            MergeCell mergeCell1 = new MergeCell() { Reference = "A1:J1" };
            MergeCell mergeCell2 = new MergeCell() { Reference = "A2:J2" };
            MergeCell mergeCell3 = new MergeCell() { Reference = "B3:J3" };
            if (query.Count() == 0)
            {
                MergeCell mergeCell4 = new MergeCell() { Reference = "B7:C7" };            
                mergeCells1.Append(mergeCell4);
            }
            mergeCells1.Append(mergeCell1);
            mergeCells1.Append(mergeCell2);
            mergeCells1.Append(mergeCell3);
 
           
           

            DataValidations dataValidations1 = new DataValidations() { Count = (UInt32Value)1U };
            DataValidation dataValidation1 = new DataValidation() { AllowBlank = true, ShowInputMessage = true, SequenceOfReferences = new ListValue<StringValue>() { InnerText = "C7:D27" } };

            dataValidations1.Append(dataValidation1);
            PageMargins pageMargins1 = new PageMargins() { Left = 0.333D, Right = 0.333D, Top = 1.3958333333333333D, Bottom = 0.94791666666666663D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup1 = new PageSetup() { Orientation = OrientationValues.Landscape, Id = "rId1" };

            HeaderFooter headerFooter1 = new HeaderFooter();
            OddHeader oddHeader1 = new OddHeader();
            oddHeader1.Text = "&C&G";
            OddFooter oddFooter1 = new OddFooter();
            oddFooter1.Text = "&C&G";

            headerFooter1.Append(oddHeader1);
            headerFooter1.Append(oddFooter1);
            LegacyDrawing legacyDrawing1 = new LegacyDrawing() { Id = "rId2" };
            LegacyDrawingHeaderFooter legacyDrawingHeaderFooter1 = new LegacyDrawingHeaderFooter() { Id = "rId3" };

            worksheet1.Append(sheetProperties1);
            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(columns1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(mergeCells1);
           
            worksheet1.Append(dataValidations1);
            worksheet1.Append(pageMargins1);
            worksheet1.Append(pageSetup1);
            worksheet1.Append(headerFooter1);
            worksheet1.Append(legacyDrawing1);
            worksheet1.Append(legacyDrawingHeaderFooter1);

            worksheetPart1.Worksheet = worksheet1;
        }    
        // Replacements
        // Generates content of worksheetPart2.
        private void GenerateWorksheetPartReplacement(WorksheetPart worksheetPart2)
        {
            Worksheet worksheet2 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet2.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet2.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetProperties sheetProperties2 = new SheetProperties() { CodeName = "Sheet7" };
            SheetDimension sheetDimension2 = new SheetDimension() { Reference = "A1:H28" };

            SheetViews sheetViews2 = new SheetViews();

            SheetView sheetView2 = new SheetView() { ShowGridLines = false, ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U, View = SheetViewValues.PageLayout };
            Selection selection2 = new Selection() { ActiveCell = "B8", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "B8" } };

            sheetView2.Append(selection2);

            sheetViews2.Append(sheetView2);
            SheetFormatProperties sheetFormatProperties2 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            Columns columns2 = new Columns();
            Column column11 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)2U, Width = 9.140625D, Style = (UInt32Value)1U };
            Column column12 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 23.85546875D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column13 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 22.85546875D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column14 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 14D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column15 = new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)6U, Width = 27.85546875D, Style = (UInt32Value)5U, CustomWidth = true };
            Column column16 = new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 15.28515625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column17 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 7.5703125D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column18 = new Column() { Min = (UInt32Value)9U, Max = (UInt32Value)16384U, Width = 9.140625D, Style = (UInt32Value)1U };

            columns2.Append(column11);
            columns2.Append(column12);
            columns2.Append(column13);
            columns2.Append(column14);
            columns2.Append(column15);
            columns2.Append(column16);
            columns2.Append(column17);
            columns2.Append(column18);

            SheetData sheetData2 = new SheetData();

            Row row30 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 15.75D, DyDescent = 0.25D };

            Cell cell219 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)125U, DataType = CellValues.String };
            CellFormula cellFormula46 = new CellFormula();
            cellFormula46.Text = "SETUP!$B$2";
            CellValue cellValue57 = new CellValue();
            cellValue57.Text = "?";
            cellFormula46.CalculateCell = true;
            cell219.Append(cellFormula46);
            cell219.Append(cellValue57);
            Cell cell220 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)126U };
            Cell cell221 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)126U };
            Cell cell222 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)126U };
            Cell cell223 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)126U };
            Cell cell224 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)126U };
            Cell cell225 = new Cell() { CellReference = "G1", StyleIndex = (UInt32Value)126U };
            Cell cell226 = new Cell() { CellReference = "H1", StyleIndex = (UInt32Value)126U };

            row30.Append(cell219);
            row30.Append(cell220);
            row30.Append(cell221);
            row30.Append(cell222);
            row30.Append(cell223);
            row30.Append(cell224);
            row30.Append(cell225);
            row30.Append(cell226);

            Row row31 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };

            Cell cell227 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)123U, DataType = CellValues.String };
            CellFormula cellFormula47 = new CellFormula();
            cellFormula47.Text = "IF(B8=\"\",\"\",\"Asset Replacement Summary Through Period Ending \"&TEXT(SETUP!B4,\"MMMMMMMMM DD, YYYY\")&\"\")";
            CellValue cellValue58 = new CellValue();
            cellValue58.Text = "";
            cellFormula47.CalculateCell = true;
            cell227.Append(cellFormula47);
            cell227.Append(cellValue58);
            Cell cell228 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)124U };
            Cell cell229 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)124U };
            Cell cell230 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)124U };
            Cell cell231 = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)124U };
            Cell cell232 = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)124U };
            Cell cell233 = new Cell() { CellReference = "G2", StyleIndex = (UInt32Value)124U };
            Cell cell234 = new Cell() { CellReference = "H2", StyleIndex = (UInt32Value)124U };

            row31.Append(cell227);
            row31.Append(cell228);
            row31.Append(cell229);
            row31.Append(cell230);
            row31.Append(cell231);
            row31.Append(cell232);
            row31.Append(cell233);
            row31.Append(cell234);

            Row row32 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 6.75D, CustomHeight = true, DyDescent = 0.25D };
            Cell cell235 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)18U };

            row32.Append(cell235);

            Row row33 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 27D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell236 = new Cell() { CellReference = "B4", StyleIndex = (UInt32Value)199U, DataType = CellValues.String };
            CellFormula cellFormula48 = new CellFormula();
            cellFormula48.Text = "\"The summary below details the assets replaced by FPR with the respective asset replacement cost that would have been incurred by \"&SETUP!$B$2&\" to refresh the non-functioning asset.\"";
            CellValue cellValue59 = new CellValue();
            cellValue59.Text = "The summary below details the assets replaced by FPR with the respective asset replacement cost that would have been incurred by ? to refresh the non-functioning asset.";
            cellFormula48.CalculateCell = true;
            cell236.Append(cellFormula48);
            cell236.Append(cellValue59);
            Cell cell237 = new Cell() { CellReference = "C4", StyleIndex = (UInt32Value)128U };
            Cell cell238 = new Cell() { CellReference = "D4", StyleIndex = (UInt32Value)128U };
            Cell cell239 = new Cell() { CellReference = "E4", StyleIndex = (UInt32Value)128U };
            Cell cell240 = new Cell() { CellReference = "F4", StyleIndex = (UInt32Value)128U };
            Cell cell241 = new Cell() { CellReference = "G4", StyleIndex = (UInt32Value)128U };

            row33.Append(cell236);
            row33.Append(cell237);
            row33.Append(cell238);
            row33.Append(cell239);
            row33.Append(cell240);
            row33.Append(cell241);

            Row row34 = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 3.75D, CustomHeight = true, DyDescent = 0.25D };
            Cell cell242 = new Cell() { CellReference = "B5", StyleIndex = (UInt32Value)20U };
            Cell cell243 = new Cell() { CellReference = "C5", StyleIndex = (UInt32Value)18U };

            row34.Append(cell242);
            row34.Append(cell243);

            Row row35 = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 3D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell244 = new Cell() { CellReference = "G6", StyleIndex = (UInt32Value)10U };
            CellFormula cellFormula49 = new CellFormula();
            cellFormula49.Text = "SUM(G8:G30)";
            CellValue cellValue60 = new CellValue();
            cellValue60.Text = "0";
            cellFormula49.CalculateCell = true;
            cell244.Append(cellFormula49);
            cell244.Append(cellValue60);

            row35.Append(cell244);

            Row row36 = new Row() { RowIndex = (UInt32Value)7U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 28.5D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell245 = new Cell() { CellReference = "B7", StyleIndex = (UInt32Value)47U, DataType = CellValues.SharedString };
            CellValue cellValue61 = new CellValue();
            cellValue61.Text = "3";

            cell245.Append(cellValue61);

            Cell cell246 = new Cell() { CellReference = "C7", StyleIndex = (UInt32Value)45U, DataType = CellValues.SharedString };
            CellValue cellValue62 = new CellValue();
            cellValue62.Text = "40";

            cell246.Append(cellValue62);

            Cell cell247 = new Cell() { CellReference = "D7", StyleIndex = (UInt32Value)45U, DataType = CellValues.SharedString };
            CellValue cellValue63 = new CellValue();
            cellValue63.Text = "39";

            cell247.Append(cellValue63);

            Cell cell248 = new Cell() { CellReference = "E7", StyleIndex = (UInt32Value)45U, DataType = CellValues.SharedString };
            CellValue cellValue64 = new CellValue();
            cellValue64.Text = "41";

            cell248.Append(cellValue64);

            Cell cell249 = new Cell() { CellReference = "F7", StyleIndex = (UInt32Value)48U, DataType = CellValues.SharedString };
            CellValue cellValue65 = new CellValue();
            cellValue65.Text = "5";

            cell249.Append(cellValue65);

            Cell cell250 = new Cell() { CellReference = "G7", StyleIndex = (UInt32Value)46U, DataType = CellValues.SharedString };
            CellValue cellValue66 = new CellValue();
            cellValue66.Text = "4";

            cell250.Append(cellValue66);

            row36.Append(cell245);
            row36.Append(cell246);
            row36.Append(cell247);
            row36.Append(cell248);
            row36.Append(cell249);
            row36.Append(cell250);
/* Insert code here for loop use two rows to get alternating styles.
 *
 * 
 * 
 * 
 * 
 * 
 *  REPLACEMENT
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 */
            sheetData2.Append(row30);
            sheetData2.Append(row31);
            sheetData2.Append(row32);
            sheetData2.Append(row33);
            sheetData2.Append(row34);
            sheetData2.Append(row35);
            sheetData2.Append(row36);
           
            GlobalViewEntities db = new GlobalViewEntities();
            DateTime  overrideDate = Convert.ToDateTime(_overrideDate);
            var query = (from r in db.AssetReplacements
                         where r.CustomerID == _customerID && r.ReplacementDate >= overrideDate
                         orderby r.ReplacementDate descending
                         select r).ToList();
            UInt32Value _rowindex = 7;
             int _even = 1;
             if (query.Count() == 0)
             {
                 _rowindex = 8;
                 Row row37 = new Row() {  Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };
                 row37.RowIndex = _rowindex;
                 Cell cell1885 = new Cell() {  StyleIndex = (UInt32Value)216U, DataType = CellValues.String };
                 CellValue cellValue247 = new CellValue();
                 cellValue247.Text = "There are no replacements to report.";
                 cell1885.CellReference = "B" + _rowindex.ToString();
                 cell1885.Append(cellValue247);

                 Cell cell251 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)54U };
                 CellValue cellOldModel = new CellValue();
                 cellOldModel.Text = "";
                 cell251.Append(cellOldModel);

                 Cell cell252 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)54U };
                 CellValue cellNewModel = new CellValue();
                 cellNewModel.Text = "";
                 cell252.Append(cellNewModel);

                 Cell cell253 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)54U, DataType = CellValues.String };
                 CellValue cellNewSN = new CellValue();
                 cellNewSN.Text = "";
                 cell253.Append(cellNewSN);

                 Cell cell254 = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)51U };
                 CellValue cellreplacecomments = new CellValue();
                 cellreplacecomments.Text = "";
                 cell254.Append(cellreplacecomments);

                 Cell cell256 = new Cell() { CellReference = "G" + _rowindex.ToString(), StyleIndex = (UInt32Value)61U, DataType = CellValues.Number };
                 CellValue cellValue67 = new CellValue();
                 cellValue67.Text ="";
                 cell256.Append(cellValue67);

                 row37.Append(cell1885);
                 row37.Append(cell251);
                 row37.Append(cell252);
                 row37.Append(cell253);
                 row37.Append(cell254);
                 row37.Append(cell256);
                 sheetData2.Append(row37);


             }
            foreach (var item in query)
            {
                _rowindex++;
                 if (_even < 3)
                {
                    Row row37 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };
                    row37.RowIndex = _rowindex;
                    Cell cell1885 = new Cell() { StyleIndex = (UInt32Value)60U, DataType = CellValues.String };
                    CellValue cellValue247 = new CellValue();
                    cellValue247.Text = item.ReplacementDate.Value.ToShortDateString();
                    cell1885.CellReference = "B" + _rowindex.ToString();
                    cell1885.Append(cellValue247);

                    Cell cell251 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)54U };
                    CellValue cellOldModel = new CellValue();
                    cellOldModel.Text = item.OldModel;
                    cell251.Append(cellOldModel);

                    Cell cell252 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)54U };
                    CellValue cellNewModel = new CellValue();
                    cellNewModel.Text = item.NewModel;
                    cell252.Append(cellNewModel);

                    Cell cell253 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)54U, DataType = CellValues.String };
                    CellValue cellNewSN = new CellValue();
                    cellNewSN.Text = item.NewSerialNumber.ToUpper();
                    cell253.Append(cellNewSN);

                    Cell cell254 = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)51U };
                    CellValue cellreplacecomments = new CellValue();
                    cellreplacecomments.Text = item.Comments;
                    cell254.Append(cellreplacecomments);

                    Cell cell256 = new Cell() { CellReference = "G" + _rowindex.ToString(), StyleIndex = (UInt32Value)61U, DataType = CellValues.Number };
                    CellValue cellValue67 = new CellValue();
                    cellValue67.Text = item.ReplacementValue.Value.ToString();
                    cell256.Append(cellValue67);

                    row37.Append(cell1885);
                    row37.Append(cell251);
                    row37.Append(cell252);
                    row37.Append(cell253);
                    row37.Append(cell254);
                    row37.Append(cell256);
                    sheetData2.Append(row37);
                     
                }
                 if (_even > 2 && _even < 5)
                 {
                                          
                    Row row37 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };
                    row37.RowIndex = _rowindex;
                    Cell cell1885 = new Cell() { StyleIndex = (UInt32Value)58U, DataType = CellValues.String };
                    CellValue cellValue247 = new CellValue();
                    cellValue247.Text = item.ReplacementDate.Value.ToShortDateString();
                    cell1885.CellReference = "B" + _rowindex.ToString();
                    cell1885.Append(cellValue247);

                    Cell cell251 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)26U };
                    CellValue cellOldModel = new CellValue();
                    cellOldModel.Text = item.OldModel;
                    cell251.Append(cellOldModel);

                    Cell cell252 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)21U };
                    CellValue cellNewModel = new CellValue();
                    cellNewModel.Text = item.NewModel;
                    cell252.Append(cellNewModel);

                    Cell cell253 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)21U, DataType = CellValues.String };
                    CellValue cellNewSN = new CellValue();
                    cellNewSN.Text = item.NewSerialNumber.ToUpper();
                    cell253.Append(cellNewSN);

                    Cell cell254 = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)8U };
                    CellValue cellreplacecomments = new CellValue();
                    cellreplacecomments.Text = item.Comments;
                    cell254.Append(cellreplacecomments);

                    Cell cell256 = new Cell() { CellReference = "G" + _rowindex.ToString(), StyleIndex = (UInt32Value)59U, DataType = CellValues.Number };
                    CellValue cellValue67 = new CellValue();
                    cellValue67.Text = item.ReplacementValue.Value.ToString();
                    cell256.Append(cellValue67);

                    row37.Append(cell1885);
                    row37.Append(cell251);
                    row37.Append(cell252);
                    row37.Append(cell253);
                    row37.Append(cell254);
                    row37.Append(cell256);
                    sheetData2.Append(row37);
                     
                 }
                 if (_even == 4) _even = 1;
                 _even++;
               
            }
            _rowindex++;
         
            if (_rowindex < 31)
            {
                for (var i = _rowindex; i < 32 ; i++)
                {
                   
                    if (_even < 3)
                    {
                        Row row37 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };
                        row37.RowIndex = i;
                        Cell cell1885 = new Cell(){ CellReference = "B" + i.ToString(), StyleIndex = (UInt32Value)60U };
                        Cell cell251 = new Cell() { CellReference = "C" + i.ToString(), StyleIndex = (UInt32Value)54U };
                        Cell cell252 = new Cell() { CellReference = "D" + i.ToString(), StyleIndex = (UInt32Value)54U };
                        Cell cell253 = new Cell() { CellReference = "E" + i.ToString(), StyleIndex = (UInt32Value)54U };
                        Cell cell254 = new Cell() { CellReference = "F" + i.ToString(), StyleIndex = (UInt32Value)54U };
                        Cell cell256 = new Cell() { CellReference = "G" + i.ToString(), StyleIndex = (UInt32Value)61U };


                        row37.Append(cell1885);
                        row37.Append(cell251);
                        row37.Append(cell252);
                        row37.Append(cell253);
                        row37.Append(cell254);
                        row37.Append(cell256);
                        sheetData2.Append(row37);



                        
                    }

                    if (_even >=  3 && _even < 5)
                    {
                        Row row39 = new Row() {   Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 14.25D, CustomHeight = true, DyDescent = 0.25D };
                        row39.RowIndex = i;
                        Cell cell263 = new Cell() { CellReference = "B" + i.ToString(), StyleIndex = (UInt32Value)58U };
                        Cell cell264 = new Cell() { CellReference = "C" + i.ToString(), StyleIndex = (UInt32Value)26U };
                        Cell cell265 = new Cell() { CellReference = "D" + i.ToString(), StyleIndex = (UInt32Value)21U };
                        Cell cell266 = new Cell() { CellReference = "E" + i.ToString(), StyleIndex = (UInt32Value)21U };
                        Cell cell267 = new Cell() { CellReference = "F" + i.ToString(), StyleIndex = (UInt32Value)8U };

                        Cell cell268 = new Cell() { CellReference = "G" + i.ToString(), StyleIndex = (UInt32Value)59U, DataType = CellValues.String };
                        CellFormula cellFormula52 = new CellFormula();
                        cellFormula52.Text = "";
                        CellValue cellValue69 = new CellValue();
                        cellValue69.Text = "";

                        cell268.Append(cellFormula52);
                        cell268.Append(cellValue69);

                        row39.Append(cell263);
                        row39.Append(cell264);
                        row39.Append(cell265);
                        row39.Append(cell266);
                        row39.Append(cell267);
                        row39.Append(cell268);
                        sheetData2.Append(row39);
                        

                    }
                    if (_even == 4) _even = 1; else _even++;
                    _rowindex++;
                }
            }
           
            Row row127 = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };
            Cell cell837 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)92U };
            Cell cell838 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)119U };
            Cell cell839 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell840 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell841 = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell843 = new Cell() { CellReference = "G" + _rowindex.ToString(), StyleIndex = (UInt32Value)121U };
           
            row127.Append(cell837);
            row127.Append(cell838);
            row127.Append(cell839);
            row127.Append(cell840);
            row127.Append(cell841);
            row127.Append(cell843);
            sheetData2.Append(row127);
            _rowindex++;
            Row row127a = new Row() { RowIndex = (UInt32Value) _rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };
            Cell cell837a = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)164U };
            Cell cell838a = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)164U };
            Cell cell839a = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)164U };
            Cell cell840a = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)164U };
            Cell cell841a = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)164U };
            CellValue TotalLabel = new CellValue();
            TotalLabel.Text = "Total Replacement Value:";
            cell841a.Append(TotalLabel);
            Cell cell843a = new Cell() { CellReference = "G" + _rowindex.ToString(), StyleIndex = (UInt32Value)167U };
            CellFormula cellFormula52a = new CellFormula();
            UInt32Value  EndRow = _rowindex - 1;
            cellFormula52a.Text = "=Sum(G8:G"+EndRow+")";
            cell843a.Append(cellFormula52a);
            row127a.Append(cell837a);
            row127a.Append(cell838a);
            row127a.Append(cell839a);
            row127a.Append(cell840a);
            row127a.Append(cell841a);            
            row127a.Append(cell843a);
            sheetData2.Append(row127a);
            _replacementTotalsString = "='Replacements'!$G$" + _rowindex.ToString();
            MergeCells mergeCells2 = new MergeCells() { Count = (UInt32Value)3U };
            MergeCell mergeCell4 = new MergeCell() { Reference = "A2:H2" };
            MergeCell mergeCell5 = new MergeCell() { Reference = "A1:H1" };
            MergeCell mergeCell6 = new MergeCell() { Reference = "B4:G4" };

            mergeCells2.Append(mergeCell4);
            mergeCells2.Append(mergeCell5);
            mergeCells2.Append(mergeCell6);

            DataValidations dataValidations2 = new DataValidations() { Count = (UInt32Value)1U };
            DataValidation dataValidation2 = new DataValidation() { AllowBlank = true, ShowInputMessage = true, SequenceOfReferences = new ListValue<StringValue>() { InnerText = "D8:D30" } };

            dataValidations2.Append(dataValidation2);
            PageMargins pageMargins2 = new PageMargins() { Left = 0.333D, Right = 0.333D, Top = 1.3333333333333333D, Bottom = 0.9375D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup2 = new PageSetup() { Orientation = OrientationValues.Landscape, Id = "rId1" };

            HeaderFooter headerFooter2 = new HeaderFooter();
            OddHeader oddHeader2 = new OddHeader();
            oddHeader2.Text = "&C&G";
            OddFooter oddFooter2 = new OddFooter();
            oddFooter2.Text = "&C&G";

            headerFooter2.Append(oddHeader2);
            headerFooter2.Append(oddFooter2);
            LegacyDrawingHeaderFooter legacyDrawingHeaderFooter2 = new LegacyDrawingHeaderFooter() { Id = "rId2" };

            worksheet2.Append(sheetProperties2);
            worksheet2.Append(sheetDimension2);
            worksheet2.Append(sheetViews2);
            worksheet2.Append(sheetFormatProperties2);
            worksheet2.Append(columns2);
            worksheet2.Append(sheetData2);
            worksheet2.Append(mergeCells2);
            worksheet2.Append(dataValidations2);
            worksheet2.Append(pageMargins2);
            worksheet2.Append(pageSetup2);
            worksheet2.Append(headerFooter2);
            worksheet2.Append(legacyDrawingHeaderFooter2);

            worksheetPart2.Worksheet = worksheet2;
        }      
        // SETUP PAGE - Completed Formatting  27-11-2012
        // Generates content of the setup page.
        // 
        private void GenerateWorksheetPartSetupPage(WorksheetPart worksheetPart3)
        {
            Worksheet worksheet3 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet3.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet3.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetProperties sheetProperties3 = new SheetProperties() { CodeName = "Sheet2" };
            SheetDimension sheetDimension3 = new SheetDimension() { Reference = "A1:C25" };

            SheetViews sheetViews3 = new SheetViews();

            SheetView sheetView3 = new SheetView() { ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U, View = SheetViewValues.PageLayout };
            Selection selection3 = new Selection() { ActiveCell = "B2", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "B2" } };

            sheetView3.Append(selection3);

            sheetViews3.Append(sheetView3);
            SheetFormatProperties sheetFormatProperties3 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            Columns columns3 = new Columns();
            Column column19 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 33D, Style = (UInt32Value)3U, CustomWidth = true };
            Column column20 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 48.85546875D, Style = (UInt32Value)5U, CustomWidth = true };
            Column column21 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)20U, Width = 9.140625D, Style = (UInt32Value)1U };
            Column column22 = new Column() { Min = (UInt32Value)21U, Max = (UInt32Value)21U, Width = 0D, Style = (UInt32Value)1U, Hidden = true, CustomWidth = true };
            Column column23 = new Column() { Min = (UInt32Value)22U, Max = (UInt32Value)16384U, Width = 9.140625D, Style = (UInt32Value)1U };

            columns3.Append(column19);
            columns3.Append(column20);
            columns3.Append(column21);
            columns3.Append(column22);
            columns3.Append(column23);

            SheetData sheetData3 = new SheetData();

            Row row60 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, Height = 15.75D, DyDescent = 0.25D };

            Cell cell389 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue90 = new CellValue();
            cellValue90.Text = "8";

            cell389.Append(cellValue90);

            row60.Append(cell389);

            Row row61 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.25D };

            Cell cell390 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue91 = new CellValue();
            cellValue91.Text = "6";

            cell390.Append(cellValue91);
            CoFreedomEntities db2 = new CoFreedomEntities();
            var query = (from c in db2.vw_CustomersOnContract
                         where c.CustomerID == _customerID
                         select c).FirstOrDefault();
            Cell cell391 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)103U, DataType = CellValues.SharedString };
            CellValue cellValue92 = new CellValue();
            cellValue92.Text = query.CustomerName;

            cell391.Append(cellValue92);

            row61.Append(cell390);
            row61.Append(cell391);

            Row row62 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.25D };

            Cell cell392 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue93 = new CellValue();
            cellValue93.Text = "9";

            cell392.Append(cellValue93);
            //var contract = (from cs in db2.SCContracts
            //                where cs.CustomerID == _customerID && cs.Active == true
            //                select cs).FirstOrDefault();
            Cell cell393 = new Cell() { CellReference = "B3", StyleIndex = (UInt32Value)103U, DataType = CellValues.SharedString };
            CellValue cellValue94 = new CellValue();
            cellValue94.Text = _overrideDate;// contract.StartDate.ToShortDateString();

            cell393.Append(cellValue94);

            row62.Append(cell392);
            row62.Append(cell393);

            Row row63 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.25D };

            Cell cell394 = new Cell() { CellReference = "A4", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue95 = new CellValue();
            cellValue95.Text = "7";

            cell394.Append(cellValue95);

            Cell cell395 = new Cell() { CellReference = "B4", StyleIndex = (UInt32Value)103U, DataType = CellValues.SharedString };
            CellValue cellValue96 = new CellValue();
            cellValue96.Text = _period;

            cell395.Append(cellValue96);

            row63.Append(cell394);
            row63.Append(cell395);

            Row row64 = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.25D };
            Cell cell396 = new Cell() { CellReference = "A6", StyleIndex = (UInt32Value)4U };

            row64.Append(cell396);

            Row row65 = new Row() { RowIndex = (UInt32Value)8U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, Height = 15.75D, ThickBot = true, DyDescent = 0.3D };

            Cell cell397 = new Cell() { CellReference = "A8", StyleIndex = (UInt32Value)22U, DataType = CellValues.SharedString };
            CellValue cellValue97 = new CellValue();
            cellValue97.Text = "27";

            cell397.Append(cellValue97);

            row65.Append(cell397);

            Row row66 = new Row() { RowIndex = (UInt32Value)9U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, Height = 22.5D, CustomHeight = true, ThickBot = true, DyDescent = 0.3D };

            Cell cell398 = new Cell() { CellReference = "A9", StyleIndex = (UInt32Value)101U, DataType = CellValues.SharedString };
            CellValue cellValue98 = new CellValue();
            cellValue98.Text = "0";

            cell398.Append(cellValue98);

            Cell cell399 = new Cell() { CellReference = "B9", StyleIndex = (UInt32Value)102U, DataType = CellValues.SharedString };
            CellValue cellValue99 = new CellValue();
            cellValue99.Text = "28";

            cell399.Append(cellValue99);

            Cell cell400 = new Cell() { CellReference = "C9", StyleIndex = (UInt32Value)122U, DataType = CellValues.String };
            CellFormula cellFormula73 = new CellFormula();
            cellFormula73.Text = "";
            CellValue cellValue100 = new CellValue();
            cellValue100.Text = "";

            // cell400.Append(cellFormula73);
            cell400.Append(cellValue100);

            row66.Append(cell398);
            row66.Append(cell399);
            row66.Append(cell400);



            sheetData3.Append(row60);
            sheetData3.Append(row61);
            sheetData3.Append(row62);
            sheetData3.Append(row63);
            sheetData3.Append(row64);
            sheetData3.Append(row65);
            sheetData3.Append(row66);

            GlobalViewEntities db3 = new GlobalViewEntities();
            var metergroups = (from pf in db3.RevisionMeterGroups
                               where pf.ERPContractID.Value == _contractID
                               select new
                               {
                                   MeterGroup = pf.ERPMeterGroupDesc,
                                   ClientCPP = pf.CPP
                               }).Distinct().ToList();
            UInt32Value rowindex = 9;
            int _even = 1;
            foreach (var row in metergroups)
            {
                rowindex++;
                if (_even == 1)
                {
                    Row row67 = new Row() { RowIndex = (UInt32Value)rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.25D };
                    Cell cell401 = new Cell() { CellReference = "A" + rowindex.ToString(), StyleIndex = (UInt32Value)106U, DataType = CellValues.String };
                    CellValue meterGroupValue = new CellValue();
                    meterGroupValue.Text = row.MeterGroup;
                    cell401.Append(meterGroupValue);
                    Cell cell402 = new Cell() { CellReference = "B" + rowindex.ToString(), StyleIndex = (UInt32Value)197U, DataType = CellValues.Number };
                    CellValue meterGroupCPPValue = new CellValue();
                    meterGroupCPPValue.Text = row.ClientCPP.ToString();
                    cell402.Append(meterGroupCPPValue);
                    row67.Append(cell401);
                    row67.Append(cell402);
                    sheetData3.Append(row67);
                    _even = 0;
                }
                else
                {
                    Row row67 = new Row() { RowIndex = (UInt32Value)rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.25D };
                    Cell cell401 = new Cell() { CellReference = "A" + rowindex.ToString(), StyleIndex = (UInt32Value)43U, DataType = CellValues.String };
                    CellValue meterGroupValue = new CellValue();
                    meterGroupValue.Text = row.MeterGroup;
                    cell401.Append(meterGroupValue);
                    Cell cell402 = new Cell() { CellReference = "B" + rowindex.ToString(), StyleIndex = (UInt32Value)198U, DataType = CellValues.Number };
                    CellValue meterGroupCPPValue = new CellValue();
                    meterGroupCPPValue.Text = row.ClientCPP.ToString();
                    cell402.Append(meterGroupCPPValue);
                    row67.Append(cell401);
                    row67.Append(cell402);
                    sheetData3.Append(row67);
                    _even = 1;
                }
            }


            rowindex++;
            UInt32Value looptimes = 0;
            if (rowindex < 25)
            {
                looptimes = 25 - rowindex;
                for (int i = 1; i <= looptimes; i++)
                {
                    if (_even == 1)
                    {
                        Row row67 = new Row() { RowIndex = (UInt32Value)rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.25D };
                        Cell cell401 = new Cell() { CellReference = "A" + rowindex.ToString(), StyleIndex = (UInt32Value)106U, DataType = CellValues.String };
                        CellValue meterGroupValue = new CellValue();
                        meterGroupValue.Text = "";
                        cell401.Append(meterGroupValue);
                        Cell cell402 = new Cell() { CellReference = "B" + rowindex.ToString(), StyleIndex = (UInt32Value)197U, DataType = CellValues.String };
                        CellValue meterGroupCPPValue = new CellValue();
                        meterGroupCPPValue.Text = " ";
                        cell402.Append(meterGroupCPPValue);
                        row67.Append(cell401);
                        row67.Append(cell402);
                        sheetData3.Append(row67);
                        _even = 0;
                    }
                    else
                    {
                        Row row67 = new Row() { RowIndex = (UInt32Value)rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.25D };
                        Cell cell401 = new Cell() { CellReference = "A" + rowindex.ToString(), StyleIndex = (UInt32Value)43U, DataType = CellValues.String };
                        CellValue meterGroupValue = new CellValue();
                        meterGroupValue.Text = "";
                        cell401.Append(meterGroupValue);
                        Cell cell402 = new Cell() { CellReference = "B" + rowindex.ToString(), StyleIndex = (UInt32Value)198U, DataType = CellValues.String };
                        CellValue meterGroupCPPValue = new CellValue();
                        meterGroupCPPValue.Text = " ";
                        cell402.Append(meterGroupCPPValue);
                        row67.Append(cell401);
                        row67.Append(cell402);
                        sheetData3.Append(row67);
                        _even = 1;
                    }
                    rowindex++;
                }
            }

            Row row82 = new Row() { RowIndex = (UInt32Value)rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:2" }, Height = 15.75D, ThickBot = true, DyDescent = 0.3D };
            Cell cell431 = new Cell() { CellReference = "A" + rowindex.ToString(), StyleIndex = (UInt32Value)108U };
            Cell cell432 = new Cell() { CellReference = "B" + rowindex.ToString(), StyleIndex = (UInt32Value)109U };

            row82.Append(cell431);
            row82.Append(cell432);


            sheetData3.Append(row82);
            PageMargins pageMargins3 = new PageMargins() { Left = 0.25D, Right = 0.25D, Top = 1.2916666666666667D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup3 = new PageSetup() { Orientation = OrientationValues.Landscape, Id = "rId1" };

            HeaderFooter headerFooter3 = new HeaderFooter();
            OddHeader oddHeader3 = new OddHeader();
            oddHeader3.Text = "&C&G";
            OddFooter oddFooter3 = new OddFooter();
            oddFooter3.Text = "&C&G";

            headerFooter3.Append(oddHeader3);
            headerFooter3.Append(oddFooter3);
            LegacyDrawingHeaderFooter legacyDrawingHeaderFooter3 = new LegacyDrawingHeaderFooter() { Id = "rId2" };

            worksheet3.Append(sheetProperties3);
            worksheet3.Append(sheetDimension3);
            worksheet3.Append(sheetViews3);
            worksheet3.Append(sheetFormatProperties3);
            worksheet3.Append(columns3);
            worksheet3.Append(sheetData3);
            worksheet3.Append(pageMargins3);
            worksheet3.Append(pageSetup3);
            worksheet3.Append(headerFooter3);
            worksheet3.Append(legacyDrawingHeaderFooter3);

            worksheetPart3.Worksheet = worksheet3;
        }
        //Survey Detail
        // Generates content of worksheetPart4.
        private void GenerateWorksheetPartSurveyDetail(WorksheetPart worksheetPart4)
        {
            Worksheet worksheet4 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet4.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet4.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet4.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetProperties sheetProperties4 = new SheetProperties() { CodeName = "Sheet12" };
            SheetDimension sheetDimension4 = new SheetDimension() { Reference = "A1:N21" };

            SheetViews sheetViews4 = new SheetViews();

            SheetView sheetView4 = new SheetView() { TopLeftCell = "A4", ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U, View = SheetViewValues.PageLayout };
            Selection selection4 = new Selection() { ActiveCell = "D4", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "D4" } };

            sheetView4.Append(selection4);

            sheetViews4.Append(sheetView4);
            SheetFormatProperties sheetFormatProperties4 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            Columns columns4 = new Columns();
            Column column24 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)3U, Width = 9.140625D, Style = (UInt32Value)1U };
            Column column25 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 15D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column26 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)14U, Width = 7.7109375D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column27 = new Column() { Min = (UInt32Value)15U, Max = (UInt32Value)16384U, Width = 9.140625D, Style = (UInt32Value)1U };

            columns4.Append(column24);
            columns4.Append(column25);
            columns4.Append(column26);
            columns4.Append(column27);

            SheetData sheetData4 = new SheetData();

            Row row83 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:14" }, Height = 18D, DyDescent = 0.25D };

            Cell cell433 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)141U, DataType = CellValues.String };
            CellFormula cellFormula74 = new CellFormula();
            cellFormula74.Text = "SETUP!B2";
            CellValue cellValue101 = new CellValue();
            cellValue101.Text = "?";
            cellFormula74.CalculateCell = true;
            cell433.Append(cellFormula74);
            cell433.Append(cellValue101);
            Cell cell434 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)142U };
            Cell cell435 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)142U };
            Cell cell436 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)142U };
            Cell cell437 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)142U };
            Cell cell438 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)142U };
            Cell cell439 = new Cell() { CellReference = "G1", StyleIndex = (UInt32Value)142U };
            Cell cell440 = new Cell() { CellReference = "H1", StyleIndex = (UInt32Value)142U };
            Cell cell441 = new Cell() { CellReference = "I1", StyleIndex = (UInt32Value)142U };
            Cell cell442 = new Cell() { CellReference = "J1", StyleIndex = (UInt32Value)142U };
            Cell cell443 = new Cell() { CellReference = "K1", StyleIndex = (UInt32Value)142U };
            Cell cell444 = new Cell() { CellReference = "L1", StyleIndex = (UInt32Value)142U };
            Cell cell445 = new Cell() { CellReference = "M1", StyleIndex = (UInt32Value)142U };
            Cell cell446 = new Cell() { CellReference = "N1", StyleIndex = (UInt32Value)142U };

            row83.Append(cell433);
            row83.Append(cell434);
            row83.Append(cell435);
            row83.Append(cell436);
            row83.Append(cell437);
            row83.Append(cell438);
            row83.Append(cell439);
            row83.Append(cell440);
            row83.Append(cell441);
            row83.Append(cell442);
            row83.Append(cell443);
            row83.Append(cell444);
            row83.Append(cell445);
            row83.Append(cell446);

            Row row84 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:14" }, DyDescent = 0.25D };

            Cell cell447 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)123U, DataType = CellValues.String };
            CellFormula cellFormula75 = new CellFormula();
            cellFormula75.Text = "\"FPR Service Support Survey for Period Ending  \"&TEXT(SETUP!$B$4,\"MMMMMMMM DD, YYYY\")&\"\"";
            CellValue cellValue102 = new CellValue();
            cellValue102.Text = "FPR Service Support Survey for Period Ending  ?";
            cellFormula75.CalculateCell = true;
            cell447.Append(cellFormula75);
            cell447.Append(cellValue102);
            Cell cell448 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)144U };
            Cell cell449 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)144U };
            Cell cell450 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)144U };
            Cell cell451 = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)144U };
            Cell cell452 = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)144U };
            Cell cell453 = new Cell() { CellReference = "G2", StyleIndex = (UInt32Value)144U };
            Cell cell454 = new Cell() { CellReference = "H2", StyleIndex = (UInt32Value)144U };
            Cell cell455 = new Cell() { CellReference = "I2", StyleIndex = (UInt32Value)144U };
            Cell cell456 = new Cell() { CellReference = "J2", StyleIndex = (UInt32Value)144U };
            Cell cell457 = new Cell() { CellReference = "K2", StyleIndex = (UInt32Value)144U };
            Cell cell458 = new Cell() { CellReference = "L2", StyleIndex = (UInt32Value)144U };
            Cell cell459 = new Cell() { CellReference = "M2", StyleIndex = (UInt32Value)144U };
            Cell cell460 = new Cell() { CellReference = "N2", StyleIndex = (UInt32Value)144U };

            row84.Append(cell447);
            row84.Append(cell448);
            row84.Append(cell449);
            row84.Append(cell450);
            row84.Append(cell451);
            row84.Append(cell452);
            row84.Append(cell453);
            row84.Append(cell454);
            row84.Append(cell455);
            row84.Append(cell456);
            row84.Append(cell457);
            row84.Append(cell458);
            row84.Append(cell459);
            row84.Append(cell460);
            
            var FromDate = Convert.ToDateTime(_period);
            var ToDate = Convert.ToDateTime(_period).AddMonths(-2);

            //CustomerPortalEntities db = new CustomerPortalEntities();
            //var query = (from r in db.vw_SurveySummaryList
            //             where r.CustomerID == _customerID && (r.SurveyDate.Value.Month <= FromDate.Month &&r.SurveyDate.Value.Year <= FromDate.Year) && (r.SurveyDate.Value.Month >= ToDate.Month && r.SurveyDate.Value.Year >= ToDate.Year)
            //             orderby r.SurveyDate descending
            //             select r).FirstOrDefault();


          //  if (query == null) return; 

            Row row85 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:14" }, Height = 18.75D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell461 = new Cell() { CellReference = "C4", StyleIndex = (UInt32Value)29U, DataType = CellValues.SharedString };
            CellValue cellValue103 = new CellValue();
            cellValue103.Text = "30";

            cell461.Append(cellValue103);
            Cell cell462 = new Cell() { CellReference = "D4", StyleIndex = (UInt32Value)27U };
            CellValue Cell462Value = new CellValue();
            Cell462Value.Text = "";//query == null ? "" : query.Name;
            cell462.Append(Cell462Value);

            Cell cell463 = new Cell() { CellReference = "E4", StyleIndex = (UInt32Value)28U };
            Cell cell464 = new Cell() { CellReference = "F4", StyleIndex = (UInt32Value)28U };
            Cell cell465 = new Cell() { CellReference = "G4", StyleIndex = (UInt32Value)28U };
            Cell cell466 = new Cell() { CellReference = "H4", StyleIndex = (UInt32Value)28U };
            Cell cell467 = new Cell() { CellReference = "I4", StyleIndex = (UInt32Value)28U };
            Cell cell468 = new Cell() { CellReference = "J4", StyleIndex = (UInt32Value)28U };
            Cell cell469 = new Cell() { CellReference = "K4", StyleIndex = (UInt32Value)28U };
            Cell cell470 = new Cell() { CellReference = "L4", StyleIndex = (UInt32Value)2U };

            row85.Append(cell461);
            row85.Append(cell462);
            row85.Append(cell463);
            row85.Append(cell464);
            row85.Append(cell465);
            row85.Append(cell466);
            row85.Append(cell467);
            row85.Append(cell468);
            row85.Append(cell469);
            row85.Append(cell470);

            Row row86 = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:14" }, Height = 18.75D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell471 = new Cell() { CellReference = "C5", StyleIndex = (UInt32Value)29U, DataType = CellValues.SharedString };
            CellValue cellValue104 = new CellValue();
            cellValue104.Text = "31";

            cell471.Append(cellValue104);
            Cell cell472 = new Cell() { CellReference = "D5", StyleIndex = (UInt32Value)27U };
            CellValue Cell472Value = new CellValue();
            Cell472Value.Text = ""; // query == null ? "" : query.Email;
            cell472.Append(Cell472Value);
            Cell cell473 = new Cell() { CellReference = "E5", StyleIndex = (UInt32Value)28U };
            Cell cell474 = new Cell() { CellReference = "F5", StyleIndex = (UInt32Value)28U };
            Cell cell475 = new Cell() { CellReference = "G5", StyleIndex = (UInt32Value)28U };
            Cell cell476 = new Cell() { CellReference = "H5", StyleIndex = (UInt32Value)28U };
            Cell cell477 = new Cell() { CellReference = "I5", StyleIndex = (UInt32Value)28U };
            Cell cell478 = new Cell() { CellReference = "J5", StyleIndex = (UInt32Value)28U };
            Cell cell479 = new Cell() { CellReference = "K5", StyleIndex = (UInt32Value)28U };
            Cell cell480 = new Cell() { CellReference = "L5", StyleIndex = (UInt32Value)2U };

            row86.Append(cell471);
            row86.Append(cell472);
            row86.Append(cell473);
            row86.Append(cell474);
            row86.Append(cell475);
            row86.Append(cell476);
            row86.Append(cell477);
            row86.Append(cell478);
            row86.Append(cell479);
            row86.Append(cell480);

            Row row87 = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:14" }, Height = 18.75D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell481 = new Cell() { CellReference = "C6", StyleIndex = (UInt32Value)29U, DataType = CellValues.SharedString };
            CellValue cellValue105 = new CellValue();
            cellValue105.Text = "32";

            cell481.Append(cellValue105);
            Cell cell482 = new Cell() { CellReference = "D6", StyleIndex = (UInt32Value)27U };
            CellValue Cell482Value = new CellValue();
            Cell482Value.Text = "";// query == null ? "" : query.SurveyDate.Value.ToShortDateString();
            cell482.Append(Cell482Value);
            Cell cell483 = new Cell() { CellReference = "E6", StyleIndex = (UInt32Value)28U };
            Cell cell484 = new Cell() { CellReference = "F6", StyleIndex = (UInt32Value)28U };
            Cell cell485 = new Cell() { CellReference = "G6", StyleIndex = (UInt32Value)28U };
            Cell cell486 = new Cell() { CellReference = "H6", StyleIndex = (UInt32Value)28U };
            Cell cell487 = new Cell() { CellReference = "I6", StyleIndex = (UInt32Value)28U };
            Cell cell488 = new Cell() { CellReference = "J6", StyleIndex = (UInt32Value)28U };
            Cell cell489 = new Cell() { CellReference = "K6", StyleIndex = (UInt32Value)28U };
            Cell cell490 = new Cell() { CellReference = "L6", StyleIndex = (UInt32Value)2U };

            row87.Append(cell481);
            row87.Append(cell482);
            row87.Append(cell483);
            row87.Append(cell484);
            row87.Append(cell485);
            row87.Append(cell486);
            row87.Append(cell487);
            row87.Append(cell488);
            row87.Append(cell489);
            row87.Append(cell490);

            Row row88 = new Row() { RowIndex = (UInt32Value)7U, Spans = new ListValue<StringValue>() { InnerText = "1:14" }, Height = 15.75D, ThickBot = true, DyDescent = 0.3D };
            Cell cell491 = new Cell() { CellReference = "B7", StyleIndex = (UInt32Value)151U };
            Cell cell492 = new Cell() { CellReference = "C7", StyleIndex = (UInt32Value)152U };
            Cell cell493 = new Cell() { CellReference = "D7", StyleIndex = (UInt32Value)152U };
            Cell cell494 = new Cell() { CellReference = "E7", StyleIndex = (UInt32Value)152U };
            Cell cell495 = new Cell() { CellReference = "F7", StyleIndex = (UInt32Value)152U };
            Cell cell496 = new Cell() { CellReference = "G7", StyleIndex = (UInt32Value)152U };
            Cell cell497 = new Cell() { CellReference = "H7", StyleIndex = (UInt32Value)152U };
            Cell cell498 = new Cell() { CellReference = "I7", StyleIndex = (UInt32Value)152U };
            Cell cell499 = new Cell() { CellReference = "J7", StyleIndex = (UInt32Value)152U };
            Cell cell500 = new Cell() { CellReference = "K7", StyleIndex = (UInt32Value)152U };
            Cell cell501 = new Cell() { CellReference = "L7", StyleIndex = (UInt32Value)152U };
            Cell cell502 = new Cell() { CellReference = "M7", StyleIndex = (UInt32Value)152U };
            Cell cell503 = new Cell() { CellReference = "N7", StyleIndex = (UInt32Value)152U };

            row88.Append(cell491);
            row88.Append(cell492);
            row88.Append(cell493);
            row88.Append(cell494);
            row88.Append(cell495);
            row88.Append(cell496);
            row88.Append(cell497);
            row88.Append(cell498);
            row88.Append(cell499);
            row88.Append(cell500);
            row88.Append(cell501);
            row88.Append(cell502);
            row88.Append(cell503);

            Row row89 = new Row() { RowIndex = (UInt32Value)8U, Spans = new ListValue<StringValue>() { InnerText = "1:14" }, Height = 15.75D, ThickBot = true, DyDescent = 0.3D };

            Cell cell504 = new Cell() { CellReference = "B8", StyleIndex = (UInt32Value)145U, DataType = CellValues.SharedString };
            CellValue cellValue106 = new CellValue();
            cellValue106.Text = "33";

            cell504.Append(cellValue106);
            Cell cell505 = new Cell() { CellReference = "C8", StyleIndex = (UInt32Value)146U };
            Cell cell506 = new Cell() { CellReference = "D8", StyleIndex = (UInt32Value)147U };

            Cell cell507 = new Cell() { CellReference = "E8", StyleIndex = (UInt32Value)99U };
            CellValue cellValue107 = new CellValue();
            cellValue107.Text = "1";

            cell507.Append(cellValue107);

            Cell cell508 = new Cell() { CellReference = "F8", StyleIndex = (UInt32Value)99U };
            CellValue cellValue108 = new CellValue();
            cellValue108.Text = "2";

            cell508.Append(cellValue108);

            Cell cell509 = new Cell() { CellReference = "G8", StyleIndex = (UInt32Value)99U };
            CellValue cellValue109 = new CellValue();
            cellValue109.Text = "3";

            cell509.Append(cellValue109);

            Cell cell510 = new Cell() { CellReference = "H8", StyleIndex = (UInt32Value)99U };
            CellValue cellValue110 = new CellValue();
            cellValue110.Text = "4";

            cell510.Append(cellValue110);

            Cell cell511 = new Cell() { CellReference = "I8", StyleIndex = (UInt32Value)99U };
            CellValue cellValue111 = new CellValue();
            cellValue111.Text = "5";

            cell511.Append(cellValue111);

            Cell cell512 = new Cell() { CellReference = "J8", StyleIndex = (UInt32Value)99U };
            CellValue cellValue112 = new CellValue();
            cellValue112.Text = "6";

            cell512.Append(cellValue112);

            Cell cell513 = new Cell() { CellReference = "K8", StyleIndex = (UInt32Value)99U };
            CellValue cellValue113 = new CellValue();
            cellValue113.Text = "7";

            cell513.Append(cellValue113);

            Cell cell514 = new Cell() { CellReference = "L8", StyleIndex = (UInt32Value)99U };
            CellValue cellValue114 = new CellValue();
            cellValue114.Text = "8";

            cell514.Append(cellValue114);

            Cell cell515 = new Cell() { CellReference = "M8", StyleIndex = (UInt32Value)99U };
            CellValue cellValue115 = new CellValue();
            cellValue115.Text = "9";

            cell515.Append(cellValue115);

            Cell cell516 = new Cell() { CellReference = "N8", StyleIndex = (UInt32Value)100U };
            CellValue cellValue116 = new CellValue();
            cellValue116.Text = "10";

            cell516.Append(cellValue116);

            row89.Append(cell504);
            row89.Append(cell505);
            row89.Append(cell506);
            row89.Append(cell507);
            row89.Append(cell508);
            row89.Append(cell509);
            row89.Append(cell510);
            row89.Append(cell511);
            row89.Append(cell512);
            row89.Append(cell513);
            row89.Append(cell514);
            row89.Append(cell515);
            row89.Append(cell516);

            Row row90 = new Row() { RowIndex = (UInt32Value)9U, Spans = new ListValue<StringValue>() { InnerText = "1:14" }, Height = 22.5D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell517 = new Cell() { CellReference = "B9", StyleIndex = (UInt32Value)148U, DataType = CellValues.SharedString };
            CellValue cellValue117 = new CellValue();
            cellValue117.Text = "13";

            cell517.Append(cellValue117);
                  
            Cell cell518 = new Cell() { CellReference = "C9",  StyleIndex = (UInt32Value)149U };
            Cell cell519 = new Cell() { CellReference = "D9", StyleIndex = (UInt32Value)150U };
            CellValue Ans1 = new CellValue();
            Ans1.Text = "";// query == null ? "" : query.Answer1 == 1 ? "*" : "";
            Cell cell520 = new Cell() { CellReference = "E9", CellValue = Ans1, StyleIndex = (UInt32Value)114U };
            CellValue Ans2 = new CellValue();
            Ans2.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer1.Value) == 2 ? "*" : "";
            Cell cell521 = new Cell() { CellReference = "F9", CellValue = Ans2, StyleIndex = (UInt32Value)114U };
            CellValue Ans3 = new CellValue();
            Ans3.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer1.Value) == 3 ? "*" : "";
            Cell cell522 = new Cell() { CellReference = "G9", CellValue = Ans3, StyleIndex = (UInt32Value)114U };
            CellValue Ans4 = new CellValue();
            Ans4.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer1.Value) == 4 ? "*" : "";
            Cell cell523 = new Cell() { CellReference = "H9", CellValue = Ans4, StyleIndex = (UInt32Value)114U };
            CellValue Ans5 = new CellValue();
            Ans5.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer1.Value) == 5 ? "*" : "";
            Cell cell524 = new Cell() { CellReference = "I9", CellValue = Ans5, StyleIndex = (UInt32Value)114U };
            CellValue Ans6 = new CellValue();
            Ans6.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer1.Value) == 6 ? "*" : "";
            Cell cell525 = new Cell() { CellReference = "J9", CellValue = Ans6, StyleIndex = (UInt32Value)114U };
            CellValue Ans7 = new CellValue();
            Ans7.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer1.Value) == 7 ? "*" : "";
            Cell cell526 = new Cell() { CellReference = "K9",CellValue = Ans7, StyleIndex = (UInt32Value)114U };
            CellValue Ans8 = new CellValue();
            Ans8.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer1.Value) == 8 ? "*" : "";
            Cell cell527 = new Cell() { CellReference = "L9", CellValue = Ans8, StyleIndex = (UInt32Value)114U };
            CellValue Ans9 = new CellValue();
            Ans9.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer1.Value) == 9 ? "*" : "";
            Cell cell528 = new Cell() { CellReference = "M9", CellValue = Ans9, StyleIndex = (UInt32Value)114U };
            CellValue Ans10 = new CellValue();
            Ans10.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer1.Value) == 10 ? "*" : "";
            Cell cell529 = new Cell() { CellReference = "N9", CellValue = Ans10, StyleIndex = (UInt32Value)115U };

            row90.Append(cell517);
            row90.Append(cell518);
            row90.Append(cell519);
            row90.Append(cell520);
            row90.Append(cell521);
            row90.Append(cell522);
            row90.Append(cell523);
            row90.Append(cell524);
            row90.Append(cell525);
            row90.Append(cell526);
            row90.Append(cell527);
            row90.Append(cell528);
            row90.Append(cell529);

            Row row91 = new Row() { RowIndex = (UInt32Value)10U, Spans = new ListValue<StringValue>() { InnerText = "1:14" }, Height = 30D, CustomHeight = true, DyDescent = 0.25D };
            Cell cell530 = new Cell() { CellReference = "B10", StyleIndex = (UInt32Value)30U };
            Cell cell531 = new Cell() { CellReference = "C10", StyleIndex = (UInt32Value)28U };

            Cell cell532 = new Cell() { CellReference = "D10", StyleIndex = (UInt32Value)37U, DataType = CellValues.SharedString };
            CellValue cellValue118 = new CellValue();
            cellValue118.Text = "36";

            cell532.Append(cellValue118);
            CellValue comment1 = new CellValue();
            comment1.Text = "";//query == null ? "" : query.Comment1;
            Cell cell533 = new Cell() { CellReference = "E10",CellValue = comment1, StyleIndex = (UInt32Value)31U };
            
            Cell cell534 = new Cell() { CellReference = "F10", StyleIndex = (UInt32Value)32U };
            Cell cell535 = new Cell() { CellReference = "G10", StyleIndex = (UInt32Value)32U };
            Cell cell536 = new Cell() { CellReference = "H10", StyleIndex = (UInt32Value)32U };
            Cell cell537 = new Cell() { CellReference = "I10", StyleIndex = (UInt32Value)32U };
            Cell cell538 = new Cell() { CellReference = "J10", StyleIndex = (UInt32Value)32U };
            Cell cell539 = new Cell() { CellReference = "K10", StyleIndex = (UInt32Value)32U };
            Cell cell540 = new Cell() { CellReference = "L10", StyleIndex = (UInt32Value)32U };
            Cell cell541 = new Cell() { CellReference = "M10", StyleIndex = (UInt32Value)32U };
            Cell cell542 = new Cell() { CellReference = "N10", StyleIndex = (UInt32Value)33U };

            row91.Append(cell530);
            row91.Append(cell531);
            row91.Append(cell532);
            row91.Append(cell533);
            row91.Append(cell534);
            row91.Append(cell535);
            row91.Append(cell536);
            row91.Append(cell537);
            row91.Append(cell538);
            row91.Append(cell539);
            row91.Append(cell540);
            row91.Append(cell541);
            row91.Append(cell542);

            Row row92 = new Row() { RowIndex = (UInt32Value)11U, Spans = new ListValue<StringValue>() { InnerText = "1:14" }, Height = 22.5D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell543 = new Cell() { CellReference = "B11", StyleIndex = (UInt32Value)133U, DataType = CellValues.SharedString };
            CellValue cellValue119 = new CellValue();
            cellValue119.Text = "14";

            cell543.Append(cellValue119);
           
            Cell cell544 = new Cell() { CellReference = "C11", StyleIndex = (UInt32Value)134U };             
            Cell cell545 = new Cell() { CellReference = "D11",  StyleIndex = (UInt32Value)135U };
            CellValue Bns1 = new CellValue();
            Bns1.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer2) == 1 ? "*" : "";
            Cell cell546 = new Cell() { CellReference = "E11",CellValue = Bns1, StyleIndex = (UInt32Value)112U };
            CellValue Bns2 = new CellValue();
            Bns2.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer2) == 2 ? "*" : "";
            Cell cell547 = new Cell() { CellReference = "F11",CellValue = Bns2, StyleIndex = (UInt32Value)112U };
            CellValue Bns3 = new CellValue();
            Bns3.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer2) == 3 ? "*" : "";
            Cell cell548 = new Cell() { CellReference = "G11", CellValue = Bns3, StyleIndex = (UInt32Value)112U };
            CellValue Bns4 = new CellValue();
            Bns4.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer2) == 4 ? "*" : "";
            Cell cell549 = new Cell() { CellReference = "H11", CellValue = Bns4, StyleIndex = (UInt32Value)112U };
            CellValue Bns5 = new CellValue();
            Bns5.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer2) == 5 ? "*" : "";
            Cell cell550 = new Cell() { CellReference = "I11", CellValue = Bns5, StyleIndex = (UInt32Value)112U };
            CellValue Bns6 = new CellValue();
            Bns6.Text = "";//query == null ? "" :  Convert.ToInt32(query.Answer2) == 6 ? "*" : "";
            Cell cell551 = new Cell() { CellReference = "J11", CellValue = Bns6, StyleIndex = (UInt32Value)112U };
            CellValue Bns7 = new CellValue();
            Bns7.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer2) == 7 ? "*" : "";
            Cell cell552 = new Cell() { CellReference = "K11", CellValue = Bns7, StyleIndex = (UInt32Value)112U };
            CellValue Bns8 = new CellValue();
            Bns8.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer2) == 8 ? "*" : "";
            Cell cell553 = new Cell() { CellReference = "L11", CellValue = Bns8, StyleIndex = (UInt32Value)112U };
            CellValue Bns9 = new CellValue();
            Bns9.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer2) == 9 ? "*" : "";
            Cell cell554 = new Cell() { CellReference = "M11", CellValue = Bns9, StyleIndex = (UInt32Value)112U };
            CellValue Bns10 = new CellValue();
            Bns10.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer2) == 10 ? "*" : "";
            Cell cell555 = new Cell() { CellReference = "N11",CellValue = Bns10, StyleIndex = (UInt32Value)113U };

            row92.Append(cell543);
            row92.Append(cell544);
            row92.Append(cell545);
            row92.Append(cell546);
            row92.Append(cell547);
            row92.Append(cell548);
            row92.Append(cell549);
            row92.Append(cell550);
            row92.Append(cell551);
            row92.Append(cell552);
            row92.Append(cell553);
            row92.Append(cell554);
            row92.Append(cell555);

            Row row93 = new Row() { RowIndex = (UInt32Value)12U, Spans = new ListValue<StringValue>() { InnerText = "1:14" }, Height = 30D, CustomHeight = true, DyDescent = 0.25D };
            Cell cell556 = new Cell() { CellReference = "B12", StyleIndex = (UInt32Value)30U };
            Cell cell557 = new Cell() { CellReference = "C12", StyleIndex = (UInt32Value)28U };

            Cell cell558 = new Cell() { CellReference = "D12", StyleIndex = (UInt32Value)37U, DataType = CellValues.SharedString };
            CellValue cellValue120 = new CellValue();
            cellValue120.Text = "36";

            cell558.Append(cellValue120);
            CellValue comment2 = new CellValue();
            comment2.Text = "";//query == null ? "" : query.Comment2;
            Cell cell559 = new Cell() { CellReference = "E12", CellValue = comment2, StyleIndex = (UInt32Value)31U };

            Cell cell560 = new Cell() { CellReference = "F12", StyleIndex = (UInt32Value)32U };
            Cell cell561 = new Cell() { CellReference = "G12", StyleIndex = (UInt32Value)32U };
            Cell cell562 = new Cell() { CellReference = "H12", StyleIndex = (UInt32Value)32U };
            Cell cell563 = new Cell() { CellReference = "I12", StyleIndex = (UInt32Value)32U };
            Cell cell564 = new Cell() { CellReference = "J12", StyleIndex = (UInt32Value)32U };
            Cell cell565 = new Cell() { CellReference = "K12", StyleIndex = (UInt32Value)32U };
            Cell cell566 = new Cell() { CellReference = "L12", StyleIndex = (UInt32Value)32U };
            Cell cell567 = new Cell() { CellReference = "M12", StyleIndex = (UInt32Value)32U };
            Cell cell568 = new Cell() { CellReference = "N12", StyleIndex = (UInt32Value)33U };

            row93.Append(cell556);
            row93.Append(cell557);
            row93.Append(cell558);
            row93.Append(cell559);
            row93.Append(cell560);
            row93.Append(cell561);
            row93.Append(cell562);
            row93.Append(cell563);
            row93.Append(cell564);
            row93.Append(cell565);
            row93.Append(cell566);
            row93.Append(cell567);
            row93.Append(cell568);

            Row row94 = new Row() { RowIndex = (UInt32Value)13U, Spans = new ListValue<StringValue>() { InnerText = "1:14" }, Height = 22.5D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell569 = new Cell() { CellReference = "B13", StyleIndex = (UInt32Value)133U, DataType = CellValues.SharedString };
            CellValue cellValue121 = new CellValue();
            cellValue121.Text = "34";

            cell569.Append(cellValue121);
            Cell cell570 = new Cell() { CellReference = "C13", StyleIndex = (UInt32Value)134U };
            Cell cell571 = new Cell() { CellReference = "D13", StyleIndex = (UInt32Value)135U };
            CellValue Cns1 = new CellValue();
            Cns1.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer3) == 1 ? "*" : "";
            Cell cell572 = new Cell() { CellReference = "E13", CellValue = Cns1, StyleIndex = (UInt32Value)110U };
            CellValue Cns2 = new CellValue();
            Cns2.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer3) == 2 ? "*" : "";
            Cell cell573 = new Cell() { CellReference = "F13",CellValue = Cns2, StyleIndex = (UInt32Value)110U };
            CellValue Cns3 = new CellValue();
            Cns3.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer3) == 3 ? "*" : "";
            Cell cell574 = new Cell() { CellReference = "G13",CellValue = Cns3, StyleIndex = (UInt32Value)110U };
            CellValue Cns4 = new CellValue();
            Cns4.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer3) == 4 ? "*" : "";
            Cell cell575 = new Cell() { CellReference = "H13",CellValue = Cns4, StyleIndex = (UInt32Value)110U };
            CellValue Cns5 = new CellValue();
            Cns5.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer3) == 5 ? "*" : "";
            Cell cell576 = new Cell() { CellReference = "I13",CellValue = Cns5, StyleIndex = (UInt32Value)110U };
            CellValue Cns6 = new CellValue();
            Cns6.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer3) == 6 ? "*" : "";
            Cell cell577 = new Cell() { CellReference = "J13",CellValue = Cns6, StyleIndex = (UInt32Value)110U };
            CellValue Cns7 = new CellValue();
            Cns7.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer3) == 7 ? "*" : "";
            Cell cell578 = new Cell() { CellReference = "K13", CellValue = Cns7, StyleIndex = (UInt32Value)110U };
            CellValue Cns8 = new CellValue();
            Cns8.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer3) == 8 ? "*" : "";
            Cell cell579 = new Cell() { CellReference = "L13", CellValue = Cns8, StyleIndex = (UInt32Value)110U };
            CellValue Cns9 = new CellValue();
            Cns9.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer3) == 9 ? "*" : "";
            Cell cell580 = new Cell() { CellReference = "M13", CellValue = Cns9, StyleIndex = (UInt32Value)110U };
            CellValue Cns10 = new CellValue();
            Cns10.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer3) == 10 ? "*" : "";
            Cell cell581 = new Cell() { CellReference = "N13", CellValue = Cns10, StyleIndex = (UInt32Value)111U };

            row94.Append(cell569);
            row94.Append(cell570);
            row94.Append(cell571);
            row94.Append(cell572);
            row94.Append(cell573);
            row94.Append(cell574);
            row94.Append(cell575);
            row94.Append(cell576);
            row94.Append(cell577);
            row94.Append(cell578);
            row94.Append(cell579);
            row94.Append(cell580);
            row94.Append(cell581);

            Row row95 = new Row() { RowIndex = (UInt32Value)14U, Spans = new ListValue<StringValue>() { InnerText = "1:14" }, Height = 37.5D, CustomHeight = true, DyDescent = 0.25D };
            Cell cell582 = new Cell() { CellReference = "B14", StyleIndex = (UInt32Value)30U };
            Cell cell583 = new Cell() { CellReference = "C14", StyleIndex = (UInt32Value)28U };

            Cell cell584 = new Cell() { CellReference = "D14", StyleIndex = (UInt32Value)37U, DataType = CellValues.SharedString };
            CellValue cellValue122 = new CellValue();
            cellValue122.Text = "36";

            cell584.Append(cellValue122);
            CellValue comment3 = new CellValue();
            comment3.Text = "";//query == null ? "" : query.Comment3;
            Cell cell585 = new Cell() { CellReference = "E14", CellValue = comment3, StyleIndex = (UInt32Value)31U };
           
            Cell cell586 = new Cell() { CellReference = "F14", StyleIndex = (UInt32Value)32U };
            Cell cell587 = new Cell() { CellReference = "G14", StyleIndex = (UInt32Value)32U };
            Cell cell588 = new Cell() { CellReference = "H14", StyleIndex = (UInt32Value)32U };
            Cell cell589 = new Cell() { CellReference = "I14", StyleIndex = (UInt32Value)32U };
            Cell cell590 = new Cell() { CellReference = "J14", StyleIndex = (UInt32Value)32U };
            Cell cell591 = new Cell() { CellReference = "K14", StyleIndex = (UInt32Value)32U };
            Cell cell592 = new Cell() { CellReference = "L14", StyleIndex = (UInt32Value)32U };
            Cell cell593 = new Cell() { CellReference = "M14", StyleIndex = (UInt32Value)32U };
            Cell cell594 = new Cell() { CellReference = "N14", StyleIndex = (UInt32Value)33U };

            row95.Append(cell582);
            row95.Append(cell583);
            row95.Append(cell584);
            row95.Append(cell585);
            row95.Append(cell586);
            row95.Append(cell587);
            row95.Append(cell588);
            row95.Append(cell589);
            row95.Append(cell590);
            row95.Append(cell591);
            row95.Append(cell592);
            row95.Append(cell593);
            row95.Append(cell594);

            Row row96 = new Row() { RowIndex = (UInt32Value)15U, Spans = new ListValue<StringValue>() { InnerText = "1:14" }, Height = 22.5D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell595 = new Cell() { CellReference = "B15", StyleIndex = (UInt32Value)133U, DataType = CellValues.SharedString };
            CellValue cellValue123 = new CellValue();
            cellValue123.Text = "35";

            cell595.Append(cellValue123);
            Cell cell596 = new Cell() { CellReference = "C15", StyleIndex = (UInt32Value)134U };
            Cell cell597 = new Cell() { CellReference = "D15", StyleIndex = (UInt32Value)135U };
            CellValue Dns1 = new CellValue();
            Cns1.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer4) == 1 ? "*" : "";
            Cell cell598 = new Cell() { CellReference = "E15",CellValue = Dns1, StyleIndex = (UInt32Value)110U };
            CellValue Dns2 = new CellValue();
            Dns2.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer4) == 2 ? "*" : "";
            Cell cell599 = new Cell() { CellReference = "F15", CellValue = Dns2, StyleIndex = (UInt32Value)110U };
            CellValue Dns3 = new CellValue();
            Dns3.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer4) == 3 ? "*" : "";
            Cell cell600 = new Cell() { CellReference = "G15", CellValue = Dns3, StyleIndex = (UInt32Value)110U };
            CellValue Dns4 = new CellValue();
            Dns4.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer4) == 4 ? "*" : "";
            Cell cell601 = new Cell() { CellReference = "H15", CellValue = Dns4, StyleIndex = (UInt32Value)110U };
            CellValue Dns5 = new CellValue();
            Dns5.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer4) == 5 ? "*" : "";
            Cell cell602 = new Cell() { CellReference = "I15", CellValue = Dns5, StyleIndex = (UInt32Value)110U };
            CellValue Dns6 = new CellValue();
            Dns6.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer4) == 6 ? "*" : "";
            Cell cell603 = new Cell() { CellReference = "J15", CellValue = Dns6, StyleIndex = (UInt32Value)110U };
            CellValue Dns7 = new CellValue();
            Dns7.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer4) == 7 ? "*" : "";
            Cell cell604 = new Cell() { CellReference = "K15", CellValue = Dns7, StyleIndex = (UInt32Value)110U };
            CellValue Dns8 = new CellValue();
            Dns8.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer4) == 8 ? "*" : "";
            Cell cell605 = new Cell() { CellReference = "L15", CellValue = Dns8, StyleIndex = (UInt32Value)110U };
            CellValue Dns9 = new CellValue();
            Dns9.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer4) == 9 ? "*" : "";
            Cell cell606 = new Cell() { CellReference = "M15", CellValue = Dns9, StyleIndex = (UInt32Value)110U };
            CellValue Dns10 = new CellValue();
            Dns10.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer4) == 10 ? "*" : "";
            Cell cell607 = new Cell() { CellReference = "N15",CellValue = Dns10, StyleIndex = (UInt32Value)111U };

            row96.Append(cell595);
            row96.Append(cell596);
            row96.Append(cell597);
            row96.Append(cell598);
            row96.Append(cell599);
            row96.Append(cell600);
            row96.Append(cell601);
            row96.Append(cell602);
            row96.Append(cell603);
            row96.Append(cell604);
            row96.Append(cell605);
            row96.Append(cell606);
            row96.Append(cell607);

            Row row97 = new Row() { RowIndex = (UInt32Value)16U, Spans = new ListValue<StringValue>() { InnerText = "1:14" }, Height = 30D, CustomHeight = true, DyDescent = 0.25D };
            Cell cell608 = new Cell() { CellReference = "B16", StyleIndex = (UInt32Value)30U };
            Cell cell609 = new Cell() { CellReference = "C16", StyleIndex = (UInt32Value)28U };

            Cell cell610 = new Cell() { CellReference = "D16", StyleIndex = (UInt32Value)37U, DataType = CellValues.SharedString };
            CellValue cellValue124 = new CellValue();
            cellValue124.Text = "36";

            cell610.Append(cellValue124);
            CellValue comment4 = new CellValue();
            comment4.Text = "";//query == null ? "" : query.Comment4;
            Cell cell611 = new Cell() { CellReference = "E16",CellValue = comment4, StyleIndex = (UInt32Value)31U };

            Cell cell612 = new Cell() { CellReference = "F16", StyleIndex = (UInt32Value)32U };
            Cell cell613 = new Cell() { CellReference = "G16", StyleIndex = (UInt32Value)32U };
            Cell cell614 = new Cell() { CellReference = "H16", StyleIndex = (UInt32Value)32U };
            Cell cell615 = new Cell() { CellReference = "I16", StyleIndex = (UInt32Value)32U };
            Cell cell616 = new Cell() { CellReference = "J16", StyleIndex = (UInt32Value)32U };
            Cell cell617 = new Cell() { CellReference = "K16", StyleIndex = (UInt32Value)32U };
            Cell cell618 = new Cell() { CellReference = "L16", StyleIndex = (UInt32Value)32U };
            Cell cell619 = new Cell() { CellReference = "M16", StyleIndex = (UInt32Value)32U };
            Cell cell620 = new Cell() { CellReference = "N16", StyleIndex = (UInt32Value)33U };

            row97.Append(cell608);
            row97.Append(cell609);
            row97.Append(cell610);
            row97.Append(cell611);
            row97.Append(cell612);
            row97.Append(cell613);
            row97.Append(cell614);
            row97.Append(cell615);
            row97.Append(cell616);
            row97.Append(cell617);
            row97.Append(cell618);
            row97.Append(cell619);
            row97.Append(cell620);

            Row row98 = new Row() { RowIndex = (UInt32Value)17U, Spans = new ListValue<StringValue>() { InnerText = "1:14" }, Height = 22.5D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell621 = new Cell() { CellReference = "B17", StyleIndex = (UInt32Value)133U, DataType = CellValues.SharedString };
            CellValue cellValue125 = new CellValue();
            cellValue125.Text = "17";

            cell621.Append(cellValue125);
            Cell cell622 = new Cell() { CellReference = "C17", StyleIndex = (UInt32Value)134U };
            Cell cell623 = new Cell() { CellReference = "D17", StyleIndex = (UInt32Value)135U };
            CellValue Ens1 = new CellValue();
            Ens1.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer5) == 1 ? "*" : "";
            Cell cell624 = new Cell() { CellReference = "E17",CellValue = Ens1, StyleIndex = (UInt32Value)110U };
            CellValue Ens2 = new CellValue();
            Ens2.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer5) == 2 ? "*" : "";
            Cell cell625 = new Cell() { CellReference = "F17",CellValue = Ens2, StyleIndex = (UInt32Value)110U };
            CellValue Ens3 = new CellValue();
            Ens3.Text = "";// query == null ? "" : Convert.ToInt32(query.Answer5) == 3 ? "*" : "";
            Cell cell626 = new Cell() { CellReference = "G17",CellValue = Ens3, StyleIndex = (UInt32Value)110U };
            CellValue Ens4 = new CellValue();
            Ens4.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer5) == 4 ? "*" : "";
            Cell cell627 = new Cell() { CellReference = "H17",CellValue = Ens4, StyleIndex = (UInt32Value)110U };
            CellValue Ens5 = new CellValue();
            Ens5.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer5) == 5 ? "*" : "";
            Cell cell628 = new Cell() { CellReference = "I17", CellValue = Ens5, StyleIndex = (UInt32Value)110U };
            CellValue Ens6 = new CellValue();
            Ens6.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer5) == 6 ? "*" : "";
            Cell cell629 = new Cell() { CellReference = "J17",CellValue = Ens6, StyleIndex = (UInt32Value)110U };
            CellValue Ens7 = new CellValue();
            Ens7.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer5) == 7 ? "*" : "";
            Cell cell630 = new Cell() { CellReference = "K17",CellValue = Ens7, StyleIndex = (UInt32Value)110U };
            CellValue Ens8 = new CellValue();
            Ens8.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer5) == 8 ? "*" : "";
            Cell cell631 = new Cell() { CellReference = "L17",CellValue = Ens8, StyleIndex = (UInt32Value)110U };
            CellValue Ens9 = new CellValue();
            Ens9.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer5) == 9 ? "*" : "";
            Cell cell632 = new Cell() { CellReference = "M17",CellValue = Ens9, StyleIndex = (UInt32Value)110U };
            CellValue Ens10 = new CellValue();
            Ens10.Text = "";//query == null ? "" : Convert.ToInt32(query.Answer5) == 10 ? "*" : "";
            Cell cell633 = new Cell() { CellReference = "N17",CellValue = Ens10, StyleIndex = (UInt32Value)111U };

            row98.Append(cell621);
            row98.Append(cell622);
            row98.Append(cell623);
            row98.Append(cell624);
            row98.Append(cell625);
            row98.Append(cell626);
            row98.Append(cell627);
            row98.Append(cell628);
            row98.Append(cell629);
            row98.Append(cell630);
            row98.Append(cell631);
            row98.Append(cell632);
            row98.Append(cell633);

            Row row99 = new Row() { RowIndex = (UInt32Value)18U, Spans = new ListValue<StringValue>() { InnerText = "1:14" }, Height = 30D, CustomHeight = true, DyDescent = 0.25D };
            Cell cell634 = new Cell() { CellReference = "B18", StyleIndex = (UInt32Value)30U };
            Cell cell635 = new Cell() { CellReference = "C18", StyleIndex = (UInt32Value)28U };

            Cell cell636 = new Cell() { CellReference = "D18", StyleIndex = (UInt32Value)37U, DataType = CellValues.SharedString };
            CellValue cellValue126 = new CellValue();
            cellValue126.Text = "36";

            cell636.Append(cellValue126);
            CellValue comment5 = new CellValue();
            comment5.Text = "";//query == null ? "" : query.Comment5;
            Cell cell637 = new Cell() { CellReference = "E18", CellValue = comment5, StyleIndex = (UInt32Value)31U };
            
            Cell cell638 = new Cell() { CellReference = "F18",StyleIndex = (UInt32Value)32U };
            Cell cell639 = new Cell() { CellReference = "G18", StyleIndex = (UInt32Value)32U };
            Cell cell640 = new Cell() { CellReference = "H18", StyleIndex = (UInt32Value)32U };
            Cell cell641 = new Cell() { CellReference = "I18", StyleIndex = (UInt32Value)32U };
            Cell cell642 = new Cell() { CellReference = "J18", StyleIndex = (UInt32Value)32U };
            Cell cell643 = new Cell() { CellReference = "K18", StyleIndex = (UInt32Value)32U };
            Cell cell644 = new Cell() { CellReference = "L18", StyleIndex = (UInt32Value)32U };
            Cell cell645 = new Cell() { CellReference = "M18", StyleIndex = (UInt32Value)32U };
            Cell cell646 = new Cell() { CellReference = "N18", StyleIndex = (UInt32Value)33U };

            row99.Append(cell634);
            row99.Append(cell635);
            row99.Append(cell636);
            row99.Append(cell637);
            row99.Append(cell638);
            row99.Append(cell639);
            row99.Append(cell640);
            row99.Append(cell641);
            row99.Append(cell642);
            row99.Append(cell643);
            row99.Append(cell644);
            row99.Append(cell645);
            row99.Append(cell646);

            Row row100 = new Row() { RowIndex = (UInt32Value)19U, Spans = new ListValue<StringValue>() { InnerText = "1:14" }, Height = 50.25D, CustomHeight = true, ThickBot = true, DyDescent = 0.3D };

            Cell cell647 = new Cell() { CellReference = "B19", StyleIndex = (UInt32Value)136U, DataType = CellValues.SharedString };
            CellValue cellValue127 = new CellValue();
            cellValue127.Text = "45";

            cell647.Append(cellValue127);
            Cell cell648 = new Cell() { CellReference = "C19", StyleIndex = (UInt32Value)137U };
            Cell cell649 = new Cell() { CellReference = "D19", StyleIndex = (UInt32Value)138U };
            Cell cell650 = new Cell() { CellReference = "E19", StyleIndex = (UInt32Value)34U };
            CellValue commentmain = new CellValue();
            commentmain.Text = "";// query == null ? "" : query.SurveyComments;
            Cell cell651 = new Cell() { CellReference = "F19",CellValue = commentmain, StyleIndex = (UInt32Value)35U };
            Cell cell652 = new Cell() { CellReference = "G19", StyleIndex = (UInt32Value)35U };
            Cell cell653 = new Cell() { CellReference = "H19", StyleIndex = (UInt32Value)35U };
            Cell cell654 = new Cell() { CellReference = "I19", StyleIndex = (UInt32Value)35U };
            Cell cell655 = new Cell() { CellReference = "J19", StyleIndex = (UInt32Value)35U };
            Cell cell656 = new Cell() { CellReference = "K19", StyleIndex = (UInt32Value)35U };
            Cell cell657 = new Cell() { CellReference = "L19", StyleIndex = (UInt32Value)35U };
            Cell cell658 = new Cell() { CellReference = "M19", StyleIndex = (UInt32Value)35U };
            Cell cell659 = new Cell() { CellReference = "N19", StyleIndex = (UInt32Value)36U };

            row100.Append(cell647);
            row100.Append(cell648);
            row100.Append(cell649);
            row100.Append(cell650);
            row100.Append(cell651);
            row100.Append(cell652);
            row100.Append(cell653);
            row100.Append(cell654);
            row100.Append(cell655);
            row100.Append(cell656);
            row100.Append(cell657);
            row100.Append(cell658);
            row100.Append(cell659);
            Row row101 = new Row() { RowIndex = (UInt32Value)20U, Spans = new ListValue<StringValue>() { InnerText = "1:14" }, Height = 5.25D, CustomHeight = true, DyDescent = 0.25D };

            Row row102 = new Row() { RowIndex = (UInt32Value)21U, Spans = new ListValue<StringValue>() { InnerText = "1:14" }, Height = 10.5D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell660 = new Cell() { CellReference = "A21", StyleIndex = (UInt32Value)139U, DataType = CellValues.SharedString };
            CellValue cellValue128 = new CellValue();
            cellValue128.Text = "44";

            cell660.Append(cellValue128);
            Cell cell661 = new Cell() { CellReference = "B21", StyleIndex = (UInt32Value)140U };
            Cell cell662 = new Cell() { CellReference = "C21", StyleIndex = (UInt32Value)140U };
            Cell cell663 = new Cell() { CellReference = "D21", StyleIndex = (UInt32Value)140U };
            Cell cell664 = new Cell() { CellReference = "E21", StyleIndex = (UInt32Value)140U };
            Cell cell665 = new Cell() { CellReference = "F21", StyleIndex = (UInt32Value)140U };
            Cell cell666 = new Cell() { CellReference = "G21", StyleIndex = (UInt32Value)140U };
            Cell cell667 = new Cell() { CellReference = "H21", StyleIndex = (UInt32Value)140U };
            Cell cell668 = new Cell() { CellReference = "I21", StyleIndex = (UInt32Value)140U };
            Cell cell669 = new Cell() { CellReference = "J21", StyleIndex = (UInt32Value)140U };
            Cell cell670 = new Cell() { CellReference = "K21", StyleIndex = (UInt32Value)140U };
            Cell cell671 = new Cell() { CellReference = "L21", StyleIndex = (UInt32Value)140U };
            Cell cell672 = new Cell() { CellReference = "M21", StyleIndex = (UInt32Value)140U };
            Cell cell673 = new Cell() { CellReference = "N21", StyleIndex = (UInt32Value)140U };

            row102.Append(cell660);
            row102.Append(cell661);
            row102.Append(cell662);
            row102.Append(cell663);
            row102.Append(cell664);
            row102.Append(cell665);
            row102.Append(cell666);
            row102.Append(cell667);
            row102.Append(cell668);
            row102.Append(cell669);
            row102.Append(cell670);
            row102.Append(cell671);
            row102.Append(cell672);
            row102.Append(cell673);

            sheetData4.Append(row83);
            sheetData4.Append(row84);
            sheetData4.Append(row85);
            sheetData4.Append(row86);
            sheetData4.Append(row87);
            sheetData4.Append(row88);
            sheetData4.Append(row89);
            sheetData4.Append(row90);
            sheetData4.Append(row91);
            sheetData4.Append(row92);
            sheetData4.Append(row93);
            sheetData4.Append(row94);
            sheetData4.Append(row95);
            sheetData4.Append(row96);
            sheetData4.Append(row97);
            sheetData4.Append(row98);
            sheetData4.Append(row99);
            sheetData4.Append(row100);
            sheetData4.Append(row101);
            sheetData4.Append(row102);

            MergeCells mergeCells3 = new MergeCells() { Count = (UInt32Value)11U };
            MergeCell mergeCell7 = new MergeCell() { Reference = "B15:D15" };
            MergeCell mergeCell8 = new MergeCell() { Reference = "B17:D17" };
            MergeCell mergeCell9 = new MergeCell() { Reference = "B19:D19" };
            MergeCell mergeCell10 = new MergeCell() { Reference = "A21:O21" };
            MergeCell mergeCell11 = new MergeCell() { Reference = "A1:O1" };
            MergeCell mergeCell12 = new MergeCell() { Reference = "A2:O2" };
            MergeCell mergeCell13 = new MergeCell() { Reference = "B8:D8" };
            MergeCell mergeCell14 = new MergeCell() { Reference = "B9:D9" };
            MergeCell mergeCell15 = new MergeCell() { Reference = "B11:D11" };
            MergeCell mergeCell16 = new MergeCell() { Reference = "B13:D13" };
            MergeCell mergeCell17 = new MergeCell() { Reference = "B7:N7" };

            mergeCells3.Append(mergeCell7);
            mergeCells3.Append(mergeCell8);
            mergeCells3.Append(mergeCell9);
            mergeCells3.Append(mergeCell10);
            mergeCells3.Append(mergeCell11);
            mergeCells3.Append(mergeCell12);
            mergeCells3.Append(mergeCell13);
            mergeCells3.Append(mergeCell14);
            mergeCells3.Append(mergeCell15);
            mergeCells3.Append(mergeCell16);
            mergeCells3.Append(mergeCell17);
            PageMargins pageMargins4 = new PageMargins() { Left = 0.333D, Right = 0.333D, Top = 1.2708333333333333D, Bottom = 0.82291666666666663D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup4 = new PageSetup() { Orientation = OrientationValues.Landscape, Id = "rId1" };

            HeaderFooter headerFooter4 = new HeaderFooter();
            OddHeader oddHeader4 = new OddHeader();
            oddHeader4.Text = "&C&G";
            OddFooter oddFooter4 = new OddFooter();
            oddFooter4.Text = "&C&G";

            headerFooter4.Append(oddHeader4);
            headerFooter4.Append(oddFooter4);
            LegacyDrawingHeaderFooter legacyDrawingHeaderFooter4 = new LegacyDrawingHeaderFooter() { Id = "rId2" };

            worksheet4.Append(sheetProperties4);
            worksheet4.Append(sheetDimension4);
            worksheet4.Append(sheetViews4);
            worksheet4.Append(sheetFormatProperties4);
            worksheet4.Append(columns4);
            worksheet4.Append(sheetData4);
            worksheet4.Append(mergeCells3);
            worksheet4.Append(pageMargins4);
            worksheet4.Append(pageSetup4);
            worksheet4.Append(headerFooter4);
            worksheet4.Append(legacyDrawingHeaderFooter4);

            worksheetPart4.Worksheet = worksheet4;
        }
        //Survey Summary
        // Generates content of worksheetPart5.
        private void GenerateWorksheetPartSurveySummary(WorksheetPart worksheetPart5)
        {
            Worksheet worksheet5 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet5.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet5.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet5.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetProperties sheetProperties5 = new SheetProperties() { CodeName = "Sheet10" };
            SheetDimension sheetDimension5 = new SheetDimension() { Reference = "A1:G29" };

            SheetViews sheetViews5 = new SheetViews();

            SheetView sheetView5 = new SheetView() { ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U, View = SheetViewValues.PageLayout };
            Selection selection5 = new Selection() { ActiveCell = "A7", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A7" } };

            sheetView5.Append(selection5);

            sheetViews5.Append(sheetView5);
            SheetFormatProperties sheetFormatProperties5 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            Columns columns5 = new Columns();
            Column column28 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 19.28515625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column29 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 24.140625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column30 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)7U, Width = 14.85546875D, Style = (UInt32Value)12U, CustomWidth = true };
            Column column31 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)16384U, Width = 9.140625D, Style = (UInt32Value)1U };

            columns5.Append(column28);
            columns5.Append(column29);
            columns5.Append(column30);
            columns5.Append(column31);

            SheetData sheetData5 = new SheetData();

            Row row103 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, Height = 15.75D, DyDescent = 0.25D };

            Cell cell674 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)125U, DataType = CellValues.String };
            CellFormula cellFormula76 = new CellFormula();
            cellFormula76.Text = "SETUP!$B$2";
            CellValue cellValue129 = new CellValue();
            cellValue129.Text = "?";
            cellFormula76.CalculateCell = true;
            cell674.Append(cellFormula76);
            cell674.Append(cellValue129);
            Cell cell675 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)132U };
            Cell cell676 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)132U };
            Cell cell677 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)132U };
            Cell cell678 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)132U };
            Cell cell679 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)132U };
            Cell cell680 = new Cell() { CellReference = "G1", StyleIndex = (UInt32Value)132U };

            row103.Append(cell674);
            row103.Append(cell675);
            row103.Append(cell676);
            row103.Append(cell677);
            row103.Append(cell678);
            row103.Append(cell679);
            row103.Append(cell680);
            sheetData5.Append(row103);
          
            CustomerPortalEntities db = new CustomerPortalEntities();
            var query = (from r in db.vw_SurveySummaryList
                         where r.CustomerID == _customerID
                         orderby r.SurveyDate descending
                         select r).ToList();
            if (query.Count() > 0)
            {
                Row row104 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };

                Cell cell681 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)123U, DataType = CellValues.String };
                CellFormula cellFormula77 = new CellFormula();
                DateTime StartDate = new DateTime();
                if (query.Count > 1)
                    StartDate = query[query.Count - 1].SurveyDate.Value;
                else
                    StartDate = query[0].SurveyDate.Value;
                DateTime EndDate = query[0].SurveyDate.Value;
                cellFormula77.Text = "=\"Service Survey Summaries Through Period Ending \" & TEXT(SETUP!$B$4,\"MMMMMMMMM DD, YYYY\") ";
                //"=\"The summary below details the \"&SETUP!$B$2&\" service calls for the assets managed by FPR for the current period.  The assets are sorted by Device ID#, to highlight potentially problematic devices (in RED).  For the period reviewed, there were \"&TEXT($I$6,\"#,###\")&\" service calls.\"";
                CellValue cellValue130 = new CellValue();
                cellValue130.Text = "";
                cellFormula77.CalculateCell = true;
                cell681.Append(cellFormula77);
                cell681.Append(cellValue130);
                Cell cell682 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)132U };
                Cell cell683 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)132U };
                Cell cell684 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)132U };
                Cell cell685 = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)132U };
                Cell cell686 = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)132U };
                Cell cell687 = new Cell() { CellReference = "G2", StyleIndex = (UInt32Value)132U };

                row104.Append(cell681);
                row104.Append(cell682);
                row104.Append(cell683);
                row104.Append(cell684);
                row104.Append(cell685);
                row104.Append(cell686);
                row104.Append(cell687);
                sheetData5.Append(row104);
            }
            Row row105 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };
            Cell cell688 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)19U };
            Cell cell689 = new Cell() { CellReference = "B3", StyleIndex = (UInt32Value)12U };

            row105.Append(cell688);
            row105.Append(cell689);

            Row row106 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, Height = 22.75D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell690 = new Cell() { CellReference = "B4", StyleIndex = (UInt32Value)199U, DataType = CellValues.String };
            CellValue cellValue131 = new CellValue();
            cellValue131.Text = "Each period, FPR petitions your feedback regarding our overall delivery of service.  The historical results provided to us are detailed below.  On a scale from 5 (worst) to 100 (best), the results below indicate how you have rated us in the past, for the respective categories.";

            cell690.Append(cellValue131);
            Cell cell691 = new Cell() { CellReference = "B4", StyleIndex = (UInt32Value)128U };
            Cell cell692 = new Cell() { CellReference = "C4", StyleIndex = (UInt32Value)128U };
            Cell cell693 = new Cell() { CellReference = "D4", StyleIndex = (UInt32Value)128U };
            Cell cell694 = new Cell() { CellReference = "E4", StyleIndex = (UInt32Value)128U };
            Cell cell695 = new Cell() { CellReference = "F4", StyleIndex = (UInt32Value)128U };
            Cell cell696 = new Cell() { CellReference = "G4", StyleIndex = (UInt32Value)128U };

            row106.Append(cell690);
            row106.Append(cell691);
            row106.Append(cell692);
            row106.Append(cell693);
            row106.Append(cell694);
            row106.Append(cell695);
            row106.Append(cell696);

            Row row107 = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, Height = 34.5D, DyDescent = 0.25D };

            Cell cell697 = new Cell() { CellReference = "A6", StyleIndex = (UInt32Value)47U, DataType = CellValues.SharedString };
            CellValue cellValue132 = new CellValue();
            cellValue132.Text = "29";

            cell697.Append(cellValue132);

            Cell cell698 = new Cell() { CellReference = "B6", StyleIndex = (UInt32Value)45U, DataType = CellValues.SharedString };
            CellValue cellValue133 = new CellValue();
            cellValue133.Text = "18";

            cell698.Append(cellValue133);

            Cell cell699 = new Cell() { CellReference = "C6", StyleIndex = (UInt32Value)45U, DataType = CellValues.SharedString };
            CellValue cellValue134 = new CellValue();
            cellValue134.Text = "13";

            cell699.Append(cellValue134);

            Cell cell700 = new Cell() { CellReference = "D6", StyleIndex = (UInt32Value)45U, DataType = CellValues.SharedString };
            CellValue cellValue135 = new CellValue();
            cellValue135.Text = "14";

            cell700.Append(cellValue135);

            Cell cell701 = new Cell() { CellReference = "E6", StyleIndex = (UInt32Value)45U, DataType = CellValues.SharedString };
            CellValue cellValue136 = new CellValue();
            cellValue136.Text = "15";

            cell701.Append(cellValue136);

            Cell cell702 = new Cell() { CellReference = "F6", StyleIndex = (UInt32Value)45U, DataType = CellValues.SharedString };
            CellValue cellValue137 = new CellValue();
            cellValue137.Text = "16";

            cell702.Append(cellValue137);

            Cell cell703 = new Cell() { CellReference = "G6", StyleIndex = (UInt32Value)46U, DataType = CellValues.SharedString };
            CellValue cellValue138 = new CellValue();
            cellValue138.Text = "17";

            cell703.Append(cellValue138);

            row107.Append(cell697);
            row107.Append(cell698);
            row107.Append(cell699);
            row107.Append(cell700);
            row107.Append(cell701);
            row107.Append(cell702);
            row107.Append(cell703);

           
            sheetData5.Append(row105);
            sheetData5.Append(row106);
            sheetData5.Append(row107);
            UInt32Value rownum = 6;
            int _even = 1;

            foreach(var row in query)
            {
                rownum++;
                if (_even == 1)
                {
                    Row row108 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };
                    row108.RowIndex = rownum;
                    Cell cell704 = new Cell() { CellReference = "A" + rownum.ToString(), StyleIndex = (UInt32Value)96U };
                    CellValue SurveyDateValue = new CellValue();
                    SurveyDateValue.Text = row.SurveyDate.Value.ToShortDateString();
                    cell704.Append(SurveyDateValue);

                    Cell cell705 = new Cell() { CellReference = "B" + rownum.ToString(), StyleIndex = (UInt32Value)117U };
                    CellValue SurveySubmittedValue = new CellValue();
                    SurveySubmittedValue.Text = row.Name;
                    cell705.Append(SurveySubmittedValue);

                    Cell cell706 = new Cell() { CellReference = "C" + rownum.ToString(), StyleIndex = (UInt32Value)97U };
                    CellValue SurveyCallExperienceValue = new CellValue();
                    SurveyCallExperienceValue.Text = row.Answer1.ToString() +"%";
                    cell706.Append(SurveyCallExperienceValue);

                    Cell cell707 = new Cell() { CellReference = "D" + rownum.ToString(), StyleIndex = (UInt32Value)97U };
                    CellValue SurveyAcctNeedsValue = new CellValue();
                    SurveyAcctNeedsValue.Text = row.Answer2.ToString() + "%";
                    cell707.Append(SurveyAcctNeedsValue);

                    Cell cell708 = new Cell() { CellReference = "E" + rownum.ToString(), StyleIndex = (UInt32Value)97U };
                    CellValue SurveyProfessionalValue = new CellValue();
                    SurveyProfessionalValue.Text = row.Answer3.ToString() + "%";
                    cell708.Append(SurveyProfessionalValue);

                    Cell cell709 = new Cell() { CellReference = "F" + rownum.ToString(), StyleIndex = (UInt32Value)97U };
                    CellValue SurveyParLevelValue = new CellValue();
                    SurveyParLevelValue.Text = row.Answer4.ToString() + "%";
                    cell709.Append(SurveyParLevelValue);

                    Cell cell710 = new Cell() { CellReference = "G" + rownum.ToString(), StyleIndex = (UInt32Value)98U };
                    CellValue SurveyServiceValue = new CellValue();
                    SurveyServiceValue.Text = row.Answer5.ToString() + "%";
                    cell710.Append(SurveyServiceValue);

                    row108.Append(cell704);
                    row108.Append(cell705);
                    row108.Append(cell706);
                    row108.Append(cell707);
                    row108.Append(cell708);
                    row108.Append(cell709);
                    row108.Append(cell710);
                    sheetData5.Append(row108);
                    rownum++;
                   
                    Row row110b = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };
                    row110b.RowIndex = rownum;
                    Cell cell710a = new Cell() { CellReference = "A" + rownum.ToString(), StyleIndex = (UInt32Value)96U };
                    CellValue NoComment = new CellValue();
                    NoComment.Text = " ";
                    cell710a.Append(NoComment);
                    Cell cell710aa = new Cell() { CellReference = "B" + rownum.ToString(), StyleIndex = (UInt32Value)97U };
                    CellValue NoComment2 = new CellValue();
                    NoComment2.Text = row.SurveyComments;
                    cell710aa.Append(NoComment2);
                    Cell cell710b = new Cell() { CellReference = "C" + rownum.ToString(), StyleIndex = (UInt32Value)97U };
                    CellValue ServiceComment = new CellValue();
                    ServiceComment.Text = row.Comment1;
                    cell710b.Append(ServiceComment);

                    Cell cell710c = new Cell() { CellReference = "D" + rownum.ToString(), StyleIndex = (UInt32Value)97U };
                    CellValue ServiceCommentc = new CellValue();
                    ServiceCommentc.Text = row.Comment2;
                    cell710c.Append(ServiceCommentc);

                    Cell cell710d = new Cell() { CellReference = "E" + rownum.ToString(), StyleIndex = (UInt32Value)97U };
                    CellValue ServiceCommentd = new CellValue();
                    ServiceCommentd.Text = row.Comment3;
                    cell710d.Append(ServiceCommentd);


                    Cell cell710e = new Cell() { CellReference = "F" + rownum.ToString(), StyleIndex = (UInt32Value)97U };
                    CellValue ServiceCommente = new CellValue();
                    ServiceCommente.Text = row.Comment4;
                    cell710e.Append(ServiceCommente);


                    Cell cell710f = new Cell() { CellReference = "G" + rownum.ToString(), StyleIndex = (UInt32Value)98U };
                    CellValue ServiceCommentf = new CellValue();
                    ServiceCommentf.Text = row.Comment5;
                    cell710f.Append(ServiceCommentf);
                    row110b.Append(cell710a);
                    row110b.Append(cell710aa);
                    row110b.Append(cell710b);
                    row110b.Append(cell710c);
                    row110b.Append(cell710d);
                    row110b.Append(cell710e);
                    row110b.Append(cell710f);

                    sheetData5.Append(row110b);
                }
            
            if (_even == 2)
            {

                Row row110 = new Row() {Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };
                row110.RowIndex = rownum;
                Cell cell704 = new Cell() { CellReference = "A" + rownum.ToString(), StyleIndex = (UInt32Value)90U };
                CellValue SurveyDateValue = new CellValue();
                SurveyDateValue.Text = row.SurveyDate.Value.ToShortDateString();
                cell704.Append(SurveyDateValue);

                Cell cell705 = new Cell() { CellReference = "B" + rownum.ToString(), StyleIndex = (UInt32Value)118U };
                CellValue SurveySubmittedValue = new CellValue();
                SurveySubmittedValue.Text = row.Name;
                cell705.Append(SurveySubmittedValue);

                Cell cell706 = new Cell() { CellReference = "C" + rownum.ToString(), StyleIndex = (UInt32Value)25U };
                CellValue SurveyCallExperienceValue = new CellValue();
                SurveyCallExperienceValue.Text = row.Answer1.ToString() + "%";
                cell706.Append(SurveyCallExperienceValue);

                Cell cell707 = new Cell() { CellReference = "D" + rownum.ToString(), StyleIndex = (UInt32Value)25U };
                CellValue SurveyAcctNeedsValue = new CellValue();
                SurveyAcctNeedsValue.Text = row.Answer2.ToString() + "%";
                cell707.Append(SurveyAcctNeedsValue);

                Cell cell708 = new Cell() { CellReference = "E" + rownum.ToString(), StyleIndex = (UInt32Value)25U };
                CellValue SurveyProfessionalValue = new CellValue();
                SurveyProfessionalValue.Text = row.Answer3.ToString() + "%";
                cell708.Append(SurveyProfessionalValue);

                Cell cell709 = new Cell() { CellReference = "F" + rownum.ToString(), StyleIndex = (UInt32Value)25U };
                CellValue SurveyParLevelValue = new CellValue();
                SurveyParLevelValue.Text = row.Answer4.ToString() + "%";
                cell709.Append(SurveyParLevelValue);

                Cell cell710 = new Cell() { CellReference = "G" + rownum.ToString(), StyleIndex = (UInt32Value)91U };
                CellValue SurveyServiceValue = new CellValue();
                SurveyServiceValue.Text = row.Answer5.ToString() +"%";
                cell710.Append(SurveyServiceValue);

                row110.Append(cell704);
                row110.Append(cell705);
                row110.Append(cell706);
                row110.Append(cell707);
                row110.Append(cell708);
                row110.Append(cell709);
                row110.Append(cell710);
                sheetData5.Append(row110);

                rownum++;
                Row row110b = new Row() {  Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };
                row110b.RowIndex = rownum;
                Cell cell710a = new Cell() { CellReference = "A" + rownum.ToString(), StyleIndex = (UInt32Value)90U };
                CellValue NoComment = new CellValue();
                NoComment.Text = " ";
                cell710a.Append(NoComment);

                Cell cell710aa = new Cell() { CellReference = "B" + rownum.ToString(), StyleIndex = (UInt32Value)25U };
                CellValue NoComment2 = new CellValue();
                NoComment2.Text = row.SurveyComments;
                cell710aa.Append(NoComment2);

                Cell cell710b = new Cell() { CellReference = "C" + rownum.ToString(), StyleIndex = (UInt32Value)25U };
                CellValue ServiceComment = new CellValue();
                ServiceComment.Text = row.Comment1;
                cell710b.Append(ServiceComment);
 
                Cell cell710c = new Cell() { CellReference = "D" + rownum.ToString(), StyleIndex = (UInt32Value)25U };
                CellValue ServiceCommentc = new CellValue();
                ServiceCommentc.Text = row.Comment2;
                cell710c.Append(ServiceCommentc);
 
                Cell cell710d = new Cell() { CellReference = "E" + rownum.ToString(), StyleIndex = (UInt32Value)25U };
                CellValue ServiceCommentd = new CellValue();
                ServiceCommentd.Text = row.Comment3;
                cell710d.Append(ServiceCommentd);

                 
                Cell cell710e = new Cell() { CellReference = "F" + rownum.ToString(), StyleIndex = (UInt32Value)25U };
                CellValue ServiceCommente = new CellValue();
                ServiceCommente.Text = row.Comment4;
                cell710e.Append(ServiceCommente);

                
                Cell cell710f = new Cell() { CellReference = "G" + rownum.ToString(), StyleIndex = (UInt32Value)91U };
                CellValue ServiceCommentf = new CellValue();
                ServiceCommentf.Text = row.Comment5;
                cell710f.Append(ServiceCommentf);


                row110b.Append(cell710a);
                row110b.Append(cell710aa);
                row110b.Append(cell710b);
                row110b.Append(cell710c);
                row110b.Append(cell710d);
                row110b.Append(cell710e);
                row110b.Append(cell710f);
                 
                sheetData5.Append(row110b);
              
            }
            if (_even == 2) _even = 1; else _even++;
        }
            
            rownum = rownum + 2;
            _even = 1;

            if (rownum < 32)
            {
                if (query.Count() == 0) rownum = rownum + 4;
                for (var i = rownum - 5; i < 32; i++)
                {

                    if (_even <= 2)
                    {
                        Row row37 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };
                        row37.RowIndex = i;
                        Cell cell1885 = new Cell() { CellReference = "A" + i.ToString(), StyleIndex = (UInt32Value)60U };
                        Cell cell251 = new Cell() { CellReference = "B" + i.ToString(), StyleIndex = (UInt32Value)54U };
                        Cell cell252 = new Cell() { CellReference = "C" + i.ToString(), StyleIndex = (UInt32Value)54U };
                        Cell cell253 = new Cell() { CellReference = "D" + i.ToString(), StyleIndex = (UInt32Value)54U };
                        Cell cell254 = new Cell() { CellReference = "E" + i.ToString(), StyleIndex = (UInt32Value)54U };
                        Cell cell255 = new Cell() { CellReference = "F" + i.ToString(), StyleIndex = (UInt32Value)54U };

                        Cell cell257 = new Cell() { CellReference = "G" + i.ToString(), StyleIndex = (UInt32Value)98U };


                        row37.Append(cell1885);
                        row37.Append(cell251);
                        row37.Append(cell252);
                        row37.Append(cell253);
                        row37.Append(cell254);
                        row37.Append(cell255);
                      
                        row37.Append(cell257);
                        sheetData5.Append(row37);

                    }

                    if (_even >= 3 && _even <= 4)
                    {
                        Row row39 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 14.25D, CustomHeight = true, DyDescent = 0.25D };
                        row39.RowIndex = i;
                        Cell cell263 = new Cell() { CellReference = "A" + i.ToString(), StyleIndex = (UInt32Value)58U };
                        Cell cell264 = new Cell() { CellReference = "B" + i.ToString(), StyleIndex = (UInt32Value)26U };
                        Cell cell265 = new Cell() { CellReference = "C" + i.ToString(), StyleIndex = (UInt32Value)21U };
                        Cell cell266 = new Cell() { CellReference = "D" + i.ToString(), StyleIndex = (UInt32Value)21U };
                        Cell cell267 = new Cell() { CellReference = "E" + i.ToString(), StyleIndex = (UInt32Value)8U };
                        Cell cell268 = new Cell() { CellReference = "F" + i.ToString(), StyleIndex = (UInt32Value)8U };
                        
                        Cell cell270 = new Cell() { CellReference = "G" + i.ToString(), StyleIndex = (UInt32Value)59U };

                        CellValue cellValue69 = new CellValue();
                        cellValue69.Text = "";


                        cell268.Append(cellValue69);

                        row39.Append(cell263);
                        row39.Append(cell264);
                        row39.Append(cell265);
                        row39.Append(cell266);
                        row39.Append(cell267);
                        row39.Append(cell268);
                        
                        row39.Append(cell270);
                        sheetData5.Append(row39);


                    }
                    if (_even == 4) _even = 1; else _even++;
                    rownum++;
                }
            }
            rownum = rownum - 5;
            Row row127 = new Row() { RowIndex = (UInt32Value)rownum, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };
            Cell cell837 = new Cell() { CellReference = "A" + rownum.ToString(), StyleIndex = (UInt32Value)92U };
            Cell cell838 = new Cell() { CellReference = "B" + rownum.ToString(), StyleIndex = (UInt32Value)119U };
            Cell cell839 = new Cell() { CellReference = "C" + rownum.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell840 = new Cell() { CellReference = "D" + rownum.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell841 = new Cell() { CellReference = "E" + rownum.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell842 = new Cell() { CellReference = "F" + rownum.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell843 = new Cell() { CellReference = "G" + rownum.ToString(), StyleIndex = (UInt32Value)121U };

            row127.Append(cell837);
            row127.Append(cell838);
            row127.Append(cell839);
            row127.Append(cell840);
            row127.Append(cell841);
            row127.Append(cell842);
             
            row127.Append(cell843);
            sheetData5.Append(row127);  
           

            MergeCells mergeCells4 = new MergeCells() { Count = (UInt32Value)3U };
            MergeCell mergeCell18 = new MergeCell() { Reference = "A1:H1" };
            MergeCell mergeCell19 = new MergeCell() { Reference = "A2:H2" };
            MergeCell mergeCell20 = new MergeCell() { Reference = "B4:G4" };

            mergeCells4.Append(mergeCell18);
            mergeCells4.Append(mergeCell19);
            mergeCells4.Append(mergeCell20);

            ConditionalFormatting conditionalFormatting1 = new ConditionalFormatting() { SequenceOfReferences = new ListValue<StringValue>() { InnerText = "C7" } };

            ConditionalFormattingRule conditionalFormattingRule1 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)19U, Priority = 116, Operator = ConditionalFormattingOperatorValues.Equal };
            Formula formula1 = new Formula();
            formula1.Text = "\"F\"";

            conditionalFormattingRule1.Append(formula1);

            ConditionalFormattingRule conditionalFormattingRule2 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)18U, Priority = 117, Operator = ConditionalFormattingOperatorValues.Equal };
            Formula formula2 = new Formula();
            formula2.Text = "\"D\"";

            conditionalFormattingRule2.Append(formula2);

            ConditionalFormattingRule conditionalFormattingRule3 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)17U, Priority = 118, Operator = ConditionalFormattingOperatorValues.Equal };
            Formula formula3 = new Formula();
            formula3.Text = "\"C\"";

            conditionalFormattingRule3.Append(formula3);

            ConditionalFormattingRule conditionalFormattingRule4 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)16U, Priority = 119, Operator = ConditionalFormattingOperatorValues.Equal };
            Formula formula4 = new Formula();
            formula4.Text = "\"B\"";

            conditionalFormattingRule4.Append(formula4);

            ConditionalFormattingRule conditionalFormattingRule5 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)15U, Priority = 120, Operator = ConditionalFormattingOperatorValues.Equal };
            Formula formula5 = new Formula();
            formula5.Text = "\"A\"";

            conditionalFormattingRule5.Append(formula5);

            conditionalFormatting1.Append(conditionalFormattingRule1);
            conditionalFormatting1.Append(conditionalFormattingRule2);
            conditionalFormatting1.Append(conditionalFormattingRule3);
            conditionalFormatting1.Append(conditionalFormattingRule4);
            conditionalFormatting1.Append(conditionalFormattingRule5);

            ConditionalFormatting conditionalFormatting2 = new ConditionalFormatting() { SequenceOfReferences = new ListValue<StringValue>() { InnerText = "D7:G7" } };

            ConditionalFormattingRule conditionalFormattingRule6 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)14U, Priority = 111, Operator = ConditionalFormattingOperatorValues.Equal };
            Formula formula6 = new Formula();
            formula6.Text = "\"F\"";

            conditionalFormattingRule6.Append(formula6);

            ConditionalFormattingRule conditionalFormattingRule7 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)13U, Priority = 112, Operator = ConditionalFormattingOperatorValues.Equal };
            Formula formula7 = new Formula();
            formula7.Text = "\"D\"";

            conditionalFormattingRule7.Append(formula7);

            ConditionalFormattingRule conditionalFormattingRule8 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)12U, Priority = 113, Operator = ConditionalFormattingOperatorValues.Equal };
            Formula formula8 = new Formula();
            formula8.Text = "\"C\"";

            conditionalFormattingRule8.Append(formula8);

            ConditionalFormattingRule conditionalFormattingRule9 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)11U, Priority = 114, Operator = ConditionalFormattingOperatorValues.Equal };
            Formula formula9 = new Formula();
            formula9.Text = "\"B\"";

            conditionalFormattingRule9.Append(formula9);

            ConditionalFormattingRule conditionalFormattingRule10 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)10U, Priority = 115, Operator = ConditionalFormattingOperatorValues.Equal };
            Formula formula10 = new Formula();
            formula10.Text = "\"A\"";

            conditionalFormattingRule10.Append(formula10);

            conditionalFormatting2.Append(conditionalFormattingRule6);
            conditionalFormatting2.Append(conditionalFormattingRule7);
            conditionalFormatting2.Append(conditionalFormattingRule8);
            conditionalFormatting2.Append(conditionalFormattingRule9);
            conditionalFormatting2.Append(conditionalFormattingRule10);

            ConditionalFormatting conditionalFormatting3 = new ConditionalFormatting() { SequenceOfReferences = new ListValue<StringValue>() { InnerText = "D8:G26" } };

            ConditionalFormattingRule conditionalFormattingRule11 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)9U, Priority = 6, Operator = ConditionalFormattingOperatorValues.Equal };
            Formula formula11 = new Formula();
            formula11.Text = "\"F\"";

            conditionalFormattingRule11.Append(formula11);

            ConditionalFormattingRule conditionalFormattingRule12 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)8U, Priority = 7, Operator = ConditionalFormattingOperatorValues.Equal };
            Formula formula12 = new Formula();
            formula12.Text = "\"D\"";

            conditionalFormattingRule12.Append(formula12);

            ConditionalFormattingRule conditionalFormattingRule13 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)7U, Priority = 8, Operator = ConditionalFormattingOperatorValues.Equal };
            Formula formula13 = new Formula();
            formula13.Text = "\"C\"";

            conditionalFormattingRule13.Append(formula13);

            ConditionalFormattingRule conditionalFormattingRule14 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)6U, Priority = 9, Operator = ConditionalFormattingOperatorValues.Equal };
            Formula formula14 = new Formula();
            formula14.Text = "\"B\"";

            conditionalFormattingRule14.Append(formula14);

            ConditionalFormattingRule conditionalFormattingRule15 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)5U, Priority = 10, Operator = ConditionalFormattingOperatorValues.Equal };
            Formula formula15 = new Formula();
            formula15.Text = "\"A\"";

            conditionalFormattingRule15.Append(formula15);

            conditionalFormatting3.Append(conditionalFormattingRule11);
            conditionalFormatting3.Append(conditionalFormattingRule12);
            conditionalFormatting3.Append(conditionalFormattingRule13);
            conditionalFormatting3.Append(conditionalFormattingRule14);
            conditionalFormatting3.Append(conditionalFormattingRule15);

            ConditionalFormatting conditionalFormatting4 = new ConditionalFormatting() { SequenceOfReferences = new ListValue<StringValue>() { InnerText = "C8:C26" } };

            ConditionalFormattingRule conditionalFormattingRule16 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)4U, Priority = 1, Operator = ConditionalFormattingOperatorValues.Equal };
            Formula formula16 = new Formula();
            formula16.Text = "\"F\"";

            conditionalFormattingRule16.Append(formula16);

            ConditionalFormattingRule conditionalFormattingRule17 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)3U, Priority = 2, Operator = ConditionalFormattingOperatorValues.Equal };
            Formula formula17 = new Formula();
            formula17.Text = "\"D\"";

            conditionalFormattingRule17.Append(formula17);

            ConditionalFormattingRule conditionalFormattingRule18 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)2U, Priority = 3, Operator = ConditionalFormattingOperatorValues.Equal };
            Formula formula18 = new Formula();
            formula18.Text = "\"C\"";

            conditionalFormattingRule18.Append(formula18);

            ConditionalFormattingRule conditionalFormattingRule19 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)1U, Priority = 4, Operator = ConditionalFormattingOperatorValues.Equal };
            Formula formula19 = new Formula();
            formula19.Text = "\"B\"";

            conditionalFormattingRule19.Append(formula19);

            ConditionalFormattingRule conditionalFormattingRule20 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)0U, Priority = 5, Operator = ConditionalFormattingOperatorValues.Equal };
            Formula formula20 = new Formula();
            formula20.Text = "\"A\"";

            conditionalFormattingRule20.Append(formula20);

            conditionalFormatting4.Append(conditionalFormattingRule16);
            conditionalFormatting4.Append(conditionalFormattingRule17);
            conditionalFormatting4.Append(conditionalFormattingRule18);
            conditionalFormatting4.Append(conditionalFormattingRule19);
            conditionalFormatting4.Append(conditionalFormattingRule20);
            PrintOptions printOptions3 = new PrintOptions() { HorizontalCentered = true };
            PageMargins pageMargins5 = new PageMargins() { Left = 0.55D, Right = 0.25D, Top = 1.3958333333333333D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup5 = new PageSetup() { Orientation = OrientationValues.Landscape, Id = "rId1" };

            HeaderFooter headerFooter5 = new HeaderFooter();
            OddHeader oddHeader5 = new OddHeader();
            oddHeader5.Text = "&C&G";
            OddFooter oddFooter5 = new OddFooter();
            oddFooter5.Text = "&C&G";

            headerFooter5.Append(oddHeader5);
            headerFooter5.Append(oddFooter5);
            LegacyDrawingHeaderFooter legacyDrawingHeaderFooter5 = new LegacyDrawingHeaderFooter() { Id = "rId2" };

            worksheet5.Append(sheetProperties5);
            worksheet5.Append(sheetDimension5);
            worksheet5.Append(sheetViews5);
            worksheet5.Append(sheetFormatProperties5);
            worksheet5.Append(columns5);
            worksheet5.Append(sheetData5);
            worksheet5.Append(mergeCells4);
            worksheet5.Append(conditionalFormatting1);
            worksheet5.Append(conditionalFormatting2);
            worksheet5.Append(conditionalFormatting3);
            worksheet5.Append(conditionalFormatting4);
            worksheet5.Append(printOptions3);
            worksheet5.Append(pageMargins5);
            worksheet5.Append(pageSetup5);
            worksheet5.Append(headerFooter5);
            worksheet5.Append(legacyDrawingHeaderFooter5);

            worksheetPart5.Worksheet = worksheet5;
        }      
        // SERVICE HISTORY
        // Generates content of worksheetPart13.
        private void GenerateWorksheetPartServiceHistory(WorksheetPart worksheetPart13)
        {
            Worksheet worksheet13 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet13.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet13.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet13.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetProperties sheetProperties5 = new SheetProperties() { CodeName = "Sheet9" };
            SheetDimension sheetDimension5 = new SheetDimension() { Reference = "A1:I29" };
            PageSetupProperties pageSetupProperties1 = new PageSetupProperties() { FitToPage = true };

            sheetProperties5.Append(pageSetupProperties1);
            SheetViews sheetViews5 = new SheetViews();

            SheetView sheetView5 = new SheetView() { ShowGridLines = false, ZoomScaleNormal = (UInt32Value)78U, WorkbookViewId = (UInt32Value)0U, View = SheetViewValues.PageLayout };
            Selection selection5 = new Selection() { ActiveCell = "B7", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "B7" } };

            sheetView5.Append(selection5);

            sheetViews5.Append(sheetView5);
            SheetFormatProperties sheetFormatProperties5 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            Columns columns5 = new Columns();
            Column column28 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 3.28515625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column29 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 12.66D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column30 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 9.11D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column31 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 11.75D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column32 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 15.00D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column33 = new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)6U, Width = 15.00D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column34 = new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 16.33D, Style = (UInt32Value)1U, CustomWidth = true};
            Column column35 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 22.33D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column36 = new Column() { Min = (UInt32Value)9U, Max = (UInt32Value)9U, Width = 22.285D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column37 = new Column() { Min = (UInt32Value)10U, Max = (UInt32Value)10U, Width = 3.285D, Style = (UInt32Value)1U, CustomWidth = true };
            columns5.Append(column28);
            columns5.Append(column29);
            columns5.Append(column30);
            columns5.Append(column31);
            columns5.Append(column32);
            columns5.Append(column33);
            columns5.Append(column34);
            columns5.Append(column35);
            columns5.Append(column36);
            columns5.Append(column37);
            SheetData sheetData5 = new SheetData();

            Row row103 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, Height = 15.75D, DyDescent = 0.25D };
            Cell cell675 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)132U };
            Cell cell674 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)125U, DataType = CellValues.String };
            CellFormula cellFormula76 = new CellFormula();
            cellFormula76.Text = "SETUP!$B$2";
            CellValue cellValue129 = new CellValue();
            cellValue129.Text = "?";
            cellFormula76.CalculateCell = true;
            cell674.Append(cellFormula76);
            cell674.Append(cellValue129);
           
            Cell cell676 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)132U };
            Cell cell677 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)132U };
            Cell cell678 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)132U };
            Cell cell679 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)132U };
            Cell cell680 = new Cell() { CellReference = "G1", StyleIndex = (UInt32Value)132U };
            Cell cell681 = new Cell() { CellReference = "H1", StyleIndex = (UInt32Value)132U };
            row103.Append(cell674);
            row103.Append(cell675);
            row103.Append(cell676);
            row103.Append(cell677);
            row103.Append(cell678);
            row103.Append(cell679);
            row103.Append(cell680);
            row103.Append(cell681);
            sheetData5.Append(row103);

            Row row103a = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, Height = 11.75D, DyDescent = 0.25D };
            Cell cell675a = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)132U };
            Cell cell674a = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)123U, DataType = CellValues.String };
            CellFormula cellFormula76a = new CellFormula();
            cellFormula76a.Text =  "=\"Service History For the Period Ending \"&TEXT(SETUP!$B$4,\"MMMMMMMM DD, YYYY\")&\"\"";
            CellValue cellValue129a = new CellValue();
            cellValue129a.Text = "?";
            cellFormula76a.CalculateCell = true;
            cell674a.Append(cellFormula76a);
            cell674a.Append(cellValue129a);
            
            Cell cell676a = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)132U };
            Cell cell677a = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)132U };
            Cell cell678a = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)132U };
            Cell cell679a = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)132U };
            Cell cell680a = new Cell() { CellReference = "G2", StyleIndex = (UInt32Value)132U };
            Cell cell681a = new Cell() { CellReference = "H2", StyleIndex = (UInt32Value)132U };
            row103a.Append(cell674a);
            row103a.Append(cell675a);
            row103a.Append(cell676a);
            row103a.Append(cell677a);
            row103a.Append(cell678a);
            row103a.Append(cell679a);
            row103a.Append(cell680a);
            row103a.Append(cell681a);
            sheetData5.Append(row103a);

            CoFreedomEntities db = new CoFreedomEntities();
            DateTime newStart = Convert.ToDateTime(_startDate);
            var firstDayOfMonth = new DateTime(newStart.Year, newStart.Month, 1);
            DateTime newEnd = Convert.ToDateTime(_period);
            var customer = (from cs in db.ARCustomers
                            where cs.CustomerID == _customerID
                            select cs.CustomerNumber).FirstOrDefault();
            var query = (from r in db.vw_CSServiceCallHistory
                         where r.CustomerID == _customerID && r.Date >= firstDayOfMonth && r.Date <= newEnd && !r.Description.Contains("Supply Order")
                         orderby r.EquipmentNumber ascending
                         select r).ToList();
            if (query.Count() > 0)
            {
                Row row104 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, Height = 23.75D, CustomHeight = true, DyDescent = 0.25D };
                Cell cell682 = new Cell() { CellReference = "A4", StyleIndex = (UInt32Value)132U };
                Cell cell682a = new Cell() { CellReference = "B4", StyleIndex = (UInt32Value)199U, DataType = CellValues.String };
                CellFormula cellFormula77 = new CellFormula();
                DateTime StartDate = new DateTime();
                if (query.Count > 1)
                    StartDate = query[query.Count - 1].Date;
                else
                    StartDate = query[0].Date ;
                DateTime EndDate = query[0].Date ;
                cellFormula77.Text = "=\"The summary below details the \"&SETUP!$B$2&\" service calls for the assets managed by FPR for the current period.  The assets are sorted by Device ID#, to highlight potentially problematic devices.  For the period reviewed, there were " + query.Count().ToString() + " service calls.\"";
                CellValue cellValue130 = new CellValue();
                cellValue130.Text = "=\"The summary below details the \"&SETUP!$B$2&\" service calls for the assets managed by FPR for the current period.  The assets are sorted by Device ID#, to highlight potentially problematic devices (in RED).  For the period reviewed, there were \"&TEXT($I$6,\"#,###\")&\" service calls.\"";
                cellFormula77.CalculateCell = true;
                cell682a.Append(cellFormula77);
                cell682a.Append(cellValue130);
                
                Cell cell683 = new Cell() { CellReference = "C4", StyleIndex = (UInt32Value)132U };
                Cell cell684 = new Cell() { CellReference = "D4", StyleIndex = (UInt32Value)132U };
                Cell cell685 = new Cell() { CellReference = "E4", StyleIndex = (UInt32Value)132U };
                Cell cell686 = new Cell() { CellReference = "F4", StyleIndex = (UInt32Value)132U };
                Cell cell687 = new Cell() { CellReference = "G4", StyleIndex = (UInt32Value)132U };
                Cell cell687a = new Cell() { CellReference = "H4", StyleIndex = (UInt32Value)132U };
                row104.Append(cell682);
                row104.Append(cell682a);
                row104.Append(cell683);
                row104.Append(cell684);
                row104.Append(cell685);
                row104.Append(cell686);
                row104.Append(cell687);
                row104.Append(cell687a);
                sheetData5.Append(row104);
            }
        

            CoFreedomEntities db3 = new CoFreedomEntities();
            var callsummary = (from cs in db3.vw_CSServiceCallSummarybyPeriod
                              where cs.CustomerID == _customerID
                              orderby cs.Period descending
                              select new
                              {
                                period    = cs.QuarterEnd,
                                NumCalls = cs.NumCalls
                              }).Take(8);


            UInt32Value rownum = 6;
            int _even = 1;
            Row row104b = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, Height = 24.5D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell697b = new Cell() { CellReference = "B6", StyleIndex = (UInt32Value)194U, DataType = CellValues.String };
            CellValue cellValue132b = new CellValue();
            cellValue132b.Text = "Service Call Summary";
            cell697b.Append(cellValue132b);
            row104b.Append(cell697b);
            sheetData5.Append(row104b);

        

            String[] AlphaRows = { "B", "C", "D", "E", "F", "G", "H", "I"};

            int i = 0;
            rownum++;


            Row row67 = new Row() { RowIndex = (UInt32Value)7U, Spans = new ListValue<StringValue>() { InnerText = "1:3" }, DyDescent = 0.25D };
            Row row67a = new Row() { RowIndex = (UInt32Value)8U, Spans = new ListValue<StringValue>() { InnerText = "1:4" }, DyDescent = 0.25D };
            Int32 recordcount = callsummary.Count();
            foreach (var rowed in callsummary)
            {
                if (i  == 0)
                {
                    Cell cell401T = new Cell() { CellReference = AlphaRows[i] + rownum.ToString(), StyleIndex = (UInt32Value)99U, DataType = CellValues.String };
                    CellValue meterGroupValue = new CellValue();
                    meterGroupValue.Text = rowed.period.Value.ToString("MMM. yyyy");
                    cell401T.Append(meterGroupValue);
                    row67.Append(cell401T);
                }
                if (i > 0 && i < (recordcount - 1))
                {
                    Cell cell401 = new Cell() { CellReference = AlphaRows[i] + rownum.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                    CellValue meterGroupValue = new CellValue();
                    meterGroupValue.Text = rowed.period.Value.ToString("MMM. yyyy");
                    cell401.Append(meterGroupValue);
                    row67.Append(cell401);
                }
                if (i == (recordcount - 1))
                {
                    Cell cell401T = new Cell() { CellReference = AlphaRows[i] + rownum.ToString(), StyleIndex = (UInt32Value)46U, DataType = CellValues.String };
                    CellValue meterGroupValue = new CellValue();
                    meterGroupValue.Text = rowed.period.Value.ToString("MMM. yyyy");
                    cell401T.Append(meterGroupValue);
                    row67.Append(cell401T);

                }
                 
             
            rownum++;

            if (i == 0)
            {
                Cell cellTotals = new Cell();
                String CellRef = AlphaRows[i] + rownum.ToString();
                cellTotals.CellReference = CellRef;
                cellTotals.StyleIndex = 215U;
                cellTotals.DataType = CellValues.String;
                CellValue totalcalls = new CellValue();
                totalcalls.Text = rowed.NumCalls.ToString();
                cellTotals.Append(totalcalls);
                row67a.Append(cellTotals);

            }
            if (i > 0  && i < (recordcount -1))
            {
                Cell cellTotals = new Cell();
                cellTotals.StyleIndex = 120U;
                String CellRef = AlphaRows[i] + rownum.ToString();
                cellTotals.CellReference = CellRef;
                CellValue totalcalls = new CellValue();
                totalcalls.Text = rowed.NumCalls.ToString();
                cellTotals.Append(totalcalls);
                row67a.Append(cellTotals);
            }
            if (i == (recordcount  - 1))
            {

                Cell cellTotals = new Cell();
                cellTotals.StyleIndex = 121U;
                String CellRef = AlphaRows[i] + rownum.ToString();
                cellTotals.CellReference = CellRef;

                CellValue totalcalls = new CellValue();
                totalcalls.Text = rowed.NumCalls.ToString();
                cellTotals.Append(totalcalls);
                row67a.Append(cellTotals);
            
            }

                i++; 
                rownum--;  
           }
            sheetData5.Append(row67);
            sheetData5.Append(row67a);

            rownum = rownum + 4;
            Row row107 = new Row() { RowIndex = (UInt32Value)rownum , Spans = new ListValue<StringValue>() { InnerText = "1:7" }, Height = 34.5D, DyDescent = 0.25D };

            Cell cell697 = new Cell() { CellReference = "B"+ rownum.ToString(), StyleIndex = (UInt32Value)47U, DataType = CellValues.String };
            CellValue cellValue132 = new CellValue();
            cellValue132.Text = "Issue Date";

            cell697.Append(cellValue132);

            Cell cell698 = new Cell() { CellReference = "C"+ rownum.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue133 = new CellValue();
            cellValue133.Text = "Submitted By";

            cell698.Append(cellValue133);

            Cell cell699 = new Cell() { CellReference = "D"+ rownum.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue134 = new CellValue();
            cellValue134.Text = "Device ID#";

            cell699.Append(cellValue134);

            Cell cell700 = new Cell() { CellReference = "E"+ rownum.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue135 = new CellValue();
            cellValue135.Text = "Model";

            cell700.Append(cellValue135);

            Cell cell701 = new Cell() { CellReference = "F"+ rownum.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue136 = new CellValue();
            cellValue136.Text = "Asset Location";

            cell701.Append(cellValue136);

            Cell cell702 = new Cell() { CellReference = "G"+ rownum.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue137 = new CellValue();
            cellValue137.Text = "Average Mo. Volume";

            cell702.Append(cellValue137);

            Cell cell703 = new Cell() { CellReference = "H"+ rownum.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue138 = new CellValue();
            cellValue138.Text = "Issue";

            cell703.Append(cellValue138);

            Cell cell799 = new Cell() { CellReference = "I"+ rownum.ToString(), StyleIndex = (UInt32Value)46U, DataType = CellValues.String };
            CellValue cellValue139 = new CellValue();
            cellValue139.Text = "Resolution Notes";

            cell799.Append(cellValue139);

            row107.Append(cell697);
            row107.Append(cell698);
            row107.Append(cell699);
            row107.Append(cell700);
            row107.Append(cell701);
            row107.Append(cell702);
            row107.Append(cell703);
            row107.Append(cell799);

        
            sheetData5.Append(row107);
           
            
            foreach (var row in query)
            {
                rownum++;
                if (_even == 1)
                {
                    Row row108 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };
                    row108.RowIndex = rownum;
                    Cell cell704 = new Cell() { CellReference = "B" + rownum.ToString(), StyleIndex = (UInt32Value)96U };
                    CellValue SurveyDateValue = new CellValue();
                    SurveyDateValue.Text = row.Date.ToShortDateString();
                    cell704.Append(SurveyDateValue);

                    Cell cell705 = new Cell() { CellReference = "C" + rownum.ToString(), StyleIndex = (UInt32Value)97U };
                    CellValue SurveySubmittedValue = new CellValue();
                    SurveySubmittedValue.Text = row.Caller;
                    cell705.Append(SurveySubmittedValue);

                    Cell cell706 = new Cell() { CellReference = "D" + rownum.ToString(), StyleIndex = (UInt32Value)97U, DataType = CellValues.String };
                    CellValue SurveyCallExperienceValue = new CellValue();
                    SurveyCallExperienceValue.Text = row.EquipmentNumber.ToString();
                    cell706.Append(SurveyCallExperienceValue);

                    Cell cell707 = new Cell() { CellReference = "E" + rownum.ToString(), StyleIndex = (UInt32Value)97U };
                    CellValue SurveyAcctNeedsValue = new CellValue();
                    SurveyAcctNeedsValue.Text = row.Model.ToString();
                    cell707.Append(SurveyAcctNeedsValue);

                    Cell cell708 = new Cell() { CellReference = "F" + rownum.ToString(), StyleIndex = (UInt32Value)97U };
                    CellValue SurveyProfessionalValue = new CellValue();
                    SurveyProfessionalValue.Text = row.Location.ToString();
                    cell708.Append(SurveyProfessionalValue);

                    Cell cell709 = new Cell() { CellReference = "G" + rownum.ToString(), StyleIndex = (UInt32Value)97U, DataType = CellValues.String };
                    CellValue SurveyParLevelValue = new CellValue();
                    SurveyParLevelValue.Text = row.EquipmentAvgMonthlyVolume.HasValue ? row.EquipmentAvgMonthlyVolume.Value.ToString("###,###") : "0" ;
                    cell709.Append(SurveyParLevelValue);

                    Cell cell710 = new Cell() { CellReference = "H" + rownum.ToString(), StyleIndex = (UInt32Value)97U, DataType = CellValues.String };
                    CellValue SurveyServiceValue = new CellValue();
                    SurveyServiceValue.Text = row.Description.Contains("Supply") ? "Internet Supply Order" : row.Description.ToString();
                    cell710.Append(SurveyServiceValue);


                    Cell cell711 = new Cell() { CellReference = "I" + rownum.ToString(), StyleIndex = (UInt32Value)98U, DataType = CellValues.String };
                    CellValue ServiceComments = new CellValue();
                    ServiceComments.Text = row.WorkOrderRemarks.ToString().Trim().Count() == 0 ? "No Work Order Remarks" : row.WorkOrderRemarks.ToString();
                    cell711.Append(ServiceComments);

                    row108.Append(cell704);
                    row108.Append(cell705);
                    row108.Append(cell706);
                    row108.Append(cell707);
                    row108.Append(cell708);
                    row108.Append(cell709);
                    row108.Append(cell710);
                    row108.Append(cell711);
                    sheetData5.Append(row108);
                    _even = 0;
                }

                else
                {

                    Row row110 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };
                    row110.RowIndex = rownum;
                    Cell cell704 = new Cell() { CellReference = "B" + rownum.ToString(), StyleIndex = (UInt32Value)90U };
                    CellValue SurveyDateValue = new CellValue();
                    SurveyDateValue.Text = row.Date.ToShortDateString();
                    cell704.Append(SurveyDateValue);

                    Cell cell705 = new Cell() { CellReference = "C" + rownum.ToString(), StyleIndex = (UInt32Value)26U };
                    CellValue SurveySubmittedValue = new CellValue();
                    SurveySubmittedValue.Text = row.Caller;
                    cell705.Append(SurveySubmittedValue);

                    Cell cell706 = new Cell() { CellReference = "D" + rownum.ToString(), StyleIndex = (UInt32Value)26U, DataType = CellValues.String };
                    CellValue SurveyCallExperienceValue = new CellValue();
                    SurveyCallExperienceValue.Text = row.EquipmentNumber.ToString();
                    cell706.Append(SurveyCallExperienceValue);

                    Cell cell707 = new Cell() { CellReference = "E" + rownum.ToString(), StyleIndex = (UInt32Value)26U };
                    CellValue SurveyAcctNeedsValue = new CellValue();
                    SurveyAcctNeedsValue.Text = row.Model.ToString();
                    cell707.Append(SurveyAcctNeedsValue);

                    Cell cell708 = new Cell() { CellReference = "F" + rownum.ToString(), StyleIndex = (UInt32Value)26U };
                    CellValue SurveyProfessionalValue = new CellValue();
                    SurveyProfessionalValue.Text = row.Location.ToString();
                    cell708.Append(SurveyProfessionalValue);

                    Cell cell709 = new Cell() { CellReference = "G" + rownum.ToString(), StyleIndex = (UInt32Value)26U, DataType = CellValues.Number };
                    CellValue SurveyParLevelValue = new CellValue();
                    SurveyParLevelValue.Text = row.EquipmentAvgMonthlyVolume.HasValue ? row.EquipmentAvgMonthlyVolume.Value.ToString() : "0";
                    cell709.Append(SurveyParLevelValue);

                    Cell cell710 = new Cell() { CellReference = "H" + rownum.ToString(), StyleIndex = (UInt32Value)26U, DataType = CellValues.String };
                    CellValue SurveyServiceValue = new CellValue();
                    SurveyServiceValue.Text = row.Description.Contains("Supply") ?"Internet Supply Order": row.Description.ToString();
                    cell710.Append(SurveyServiceValue);

                    Cell cell711 = new Cell() { CellReference = "I" + rownum.ToString(), StyleIndex = (UInt32Value)91U, DataType = CellValues.String };
                    CellValue ServiceComments = new CellValue();
                    ServiceComments.Text = row.WorkOrderRemarks.ToString().Trim().Count() == 0 ? "No Work Order Remarks" : row.WorkOrderRemarks.ToString();
                    cell711.Append(ServiceComments);

                    row110.Append(cell704);
                    row110.Append(cell705);
                    row110.Append(cell706);
                    row110.Append(cell707);
                    row110.Append(cell708);
                    row110.Append(cell709);
                    row110.Append(cell710);
                    row110.Append(cell711);
                    sheetData5.Append(row110);
                    _even = 1;
                }
                
            }

            rownum++;
          
       

            Row row127 = new Row() { RowIndex = (UInt32Value)rownum, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };
            Cell cell837 = new Cell() { CellReference = "B" + rownum.ToString(), StyleIndex = (UInt32Value)92U };
            CellValue EmptyCell = new CellValue();
            EmptyCell.Text = query.Count == 0 ? "No Service Calls to report." : "";
            cell837.Append(EmptyCell);
            Cell cell838 = new Cell() { CellReference = "C" + rownum.ToString(), StyleIndex = (UInt32Value)119U };
            Cell cell839 = new Cell() { CellReference = "D" + rownum.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell840 = new Cell() { CellReference = "E" + rownum.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell841 = new Cell() { CellReference = "F" + rownum.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell842 = new Cell() { CellReference = "G" + rownum.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell843 = new Cell() { CellReference = "H" + rownum.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell844 = new Cell() { CellReference = "I" + rownum.ToString(), StyleIndex = (UInt32Value)121U };

            row127.Append(cell837);
            row127.Append(cell838);
            row127.Append(cell839);
            row127.Append(cell840);
            row127.Append(cell841);
            row127.Append(cell842);
            row127.Append(cell843);
            row127.Append(cell844);
            sheetData5.Append(row127);


            MergeCells mergeCells4 = new MergeCells() { Count = (UInt32Value)3U };
            MergeCell mergeCell18 = new MergeCell() { Reference = "A1:J1" };
            MergeCell mergeCell19 = new MergeCell() { Reference = "A2:J2" };
            MergeCell mergeCell20 = new MergeCell() { Reference = "A3:J3" };
            MergeCell mergeCell21 = new MergeCell() { Reference = "B4:I4" };
            mergeCells4.Append(mergeCell18);
            mergeCells4.Append(mergeCell19);
            mergeCells4.Append(mergeCell20);
            mergeCells4.Append(mergeCell21);
      
            PrintOptions printOptions3 = new PrintOptions() { HorizontalCentered = true };
            PageMargins pageMargins5 = new PageMargins() { Left = 0.55D, Right = 0.25D, Top = 1.3958333333333333D, Bottom = 0.75D, Header = 0.25D, Footer = 0.25D };
            PageSetup pageSetup5 = new PageSetup() {   Orientation = OrientationValues.Landscape, Id = "rId1" };
            pageSetup5.FitToWidth = 1U;
            pageSetup5.FitToHeight = 0U;
            pageSetup5.Scale = 100;
            HeaderFooter headerFooter5 = new HeaderFooter();
            OddHeader oddHeader5 = new OddHeader();
            oddHeader5.Text = "&C&G";
            OddFooter oddFooter5 = new OddFooter();
            oddFooter5.Text = "&C&G";

            headerFooter5.Append(oddHeader5);
            headerFooter5.Append(oddFooter5);
            LegacyDrawingHeaderFooter legacyDrawingHeaderFooter5 = new LegacyDrawingHeaderFooter() { Id = "rId2" };

            worksheet13.Append(sheetProperties5);
            worksheet13.Append(sheetDimension5);
            worksheet13.Append(sheetViews5);
            worksheet13.Append(sheetFormatProperties5);
            worksheet13.Append(columns5);
            worksheet13.Append(sheetData5);
            worksheet13.Append(mergeCells4);
            worksheet13.Append(printOptions3);
            worksheet13.Append(pageMargins5);
            worksheet13.Append(pageSetup5);
            worksheet13.Append(headerFooter5);
            worksheet13.Append(legacyDrawingHeaderFooter5);

            worksheetPart13.Worksheet = worksheet13;
        }    
        //EasyLink
        // Generates content of worksheetPart6.
        private void GenerateWorksheetPartEasylink(WorksheetPart worksheetPart6)
        {
            Worksheet worksheet6 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet6.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet6.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet6.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetProperties sheetProperties6 = new SheetProperties() {  CodeName = "Sheet9" };
            SheetDimension sheetDimension6 = new SheetDimension() { Reference = "A1:E29" };
            
            SheetViews sheetViews6 = new SheetViews();

            SheetView sheetView6 = new SheetView() { ShowGridLines = false, ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U, View = SheetViewValues.PageLayout };
            Selection selection6 = new Selection() { ActiveCell = "F1", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "F1" } };

            sheetView6.Append(selection6);

            sheetViews6.Append(sheetView6);
            SheetFormatProperties sheetFormatProperties6 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            Columns columns6 = new Columns();
            Column column32 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 10.28515625D, Style = (UInt32Value)42U, CustomWidth = true };
            Column column33 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 28.7109375D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column34 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 16.28515625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column35 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 1.1D, Style = (UInt32Value)12U, CustomWidth = true };
            Column column36 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 30.140625D, Style = (UInt32Value)12U, CustomWidth = true };
            Column column37 = new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)6, Width = 9.140625D, Style = (UInt32Value)1U };

            columns6.Append(column32);
            columns6.Append(column33);
            columns6.Append(column34);
            columns6.Append(column35);
            columns6.Append(column36);
            columns6.Append(column37);

            SheetData sheetData6 = new SheetData();

            CustomerPortalEntities db = new CustomerPortalEntities();
            var FromDate = Convert.ToDateTime(_period).AddMonths(-2);
            var ToDate =  Convert.ToDateTime(_period);

            var querygroup =  db.EasylinkReportByDate(_customerID, FromDate.ToShortDateString(), ToDate.ToShortDateString()).ToList();
                             
 

            Int32 rowCount = querygroup.Count();
            Int32 totalRows = rowCount + 8;

            Row row128 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:5" }, Height = 15.75D, DyDescent = 0.25D };

            Cell cell844 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)125U, DataType = CellValues.String };
            CellFormula cellFormula78 = new CellFormula();
            cellFormula78.Text = "SETUP!$B$2";
            CellValue cellValue139 = new CellValue();
            cellValue139.Text = "?";
            cellFormula78.CalculateCell = true;
            cell844.Append(cellFormula78);
            cell844.Append(cellValue139);
            Cell cell845 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)132U };
            Cell cell846 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)132U };
            Cell cell847 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)132U };
            Cell cell848 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)132U };

            row128.Append(cell844);
            row128.Append(cell845);
            row128.Append(cell846);
            row128.Append(cell847);
            row128.Append(cell848);

            Row row129 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>()  { InnerText = "1:5" }, DyDescent = 0.25D };

            Cell cell849 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)123U, DataType = CellValues.String };
            CellFormula cellFormula79 = new CellFormula();
            cellFormula79.Text = "=\"Easylink Usage by User for Period Ending \"&(TEXT(SETUP!B4,\"MMMMMMMMM DD, YYYY\")&\"\")";
            CellValue cellValue140 = new CellValue();
            cellValue140.Text = "";
            cellFormula79.CalculateCell = true;
            cell849.Append(cellFormula79);
            cell849.Append(cellValue140);
            Cell cell850 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)132U };
            Cell cell851 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)132U };
            Cell cell852 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)132U };
            Cell cell853 = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)132U };

            row129.Append(cell849);
            row129.Append(cell850);
            row129.Append(cell851);
            row129.Append(cell852);
            row129.Append(cell853);

            Row row130 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:5" }, Height = 8.25D, CustomHeight = true, DyDescent = 0.25D };
            Cell cell854 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)41U };
            Cell cell855 = new Cell() { CellReference = "B3", StyleIndex = (UInt32Value)20U };
            Cell cell856 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)20U };
            Cell cell857 = new Cell() { CellReference = "D3", StyleIndex = (UInt32Value)20U };
            Cell cell858 = new Cell() { CellReference = "E3", StyleIndex = (UInt32Value)20U };

            row130.Append(cell854);
            row130.Append(cell855);
            row130.Append(cell856);
            row130.Append(cell857);
            row130.Append(cell858);

            Row row131 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:5" }, Height = 27.25D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell859 = new Cell() { CellReference = "B4", StyleIndex = (UInt32Value)199U, DataType = CellValues.SharedString };
            CellValue cellValue141 = new CellValue();
            cellValue141.Text = "26";

            cell859.Append(cellValue141);

            row131.Append(cell859);

            Row row132 = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:5" }, Height = 6.75D, CustomHeight = true, DyDescent = 0.25D };
            Cell cell860 = new Cell() { CellReference = "A5", StyleIndex = (UInt32Value)41U };
            Cell cell861 = new Cell() { CellReference = "B5", StyleIndex = (UInt32Value)20U };
            Cell cell862 = new Cell() { CellReference = "C5", StyleIndex = (UInt32Value)20U };
            Cell cell863 = new Cell() { CellReference = "D5", StyleIndex = (UInt32Value)20U };
            Cell cell864 = new Cell() { CellReference = "E5", StyleIndex = (UInt32Value)20U };

            row132.Append(cell860);
            row132.Append(cell861);
            row132.Append(cell862);
            row132.Append(cell863);
            row132.Append(cell864);

            sheetData6.Append(row128);
            sheetData6.Append(row129);
            sheetData6.Append(row130);
            sheetData6.Append(row131);
            sheetData6.Append(row132);
           
                    Row row133 = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:5" }, DyDescent = 0.25D };

                    Cell cell865 = new Cell() { CellReference = "B6", StyleIndex = (UInt32Value)13U, DataType = CellValues.SharedString };
                    CellValue cellValue142 = new CellValue();
                    cellValue142.Text = "22";

                    cell865.Append(cellValue142);

                    Cell cell866 = new Cell() { CellReference = "C6", StyleIndex = (UInt32Value)14U };
                    CellFormula cellFormula80 = new CellFormula();
                    cellFormula80.Text = "IF(B9=\"There is currently no Easylink usage.\",0,SUM(A9:A" + totalRows + "))";
                    CellValue cellValue143 = new CellValue();
                    cellValue143.Text = "0";
                    cellFormula80.CalculateCell = true;
                    cell866.Append(cellFormula80);
                    cell866.Append(cellValue143);
                    Cell cell867 = new Cell() { CellReference = "D6", StyleIndex = (UInt32Value)1U };

                    row133.Append(cell865);
                    row133.Append(cell866);
                    row133.Append(cell867);

                    Row row134 = new Row() { RowIndex = (UInt32Value)7U, Spans = new ListValue<StringValue>() { InnerText = "1:5" }, DyDescent = 0.25D };

                    Cell cell868 = new Cell() { CellReference = "B7", StyleIndex = (UInt32Value)13U, DataType = CellValues.SharedString };
                    CellValue cellValue144 = new CellValue();
                    cellValue144.Text = "23";

                    cell868.Append(cellValue144);

                    Cell cell869 = new Cell() { CellReference = "C7", StyleIndex = (UInt32Value)38U };
                    CellFormula cellFormula81 = new CellFormula();
                    cellFormula81.Text = "SUM(E9:E" + totalRows + ")";
                    CellValue cellValue145 = new CellValue();
                    cellValue145.Text = "0";
                    cellFormula81.CalculateCell = true;
                    cell869.Append(cellFormula81);
                    cell869.Append(cellValue145);
                    Cell cell870 = new Cell() { CellReference = "D7", StyleIndex = (UInt32Value)1U };

                    row134.Append(cell868);
                    row134.Append(cell869);
                    row134.Append(cell870);
                   
                   
                    Row row135 = new Row() { RowIndex = (UInt32Value)8U, Spans = new ListValue<StringValue>() { InnerText = "1:5" }, Height = 30D, CustomHeight = true, DyDescent = 0.25D };

                    Cell cell871 = new Cell() { CellReference = "B8", StyleIndex = (UInt32Value)47U, DataType = CellValues.SharedString };
                    CellValue cellValue146 = new CellValue();
                    cellValue146.Text = "19";

                    cell871.Append(cellValue146);
                    Cell cell872 = new Cell() { CellReference = "C8", StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                    CellValue FaxLabel = new CellValue();
                    FaxLabel.Text = "Fax Number";

                    cell872.Append(FaxLabel);
                    Cell cell873 = new Cell() { CellReference = "D8", StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                    CellValue cellValue147 = new CellValue();
                    cellValue147.Text = "";

                    cell873.Append(cellValue147);

                    Cell cell874 = new Cell() { CellReference = "E8", StyleIndex = (UInt32Value)46U, DataType = CellValues.SharedString };
                    CellValue cellValue148 = new CellValue();
                    cellValue148.Text = "21";

                    cell874.Append(cellValue148);

                    row135.Append(cell871);
                    row135.Append(cell872);
                    row135.Append(cell873);
                    row135.Append(cell874);

                   
                    sheetData6.Append(row133);
                    sheetData6.Append(row134);
                    sheetData6.Append(row135);
                    
                    UInt32Value rownum = 8U;
                    if (rowCount == 0)
                    {

                        rownum++;
                        Row row136 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:5" }, DyDescent = 0.25D };
                        row136.RowIndex = rownum;
                        Cell cell875 = new Cell() { CellReference = "A" + rownum.ToString(), StyleIndex = (UInt32Value)42U };
                        CellValue cellValue149 = new CellValue();
                        cellValue149.Text = "";
                       
 
                        cell875.Append(cellValue149);

                        Cell cell876 = new Cell() { CellReference = "B" + rownum.ToString(), StyleIndex = (UInt32Value)60U, DataType = CellValues.String };
                        CellValue cellValue150 = new CellValue();
                        cellValue150.Text =  "There is currently no Easylink usage.";

                        cell876.Append(cellValue150);
                        Cell cell877 = new Cell() { CellReference = "C" + rownum.ToString(), StyleIndex = (UInt32Value)84U };
                        Cell cell878 = new Cell() { CellReference = "D" + rownum.ToString(), StyleIndex = (UInt32Value)85U };
                        Cell cell879 = new Cell() { CellReference = "E" + rownum.ToString(), StyleIndex = (UInt32Value)86U };
                        CellValue TotalPageValue = new CellValue();
                        TotalPageValue.Text ="";
                        cell879.Append(TotalPageValue);
                        row136.Append(cell875);
                        row136.Append(cell876);
                        row136.Append(cell877);
                        row136.Append(cell878);
                        row136.Append(cell879);


                        sheetData6.Append(row136);
                        rownum++;
                    }
                    Int32 evenodd = 0;

                    foreach (var row in querygroup)
                    {
                    rownum++;
                    if (evenodd < 2)
                    {
               
                            Row row136 = new Row() {Spans = new ListValue<StringValue>() { InnerText = "1:5" }, DyDescent = 0.25D };
                            row136.RowIndex = rownum;
                            Cell cell875 = new Cell() { CellReference = "A"+ rownum.ToString(), StyleIndex = (UInt32Value)42U };
                            CellFormula cellFormula82 = new CellFormula();
                            cellFormula82.Text = "IF(B" + rownum.ToString() + "=\"\",\"\",1)";
                            CellValue cellValue149 = new CellValue();
                            cellValue149.Text = "1";
                            cellFormula82.CalculateCell = true;
                            cell875.Append(cellFormula82);
                            cell875.Append(cellValue149);

                            Cell cell876 = new Cell() { CellReference = "B" + rownum.ToString(), StyleIndex = (UInt32Value)60U, DataType = CellValues.String };
                            CellValue cellValue150 = new CellValue();
                            cellValue150.Text = querygroup.Count() > 0 ? row.emailaddress : "There is currently no Easylink usage.";

                            cell876.Append(cellValue150);
                            Cell cell877 = new Cell() { CellReference = "C" + rownum.ToString(), StyleIndex = (UInt32Value)173U };
                            CellValue FaxNumber = new CellValue();
                            FaxNumber.Text = row.FaxNumber;
                            cell877.Append(FaxNumber);
                            Cell cell878 = new Cell() { CellReference = "D" + rownum.ToString(), StyleIndex = (UInt32Value)85U };
                            Cell cell879 = new Cell() { CellReference = "E" + rownum.ToString(), StyleIndex = (UInt32Value)86U };
                            CellValue TotalPageValue = new CellValue();
                            TotalPageValue.Text = row.Pages.ToString();
                            cell879.Append(TotalPageValue);
                            row136.Append(cell875);
                            row136.Append(cell876);
                            row136.Append(cell877);
                            row136.Append(cell878);
                            row136.Append(cell879);

                       
                            sheetData6.Append(row136);
                         
                        }
                        if (evenodd >= 2 && evenodd < 4)
                        {
                            Row row138 = new Row() {  Spans = new ListValue<StringValue>() { InnerText = "1:5" }, DyDescent = 0.25D };
                            row138.RowIndex = rownum;
                            Cell cell885 = new Cell() { CellReference = "A" + rownum.ToString(), StyleIndex = (UInt32Value)42U };
                            CellFormula cellFormula82 = new CellFormula();
                            cellFormula82.Text = "IF(B" + rownum.ToString() + " =\"\",\"\",1)";
                            CellValue cellValue149 = new CellValue();
                            cellValue149.Text = "1";
                            cellFormula82.CalculateCell = true;
                            cell885.Append(cellFormula82);
                            cell885.Append(cellValue149);   
                           
                            Cell cell886 = new Cell() { CellReference = "B" + rownum.ToString(), StyleIndex = (UInt32Value)58U };
                            CellValue cellValue152 = new CellValue();
                            cellValue152.Text = row.emailaddress;
                            cell886.Append(cellValue152);
                            
                            Cell cell887 = new Cell() { CellReference = "C" + rownum.ToString(), StyleIndex = (UInt32Value)172U };
                            CellValue FaxNumber = new CellValue();
                            FaxNumber.Text = row.FaxNumber;
                            cell887.Append(FaxNumber);

                           // Cell cell887 = new Cell() { CellReference = "C" + rownum.ToString(), StyleIndex = (UInt32Value)7U };
                            Cell cell888 = new Cell() { CellReference = "D" + rownum.ToString(), StyleIndex = (UInt32Value)15U };
                            Cell cell889 = new Cell() { CellReference = "E" + rownum.ToString(), StyleIndex = (UInt32Value)80U };
                            CellValue TotalPageValue = new CellValue();
                            TotalPageValue.Text = row.Pages.ToString();
                            cell889.Append(TotalPageValue);
                            row138.Append(cell885);
                            row138.Append(cell886);
                            row138.Append(cell887);
                            row138.Append(cell888);
                            row138.Append(cell889);

                  
                            sheetData6.Append(row138);
                             
                        }
                        if (evenodd == 3) evenodd = 0; else evenodd++;
            }
                    cellFormula81.CalculateCell = true;
                    if (querygroup.Count() > 0) rownum++;
                    Row row127 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };
                    row127.RowIndex = rownum;
                    Cell cell837 = new Cell() { CellReference = "B" + rownum.ToString(), StyleIndex = (UInt32Value)92U };
                    
                    Cell cell839 = new Cell() { CellReference = "C" + rownum.ToString(), StyleIndex = (UInt32Value)120U };
                    Cell cell840 = new Cell() { CellReference = "D" + rownum.ToString(), StyleIndex = (UInt32Value)120U };
                    Cell cell843 = new Cell() { CellReference = "E" + rownum.ToString(), StyleIndex = (UInt32Value)121U };

                    row127.Append(cell837);
                   // row127.Append(cell838);
                    row127.Append(cell839);
                    row127.Append(cell840);
                   
                    row127.Append(cell843);
                    sheetData6.Append(row127);

                MergeCells mergeCells5 = new MergeCells() { Count = (UInt32Value)3U };
                MergeCell mergeCell21 = new MergeCell() { Reference = "A1:I1" };
                MergeCell mergeCell22 = new MergeCell() { Reference = "A2:I2" };
                MergeCell mergeCell23 = new MergeCell() { Reference = "B4:I4" };

                mergeCells5.Append(mergeCell21);
                mergeCells5.Append(mergeCell22);
                mergeCells5.Append(mergeCell23);
                PrintOptions printOptions3 = new PrintOptions() { HorizontalCentered = true };
                PageMargins pageMargins6 = new PageMargins() { Left = 0.553D, Right = 0.333D, Top = 1.3541666666666667D, Bottom = 1.0208333333333333D, Header = 0.3D, Footer = 0.3D };
                PageSetup pageSetup6 = new PageSetup() {   Orientation = OrientationValues.Landscape, Id = "rId1" };
              
                HeaderFooter headerFooter6 = new HeaderFooter();
                OddHeader oddHeader6 = new OddHeader();
                oddHeader6.Text = "&C&G";
                OddFooter oddFooter6 = new OddFooter();
                oddFooter6.Text = "&C&G";

                headerFooter6.Append(oddHeader6);
                headerFooter6.Append(oddFooter6);
                LegacyDrawingHeaderFooter legacyDrawingHeaderFooter6 = new LegacyDrawingHeaderFooter() { Id = "rId2" };

                worksheet6.Append(sheetProperties6);
                worksheet6.Append(sheetDimension6);
                worksheet6.Append(sheetViews6);
                worksheet6.Append(sheetFormatProperties6);
                worksheet6.Append(columns6);
                worksheet6.Append(sheetData6);
                worksheet6.Append(mergeCells5);
                worksheet6.Append(printOptions3);
                worksheet6.Append(pageMargins6);
                worksheet6.Append(pageSetup6);
                worksheet6.Append(headerFooter6);
                worksheet6.Append(legacyDrawingHeaderFooter6);

                worksheetPart6.Worksheet = worksheet6;
            
        }
        //Revision Summary 
        // Generates content of worksheetPart2.
        
        private void GenerateWorksheetPartRevisionSummary(WorksheetPart worksheetPart2)
        {
            Worksheet worksheet2 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet2.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet2.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetDimension sheetDimension2 = new SheetDimension() { Reference = "A1:I31" };

            SheetViews sheetViews2 = new SheetViews();

            SheetView sheetView2 = new SheetView() { ShowGridLines = false, ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U, View = SheetViewValues.PageLayout };
            Selection selection2 = new Selection() { ActiveCell = "A7", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A7" } };

            sheetView2.Append(selection2);

            sheetViews2.Append(sheetView2);
            SheetFormatProperties sheetFormatProperties2 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            Columns columns2 = new Columns();
            Column column28 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 14.28515625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column29 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 19.140625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column30 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 10.5D, Style = (UInt32Value)12U, CustomWidth = true };
            Column column31 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 12.5D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column32 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 13.5D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column33 = new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)6U, Width = 10.5D, Style = (UInt32Value)12U, CustomWidth = true };
            Column column34 = new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 10.5D, Style = (UInt32Value)12U, CustomWidth = true };
            Column column35 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 11.5D, Style = (UInt32Value)12U, CustomWidth = true };
            Column column36 = new Column() { Min = (UInt32Value)9U, Max = (UInt32Value)9U, Width = 11.5D, Style = (UInt32Value)12U, CustomWidth = true };
            Column column37 = new Column() { Min = (UInt32Value)10U, Max = (UInt32Value)10U, Width = 12.888D, Style = (UInt32Value)12U, CustomWidth = true };

            columns2.Append(column28);
            columns2.Append(column29);
            columns2.Append(column30);
            columns2.Append(column31);
            columns2.Append(column32);
            columns2.Append(column33);
            columns2.Append(column34);
            columns2.Append(column35);
            columns2.Append(column36);
            columns2.Append(column37);
            SheetData sheetData2 = new SheetData();

            Row row103 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, Height = 15.75D, DyDescent = 0.25D };
           
            Cell cell674 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)125U, DataType = CellValues.String };
            CellFormula cellFormula76 = new CellFormula();
            cellFormula76.Text = "SETUP!$B$2";
            CellValue cellValue129 = new CellValue();
            cellValue129.Text = "?";
            cellFormula76.CalculateCell = true;
            cell674.Append(cellFormula76);
            cell674.Append(cellValue129);
           
            Cell cell676 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)132U };
            Cell cell677 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)132U };
            Cell cell678 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)132U };
            Cell cell679 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)132U };
            Cell cell680 = new Cell() { CellReference = "G1", StyleIndex = (UInt32Value)132U };
            Cell cell680a = new Cell() { CellReference = "H1", StyleIndex = (UInt32Value)132U };
            Cell cell675 = new Cell() { CellReference = "J1", StyleIndex = (UInt32Value)132U };
            row103.Append(cell674);
          
            row103.Append(cell676);
            row103.Append(cell677);
            row103.Append(cell678);
            row103.Append(cell679);
            row103.Append(cell680);
            row103.Append(cell680a);
            row103.Append(cell675);
            sheetData2.Append(row103);

            Row row103a = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };

            
            Cell cell682a = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)123U, DataType = CellValues.String };
            CellFormula cellFormula77a = new CellFormula();
            cellFormula77a.Text = "=\"REVision Summary Through the Period Ending \"&TEXT(SETUP!$B$4,\"MMMMMMMM DD, YYYY\")&\"\"";
            //"=\"The summary below details the \"&SETUP!$B$2&\" service calls for the assets managed by FPR for the current period.  The assets are sorted by Device ID#, to highlight potentially problematic devices (in RED).  For the period reviewed, there were \"&TEXT($I$6,\"#,###\")&\" service calls.\"";
            CellValue cellValue130a = new CellValue();
            cellValue130a.Text = "=\"REVision Summary Through the Period Ending \"&TEXT(SETUP!$B$4,\"MMMMMMMM DD, YYYY\")&\"\"";
            cellFormula77a.CalculateCell = true;
            cell682a.Append(cellFormula77a);
            cell682a.Append(cellValue130a);
           
            Cell cell683a = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)132U };
            Cell cell684a = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)132U };
            Cell cell685a = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)132U };
            Cell cell686a = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)132U };
            Cell cell687a = new Cell() { CellReference = "G2", StyleIndex = (UInt32Value)132U };
            Cell cell688a = new Cell() { CellReference = "H2", StyleIndex = (UInt32Value)132U };
            Cell cell681a = new Cell() { CellReference = "J2", StyleIndex = (UInt32Value)123U };
            row103a.Append(cell682a);
            row103a.Append(cell683a);
            row103a.Append(cell684a);
            row103a.Append(cell685a);
            row103a.Append(cell686a);
            row103a.Append(cell687a);
            row103a.Append(cell688a);
            row103a.Append(cell681a);
            sheetData2.Append(row103a);

                Row row104 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, Height = 51.5D, CustomHeight = true, DyDescent = 0.25D };
                Cell cell682 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)132U };
                Cell cell681 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)199U, DataType = CellValues.String };
                CellValue cellValue130 = new CellValue();
                cellValue130.Text = "REVision is a summary of the savings delivered by FPR in two key areas:  a)  the difference in base contract expense with FPR as it compares to your \"before\" FPR base operating expense and  b) your net savings per period as it relates to the overage expense you incurred with FPR (if applicable) compared to what you \"would have\" spent in overage expense had you not contracted with FPR to manage this portion of your business.";               
                cell681.Append(cellValue130);
               
                Cell cell683 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)132U };
                Cell cell684 = new Cell() { CellReference = "D3", StyleIndex = (UInt32Value)132U };
                Cell cell685 = new Cell() { CellReference = "E3", StyleIndex = (UInt32Value)132U };
                Cell cell686 = new Cell() { CellReference = "F3", StyleIndex = (UInt32Value)132U };
                Cell cell687 = new Cell() { CellReference = "G3", StyleIndex = (UInt32Value)132U };
                Cell cell688 = new Cell() { CellReference = "H3", StyleIndex = (UInt32Value)132U };
                row104.Append(cell681);
                row104.Append(cell682);
                row104.Append(cell683);
                row104.Append(cell684);
                row104.Append(cell685);
                row104.Append(cell686);
                row104.Append(cell687);
                row104.Append(cell688);
                sheetData2.Append(row104);

                Row row104a = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, Height = 15.0D, CustomHeight = true, DyDescent = 0.25D };
            
                sheetData2.Append(row104a);
   
   

            Row row107 = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, Height = 34.5D, DyDescent = 0.25D };
            Cell cell697 = new Cell() { CellReference = "B5", StyleIndex = (UInt32Value)47U, DataType = CellValues.String };
            CellValue cellValue132 = new CellValue();
            cellValue132.Text = "Reconciliation Period";

            cell697.Append(cellValue132);

            Cell cell698 = new Cell() { CellReference = "C5", StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue133 = new CellValue();
            cellValue133.Text = "Old Cost";

            cell698.Append(cellValue133);

            Cell cell699a = new Cell() { CellReference = "D5", StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue134a = new CellValue();
            cellValue134a.Text = "Overage Cost W/O FPR";

            cell699a.Append(cellValue134a);

            Cell cell700 = new Cell() { CellReference = "E5", StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue135 = new CellValue();
            cellValue135.Text = "Credits";

            cell700.Append(cellValue135);

            Cell cell701 = new Cell() { CellReference = "F5", StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue136 = new CellValue();
            cellValue136.Text = "New Cost";

            cell701.Append(cellValue136);

            Cell cell702 = new Cell() { CellReference = "G5", StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue137 = new CellValue();
            cellValue137.Text = "FPR Overage Cost";

            cell702.Append(cellValue137);

            Cell cell703 = new Cell() { CellReference = "H5", StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue138 = new CellValue();
            cellValue138.Text = "Savings / Expense Increase";

            cell703.Append(cellValue138);

            Cell cell799 = new Cell() { CellReference = "I5", StyleIndex = (UInt32Value)46U, DataType = CellValues.String };
            CellValue cellValue139 = new CellValue();
            cellValue139.Text = "Percent Savings";

            cell799.Append(cellValue139);
            
            row107.Append(cell697);
            row107.Append(cell698);
            row107.Append(cell699a);
            row107.Append(cell700);
            row107.Append(cell701);
            row107.Append(cell702);
            row107.Append(cell703);
            row107.Append(cell799);


            sheetData2.Append(row107);

          


            //Data loop

            ExcelRevisionExport export = new ExcelRevisionExport();
            UInt32Value RowIndex = 6 ;
            var VisionDataSummary = new List<GVWebapi.Models.VisionData>();
            DateTime overridedate = Convert.ToDateTime(_overrideDate);
            DateTime perioddate = Convert.ToDateTime(_period);
            VisionDataSummary = export.RevisionSummary(_contractID).Where(o=> o.ClientStartDate >= overridedate && o.ClientPeriodDate <= perioddate).ToList();
            int _even = 0;
            foreach (GVWebapi.Models.VisionData VDL2 in VisionDataSummary)
            {
               
                if (_even == 0)
                {
                    sheetData2.Append(CreateContentSummaryBLU(RowIndex, VDL2));
                    _even = 1;
                  
                }
                else
                {
                    sheetData2.Append(CreateContentSummaryWHT(RowIndex, VDL2));
                  
                    _even = 0;
                }
                RowIndex++;
            }


            UInt32Value StartIndex = 6;
            UInt32Value EndIndex = RowIndex - 1;
            UInt32Value RowTotals = RowIndex;
           

            //Summary Rows
            Row row12 = new Row() { RowIndex = (UInt32Value)RowTotals, Spans = new ListValue<StringValue>() { InnerText = "1:9" }, Hidden = true, DyDescent = 0.2D };

            Cell cell36 = new Cell() { CellReference = "B" + RowTotals, StyleIndex = (UInt32Value)132U, DataType = CellValues.String };
            CellValue cellValue22 = new CellValue();
            cellValue22.Text = "5";

            cell36.Append(cellValue22);

            Cell cell37 = new Cell() { CellReference = "C" + RowTotals, StyleIndex = (UInt32Value)132U };
            CellFormula cellFormula3 = new CellFormula();
            cellFormula3.Text = "SUM(C" + StartIndex.ToString() + ":C" + EndIndex.ToString() + ")";
            CellValue cellValue23 = new CellValue();
            cellValue23.Text = "0.00";

            cell37.Append(cellFormula3);
            cell37.Append(cellValue23);

            Cell cell38 = new Cell() { CellReference = "D" + RowTotals, StyleIndex = (UInt32Value)132U };
            CellFormula cellFormula4 = new CellFormula();
            cellFormula4.Text = "SUM(D" + StartIndex + ":D" + EndIndex.ToString() + ")";
            CellValue cellValue24 = new CellValue();

            cell38.Append(cellFormula4);
            cell38.Append(cellValue24);

            Cell cell39 = new Cell() { CellReference = "E" + RowTotals, StyleIndex = (UInt32Value)132U };
            CellFormula cellFormula105 = new CellFormula();
            cellFormula105.Text = "SUM(E" + StartIndex + ":E" + EndIndex.ToString() + ")";
            CellValue cellValue125 = new CellValue();
            cell39.Append(cellFormula105);
            cell39.Append(cellValue125);

            Cell cell40 = new Cell() { CellReference = "F" + RowTotals, StyleIndex = (UInt32Value)132U };
            CellFormula cellFormula5 = new CellFormula();
            cellFormula5.Text = "SUM(F" + StartIndex + ":F" + EndIndex.ToString() + ")";
            CellValue cellValue25 = new CellValue();


            cell40.Append(cellFormula5);
            cell40.Append(cellValue25);

            Cell cell41 = new Cell() { CellReference = "G" + RowTotals, StyleIndex = (UInt32Value)132U };
            CellFormula cellFormula6 = new CellFormula();
            cellFormula6.Text = "SUM(G" + StartIndex + ":G" + EndIndex.ToString() + ")";
            CellValue cellValue26 = new CellValue();


            cell41.Append(cellFormula6);
            cell41.Append(cellValue26);
            Cell cell42 = new Cell() { CellReference = "H" + RowTotals, StyleIndex = (UInt32Value)132U };
            CellFormula cellFormula7 = new CellFormula();
            cellFormula7.Text = "SUM(H" + StartIndex + ":H" + EndIndex.ToString() + ")";
            CellValue cellValue27 = new CellValue();

            cell42.Append(cellFormula7);
            cell42.Append(cellValue27);


            Cell cell43 = new Cell() { CellReference = "I" + RowTotals, StyleIndex = (UInt32Value)132U };


            row12.Append(cell36);
            row12.Append(cell37);
            row12.Append(cell38);
            row12.Append(cell39);
            row12.Append(cell40);
            row12.Append(cell41);
            row12.Append(cell42);
            row12.Append(cell43);
 //Summary Data           
            RowTotals++;
            UInt32Value ClientTotals = RowTotals;
            EndIndex++;
            Row row13 = new Row() { RowIndex = (UInt32Value)RowTotals, Spans = new ListValue<StringValue>() { InnerText = "1:9" }, DyDescent = 0.2D };
             Cell cell48b = new Cell() { CellReference = "B" + RowTotals, StyleIndex = (UInt32Value)174U };
            Cell cell48c = new Cell() { CellReference = "C" + RowTotals, StyleIndex = (UInt32Value)174U };


            Cell cell48 = new Cell() { CellReference = "D" + RowTotals, StyleIndex = (UInt32Value)209U, DataType = CellValues.String };
            CellValue cvD = new CellValue();
            cvD.Text = "What You Would Have Spent Without FPR:";

            cell48.Append(cvD);

            Cell cell49 = new Cell() { CellReference = "E" + RowTotals, StyleIndex = (UInt32Value)209U };
            CellFormula cellFormulaE7 = new CellFormula();
            cellFormulaE7.Text = "SUM(C" + EndIndex.ToString() + "+D" + EndIndex.ToString() + ")";
            CellValue cellValue28 = new CellValue();
            cellValue28.Text = "133381.19041487278";
            cell49.Append(cellFormulaE7);
            cell49.Append(cellValue28);
            Cell cell48f = new Cell() { CellReference = "F" + RowTotals, StyleIndex = (UInt32Value)174U };
            Cell cell48g = new Cell() { CellReference = "G" + RowTotals, StyleIndex = (UInt32Value)174U };
            Cell cell48h = new Cell() { CellReference = "H" + RowTotals, StyleIndex = (UInt32Value)174U };
            Cell cell48i = new Cell() { CellReference = "I" + RowTotals, StyleIndex = (UInt32Value)174U };

            
            row13.Append(cell48b);
            row13.Append(cell48c);
            row13.Append(cell48);
            row13.Append(cell49);
            row13.Append(cell48f);
            row13.Append(cell48g);
            row13.Append(cell48h);
            row13.Append(cell48i);
            
 
             RowTotals++;
            UInt32Value FPRTotals = RowTotals;
            Row row14 = new Row() { RowIndex = (UInt32Value)RowTotals, Spans = new ListValue<StringValue>() { InnerText = "1:9" }, DyDescent = 0.2D };

            Cell cell57 = new Cell() { CellReference = "D" + RowTotals, StyleIndex = (UInt32Value)164U, DataType = CellValues.String };
            CellValue cellValue29 = new CellValue();
            cellValue29.Text = "What You Did Spend With FPR:";

            cell57.Append(cellValue29);

            Cell cell58 = new Cell() { CellReference = "E" + RowTotals, StyleIndex = (UInt32Value)210U };
            CellFormula cellFormula8 = new CellFormula();
            cellFormula8.Text = "SUM(F" + EndIndex.ToString() + " + G" + EndIndex.ToString() + ")";
            CellValue cellValue30 = new CellValue();
            cellValue30.Text = "0";

            cell58.Append(cellFormula8);
            cell58.Append(cellValue30);

            row14.Append(cell57);
            row14.Append(cell58);

            RowTotals++;
            UInt32Value SavingTotals = RowTotals;
            Row row15 = new Row() { RowIndex = (UInt32Value)RowTotals, Spans = new ListValue<StringValue>() { InnerText = "1:9" }, DyDescent = 0.2D };


            Cell cell66 = new Cell() { CellReference = "D" + RowTotals, StyleIndex = (UInt32Value)164U, DataType = CellValues.String };
            CellValue cellValue31 = new CellValue();
            cellValue31.Text = "What FPR Has Saved you for the Period Reviewed:";

            cell66.Append(cellValue31);

            Cell cell67 = new Cell() { CellReference = "E" + RowTotals, StyleIndex = (UInt32Value)210U };
            CellFormula cellFormula9 = new CellFormula();
            UInt32Value CreditRow = FPRTotals + 2;
            cellFormula9.Text = "SUM(E" + ClientTotals + "-E" + FPRTotals + ")";
            CellValue cellValue32 = new CellValue();
            cellValue32.Text = "0";

            cell67.Append(cellFormula9);
            cell67.Append(cellValue32);


            row15.Append(cell66);
            row15.Append(cell67);
            
            _revisionTotalString = "='Revision Summary'!$E$" + RowTotals;

            RowTotals++;
            Row row161 = new Row() { RowIndex = (UInt32Value)RowTotals, Spans = new ListValue<StringValue>() { InnerText = "1:9" }, DyDescent = 0.2D };


            Cell cell175 = new Cell() { CellReference = "D" + RowTotals, StyleIndex = (UInt32Value)164U, DataType = CellValues.String };
            CellValue cellValue133a = new CellValue("Credits:");
            cell175.Append(cellValue133a);

            Cell cell176 = new Cell() { CellReference = "E" + RowTotals, StyleIndex = (UInt32Value)210U };
            CellFormula cellFormula110 = new CellFormula();
            cellFormula110.Text = "=E" + EndIndex.ToString();
            CellValue cellValue134b = new CellValue();
            cellValue134b.Text = "0";

            cell176.Append(cellFormula110);
            cell176.Append(cellValue134b);


            row161.Append(cell175);
            row161.Append(cell176);



            RowTotals++;
            _summaryindex = ClientTotals;
            //Row row16 = new Row() { RowIndex = (UInt32Value)RowTotals, Spans = new ListValue<StringValue>() { InnerText = "1:9" }, DyDescent = 0.2D };


            //Cell cell75 = new Cell() { CellReference = "D" + RowTotals, StyleIndex = (UInt32Value)164U, DataType = CellValues.String };
            //CellValue cellValue33 = new CellValue();
            //cellValue33.Text = "On-Going Percent Savings:";

            //cell75.Append(cellValue33);

            //Cell cell76 = new Cell() { CellReference = "E" + RowTotals, StyleIndex = (UInt32Value)208U };
            //CellFormula cellFormula10 = new CellFormula();
            //cellFormula10.Text = "SUM((E" + SavingTotals + "/E" + ClientTotals + "))";
            //CellValue cellValue34 = new CellValue();
            //cellValue34.Text = "0";

            //cell76.Append(cellFormula10);
            //cell76.Append(cellValue34);


            //row16.Append(cell75);
            //row16.Append(cell76);


           // sheetData2.Append(row11a);
            sheetData2.Append(row12);
            sheetData2.Append(row13);
            sheetData2.Append(row14);
            sheetData2.Append(row15);
            sheetData2.Append(row161);
         //   sheetData2.Append(row16);


            foreach (var cell in sheetData2.Descendants<CellFormula>())
                cell.CalculateCell = true;


            MergeCells mergeCells104 = new MergeCells() { Count = (UInt32Value)4U };
            MergeCell mergeCell118 = new MergeCell() { Reference = "A1:J1" };
            MergeCell mergeCell119 = new MergeCell() { Reference = "A2:J2" };
            MergeCell mergeCell120 = new MergeCell() { Reference = "A3:J3" };
            MergeCell mergeCell121 = new MergeCell() { Reference = "A4:J4" };

            mergeCells104.Append(mergeCell118);
            mergeCells104.Append(mergeCell119);
            mergeCells104.Append(mergeCell120);
            mergeCells104.Append(mergeCell121);

            ConditionalFormatting conditionalFormatting1 = new ConditionalFormatting() { SequenceOfReferences = new ListValue<StringValue>() { InnerText = "G5" } };

            ConditionalFormattingRule conditionalFormattingRule1 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)0U, Priority = 1, StopIfTrue = true, Operator = ConditionalFormattingOperatorValues.LessThan };
            Formula formula1 = new Formula();
            formula1.Text = "0";

            conditionalFormattingRule1.Append(formula1);

            conditionalFormatting1.Append(conditionalFormattingRule1);

            PrintOptions printOptions3 = new PrintOptions() { HorizontalCentered = true };
            PageMargins pageMargins2 = new PageMargins() { Left = 0.553D, Right = 0.333D, Top = 1.25D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup2 = new PageSetup() { Orientation = OrientationValues.Landscape, HorizontalDpi = (UInt32Value)0U, VerticalDpi = (UInt32Value)0U, Id = "rId1" };

            HeaderFooter headerFooter2 = new HeaderFooter();
            OddHeader oddHeader2 = new OddHeader();
            oddHeader2.Text = "&C&G";
            OddFooter oddFooter2 = new OddFooter();
            oddFooter2.Text = "&C&G";

            headerFooter2.Append(oddHeader2);
            headerFooter2.Append(oddFooter2);
            LegacyDrawingHeaderFooter legacyDrawingHeaderFooter2 = new LegacyDrawingHeaderFooter() { Id = "rId2" };

            worksheet2.Append(sheetDimension2);
            worksheet2.Append(sheetViews2);
            worksheet2.Append(sheetFormatProperties2);
            worksheet2.Append(columns2);
            worksheet2.Append(sheetData2);
       //     worksheet2.Append(printOptions3);
            worksheet2.Append(mergeCells104);
            worksheet2.Append(conditionalFormatting1);
            worksheet2.Append(pageMargins2);
            worksheet2.Append(pageSetup2);
            worksheet2.Append(headerFooter2);
            worksheet2.Append(legacyDrawingHeaderFooter2);

            worksheetPart2.Worksheet = worksheet2;
        }
        private Row CreateContentSummaryBLU(UInt32Value index, GVWebapi.Models.VisionData VisionDataList)
        {
            Row r = new Row() {  Spans = new ListValue<StringValue>() { InnerText = "1:9" }, Height = 17.2D, CustomHeight = true, DyDescent = 0.2D };
            r.RowIndex = index;

            Cell c;
            c = new Cell() { StyleIndex = (UInt32Value)169U, DataType = CellValues.String };
            c.CellReference = "B" + index.ToString();
            c.CellValue = new CellValue(VisionDataList.ClientPeriodDates);
            r.Append(c);

            c = new Cell() { StyleIndex = (UInt32Value)75U };
            c.CellReference = "C" + index.ToString();

            CellValue cvB = new CellValue();
            cvB.Text = VisionDataList.ClientCost.ToString();
            c.Append(cvB);
            r.Append(c);

            c = new Cell() { DataType = CellValues.Number, StyleIndex = (UInt32Value)75U };
            c.CellReference = "D" + index.ToString();
            c.DataType.Value = CellValues.Number;
            c.CellValue = new CellValue(VisionDataList.ClientOverageCost.ToString());
            r.Append(c);

            c = new Cell() { StyleIndex = (UInt32Value)75U };
            c.CellReference = "E" + index.ToString();
            c.CellValue = new CellValue(VisionDataList.Credits.ToString());
            r.Append(c);

            c = new Cell() { StyleIndex = (UInt32Value)75U };
            c.CellReference = "F" + index.ToString();
            c.CellValue = new CellValue(VisionDataList.FPRCost.ToString());
            r.Append(c);

            c = new Cell() { StyleIndex = (UInt32Value)75U };

            c.CellReference = "G" + index.ToString();
            c.CellValue = new CellValue(VisionDataList.FPROverageCost.ToString());
            r.Append(c);

            c = new Cell() { StyleIndex = (UInt32Value)75U };
            c.CellReference = "H" + index.ToString();
            CellFormula cfG = new CellFormula();
            cfG.Text = "SUM((C" + index.ToString() + "+D" + index.ToString() + ")-(-E" + index.ToString() + "+F" + index.ToString() + "+G" + index.ToString() + "))";
            CellValue cv = new CellValue();
            c.Append(cfG);
            c.Append(cv);
            r.Append(c);





            c = new Cell() { StyleIndex = (UInt32Value)168U };
            c.CellReference = "I" + index.ToString();
            CellFormula cfH = new CellFormula();
            cfH.Text = "IF(G" + index.ToString() + " = \"\",\"\",SUM((C" + index.ToString() + "+D" + index.ToString() + ")-(F" + index.ToString() + "+G" + index.ToString() + "-E" + index.ToString() + "))/ (C" + index.ToString() + "+D" + index.ToString() + "))";
            CellValue cv2 = new CellValue();

            c.Append(cfH);
            r.Append(c);

            return r;
        }
        private Row CreateContentSummaryWHT(UInt32Value index, GVWebapi.Models.VisionData VisionDataList)
        {
            Row r = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:9" }, Height = 17.2D, CustomHeight = true, DyDescent = 0.2D };
            r.RowIndex = index;

            Cell c;
            c = new Cell() { StyleIndex = (UInt32Value)171U, DataType = CellValues.String };
            c.CellReference = "B" + index.ToString();
            c.CellValue = new CellValue(VisionDataList.ClientPeriodDates);
            r.Append(c);

            c = new Cell() { StyleIndex = (UInt32Value)17U };
            c.CellReference = "C" + index.ToString();

            CellValue cvB = new CellValue();
            cvB.Text = VisionDataList.ClientCost.ToString();
            c.Append(cvB);
            r.Append(c);

            c = new Cell() { DataType = CellValues.Number, StyleIndex = (UInt32Value)17U };
            c.CellReference = "D" + index.ToString();
            c.DataType.Value = CellValues.Number;
            c.CellValue = new CellValue(VisionDataList.ClientOverageCost.ToString());
            r.Append(c);

            c = new Cell() { StyleIndex = (UInt32Value)17U };
            c.CellReference = "E" + index.ToString();
            c.CellValue = new CellValue(VisionDataList.Credits.ToString());
            r.Append(c);

            c = new Cell() { StyleIndex = (UInt32Value) 17U };
            c.CellReference = "F" + index.ToString();
            c.CellValue = new CellValue(VisionDataList.FPRCost.ToString());
            r.Append(c);

            c = new Cell() { StyleIndex = (UInt32Value)17U };

            c.CellReference = "G" + index.ToString();
            c.CellValue = new CellValue(VisionDataList.FPROverageCost.ToString());
            r.Append(c);

            c = new Cell() { StyleIndex = (UInt32Value)17U };
            c.CellReference = "H" + index.ToString();
            CellFormula cfG = new CellFormula();
            cfG.Text = "SUM((C" + index.ToString() + "+D" + index.ToString() + ")-(-E" + index.ToString() + "+F" + index.ToString() + "+G" + index.ToString() + "))";
            CellValue cv = new CellValue();
            c.Append(cfG);
            c.Append(cv);
            r.Append(c);





            c = new Cell() { StyleIndex = (UInt32Value)170U };
            c.CellReference = "I" + index.ToString();
            CellFormula cfH = new CellFormula();
            cfH.Text = "IF(G" + index.ToString() + " = \"\",\"\",SUM((C" + index.ToString() + "+D" + index.ToString() + ")-(F" + index.ToString() + "+G" + index.ToString() + "-E" + index.ToString() + "))/ (C" + index.ToString() + "+D" + index.ToString() + ") )";
            CellValue cv2 = new CellValue();

            c.Append(cfH);
            r.Append(c);

            return r;
        }
        //Executive Summary
        //
        private void GenerateWorksheetPartExecutiveSummary(WorksheetPart worksheetPart3)
        {
            Worksheet worksheet3 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet3.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet3.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetProperties sheetProperties3 = new SheetProperties() { CodeName = "Sheet2" };
            SheetDimension sheetDimension3 = new SheetDimension() { Reference = "A1:I32" };

            SheetViews sheetViews3 = new SheetViews();

            SheetView sheetView3 = new SheetView() { ShowGridLines = false, ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U, View = SheetViewValues.PageLayout };
            Selection selection3 = new Selection() { ActiveCell = "B8", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "B8" } };

            sheetView3.Append(selection3);

            sheetViews3.Append(sheetView3);
            SheetFormatProperties sheetFormatProperties3 = new SheetFormatProperties() { DefaultRowHeight = 15D, DefaultColumnWidth=9.16D, DyDescent = 0.25D };

            Columns columns6 = new Columns();
            Column column32 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 11.28515625D, Style = (UInt32Value)42U, CustomWidth = true };
            Column column33 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 28.7109375D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column34 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 11.28515625D, Style = (UInt32Value)1U, CustomWidth = true };
            Column column35 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 11.7109375D, Style = (UInt32Value)12U, CustomWidth = true };
            Column column36 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 11.140625D, Style = (UInt32Value)12U, CustomWidth = true };
            Column column37 = new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)6U, Width = 10.140625D, Style = (UInt32Value)12U, CustomWidth = true };
            Column column38 = new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 10.140625D, Style = (UInt32Value)12U, CustomWidth = true };
            Column column39 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 9.140625D, Style = (UInt32Value)12U, CustomWidth = true };
            Column column40 = new Column() { Min = (UInt32Value)9U, Max = (UInt32Value)9U, Width = 11.333D, Style = (UInt32Value)12U, CustomWidth = true };

            columns6.Append(column32);
            columns6.Append(column33);
            columns6.Append(column34);
            columns6.Append(column35);
            columns6.Append(column36);
            columns6.Append(column37);
            columns6.Append(column38);
            columns6.Append(column39);
            columns6.Append(column40);
            SheetData sheetData3 = new SheetData();



            Row row103 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, Height = 15.75D, DyDescent = 0.25D };
            Cell cell680a = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)132U };
            Cell cell674 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)125U, DataType = CellValues.String };
            CellFormula cellFormula76 = new CellFormula();
            cellFormula76.Text = "SETUP!$B$2";
            CellValue cellValue129 = new CellValue();
            cellValue129.Text = "?";
            cellFormula76.CalculateCell = true;
            cell674.Append(cellFormula76);
            cell674.Append(cellValue129);
            Cell cell675 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)132U };
            Cell cell676 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)132U };
            Cell cell677 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)132U };
            Cell cell678 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)132U };
            Cell cell679 = new Cell() { CellReference = "G1", StyleIndex = (UInt32Value)132U };
            Cell cell680 = new Cell() { CellReference = "H1", StyleIndex = (UInt32Value)132U };
           
            row103.Append(cell674);
            row103.Append(cell675);
            row103.Append(cell676);
            row103.Append(cell677);
            row103.Append(cell678);
            row103.Append(cell679);
            row103.Append(cell680);
            row103.Append(cell680a);
            sheetData3.Append(row103);

            //Report Header
            Row row48 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:11" }, Height = 20.5D, CustomHeight = true, DyDescent = 0.3D };
            Cell cell347 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)131U };
            Cell cell339 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)123U, DataType = CellValues.String };
            CellFormula cellFormula77a = new CellFormula();
            DateTime StartDate = Convert.ToDateTime(_startDate);

            DateTime EndDate = Convert.ToDateTime(_period);

            cellFormula77a.Text = "=\"Managed Savings Summary From "+ StartDate.ToString("MMMMM 1, yyyy") +"  Through \"&TEXT(SETUP!$B$4,\"MMMMMMMM DD, YYYY\")&\"\"";
            //"=\"The summary below details the \"&SETUP!$B$2&\" service calls for the assets managed by FPR for the current period.  The assets are sorted by Device ID#, to highlight potentially problematic devices (in RED).  For the period reviewed, there were \"&TEXT($I$6,\"#,###\")&\" service calls.\"";
            CellValue cellValue130a = new CellValue();
            cellValue130a.Text = "=\"Managed Savings Summary From " + StartDate.ToString("MMMMM 1, yyyy") + "  Through \"&TEXT(SETUP!$B$4,\"MMMMMMMM DD, YYYY\")&\"\"";
            cellFormula77a.CalculateCell = true;
            cell339.Append(cellFormula77a);
        
            Cell cell340 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)131U };
            Cell cell341 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)131U };
            Cell cell342 = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)131U };
            Cell cell343 = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)131U };
            Cell cell344 = new Cell() { CellReference = "G2", StyleIndex = (UInt32Value)131U };
            Cell cell345 = new Cell() { CellReference = "H2", StyleIndex = (UInt32Value)131U };
            Cell cell346 = new Cell() { CellReference = "I2", StyleIndex = (UInt32Value)131U };
           
           

            row48.Append(cell339);
            row48.Append(cell340);
            row48.Append(cell341);
            row48.Append(cell342);
            row48.Append(cell343);
            row48.Append(cell344);
            row48.Append(cell345);
            row48.Append(cell346);
            row48.Append(cell347);
          

            Row row49 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:11" }, Height = 38.5D, CustomHeight = true, DyDescent = 0.3D };
            Cell cell358 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)26U };
            Cell cell350 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)199U };
            CellValue cellValue35 = new CellValue();
            CellFormula cellFormula35 = new CellFormula();
            cellFormula35.Text = "=IF(SETUP!B4=\"?\",\"\",\"Freedom Profit Recovery (FPR) is your Managed Savings Provider in the area of document output.  This report details and summarizes the overall savings delivered by FPR over the past \"&ROUND(SUM((SETUP!B4-SETUP!B3)/30.42),1)&\" months.  The areas where FPR is able to effectively deliver and accurately detail cost savings are listed below.\")";
            cellValue35.Text = "=IF(SETUP!B4=\"?\",\"\",\"Freedom Profit Recovery (FPR) is your Managed Savings Provider in the area of document output.  This report details and summarizes the overall savings delivered by FPR over the past \"&ROUND(SUM((SETUP!B4-SETUP!B3)/30.42),1)&\" months.  The areas where FPR is able to effectively deliver and accurately detail cost savings are listed below.\")";
            cellFormula35.CalculateCell = true;
            cell350.Append(cellFormula35);
            cell350.Append(cellValue35);
        
            Cell cell351 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)25U };
            Cell cell352 = new Cell() { CellReference = "D3", StyleIndex = (UInt32Value)26U };
            Cell cell353 = new Cell() { CellReference = "E3", StyleIndex = (UInt32Value)26U };
            Cell cell354 = new Cell() { CellReference = "F3", StyleIndex = (UInt32Value)26U };
            Cell cell355 = new Cell() { CellReference = "G3", StyleIndex = (UInt32Value)26U };
            Cell cell356 = new Cell() { CellReference = "H3", StyleIndex = (UInt32Value)26U };
            Cell cell357 = new Cell() { CellReference = "I3", StyleIndex = (UInt32Value)26U };
          
      
            row49.Append(cell350);
            row49.Append(cell351);
            row49.Append(cell352);
            row49.Append(cell353);
            row49.Append(cell354);
            row49.Append(cell355);
            row49.Append(cell356);
            row49.Append(cell357);
            row49.Append(cell358);
 

            Row row49a = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:11" },   DyDescent = 0.3D };
            Cell cell350a = new Cell() { CellReference = "A4", StyleIndex = (UInt32Value)127U };
            Cell cell351a = new Cell() { CellReference = "B4", StyleIndex = (UInt32Value)25U };
            Cell cell352a = new Cell() { CellReference = "C4", StyleIndex = (UInt32Value)26U };
            Cell cell353a = new Cell() { CellReference = "D4", StyleIndex = (UInt32Value)26U };
            Cell cell354a = new Cell() { CellReference = "E4", StyleIndex = (UInt32Value)26U };
            Cell cell355a = new Cell() { CellReference = "F4", StyleIndex = (UInt32Value)26U };
            Cell cell356a = new Cell() { CellReference = "G4", StyleIndex = (UInt32Value)26U };
            Cell cell357a = new Cell() { CellReference = "H4", StyleIndex = (UInt32Value)26U };
            Cell cell358a = new Cell() { CellReference = "I4", StyleIndex = (UInt32Value)26U };
 
            row49a.Append(cell350a);
            row49a.Append(cell351a);
            row49a.Append(cell352a);
            row49a.Append(cell353a);
            row49a.Append(cell354a);
            row49a.Append(cell355a);
            row49a.Append(cell356a);
            row49a.Append(cell357a);
            row49a.Append(cell358a);
         
            

            Row row49b = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:11" },   DyDescent = 0.3D };
            Cell cell350b = new Cell() { CellReference = "A5", StyleIndex = (UInt32Value)127U };
            Cell cell351b = new Cell() { CellReference = "B5", StyleIndex = (UInt32Value)25U };
            Cell cell352b = new Cell() { CellReference = "C5", StyleIndex = (UInt32Value)26U };
            Cell cell353b = new Cell() { CellReference = "D5", StyleIndex = (UInt32Value)164U };
            CellValue assetReplacement = new CellValue();
            assetReplacement.Text = "Asset Replacement Savings";
            cell353b.Append(assetReplacement);
            Cell cell354b = new Cell() { CellReference = "E5", StyleIndex = (UInt32Value)204U };
            CellFormula ReplacementTotal = new CellFormula();
            ReplacementTotal.Text = _replacementTotalsString;
            cell354b.Append(ReplacementTotal);

            Cell cell355b = new Cell() { CellReference = "F5", StyleIndex = (UInt32Value)26U };
            Cell cell356b = new Cell() { CellReference = "G5", StyleIndex = (UInt32Value)26U };
            Cell cell357b = new Cell() { CellReference = "H5", StyleIndex = (UInt32Value)26U };
            Cell cell358b = new Cell() { CellReference = "I5", StyleIndex = (UInt32Value)26U };
 
            row49b.Append(cell350b);
            row49b.Append(cell351b);
            row49b.Append(cell352b);
            row49b.Append(cell353b);
            row49b.Append(cell354b);
            row49b.Append(cell355b);
            row49b.Append(cell356b);
            row49b.Append(cell357b);
            row49b.Append(cell358b);
 
            Row row49c = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:11" }, DyDescent = 0.3D };
            Cell cell350c = new Cell() { CellReference = "A6", StyleIndex = (UInt32Value)127U };
            Cell cell351c = new Cell() { CellReference = "B6", StyleIndex = (UInt32Value)25U };
            Cell cell352c = new Cell() { CellReference = "C6", StyleIndex = (UInt32Value)26U };
            Cell cell353c = new Cell() { CellReference = "D6", StyleIndex = (UInt32Value)164U };
            CellValue CostAvoidance = new CellValue();
            CostAvoidance.Text = "Cost Avoidance";
            cell353c.Append(CostAvoidance);
            Cell cell354c = new Cell() { CellReference = "E6", StyleIndex = (UInt32Value)204U };
            CellFormula CostAvoidanceTotal = new CellFormula();
            CostAvoidanceTotal.Text = _costavoidanceTotalString;
            cell354c.Append(CostAvoidanceTotal);
            Cell cell355c = new Cell() { CellReference = "F6", StyleIndex = (UInt32Value)26U };
            Cell cell356c = new Cell() { CellReference = "G6", StyleIndex = (UInt32Value)26U };
            Cell cell357c = new Cell() { CellReference = "H6", StyleIndex = (UInt32Value)26U };
            Cell cell358c = new Cell() { CellReference = "I6", StyleIndex = (UInt32Value)26U };
 
            row49c.Append(cell350c);
            row49c.Append(cell351c);
            row49c.Append(cell352c);
            row49c.Append(cell353c);
            row49c.Append(cell354c);
            row49c.Append(cell355c);
            row49c.Append(cell356c);
            row49c.Append(cell357c);
            row49c.Append(cell358c);
 

            Row row49d = new Row() { RowIndex = (UInt32Value)7U, Spans = new ListValue<StringValue>() { InnerText = "1:11" }, DyDescent = 0.3D };
            Cell cell350d = new Cell() { CellReference = "A7", StyleIndex = (UInt32Value)127U };
            Cell cell351d = new Cell() { CellReference = "B7", StyleIndex = (UInt32Value)25U };
            Cell cell352d = new Cell() { CellReference = "C7", StyleIndex = (UInt32Value)26U };
            Cell cell353d = new Cell() { CellReference = "D7", StyleIndex = (UInt32Value)164U };
            CellValue RolloverSavings = new CellValue();
            RolloverSavings.Text = "Rollover Savings";
            cell353d.Append(RolloverSavings);
            Cell cell354d = new Cell() { CellReference = "E7", StyleIndex = (UInt32Value)204U , DataType = CellValues.Number};
            CellValue RolloverFormula = new CellValue();
            RolloverFormula.Text = _rolloverTotalString;
            cell354d.Append(RolloverFormula);
            Cell cell355d = new Cell() { CellReference = "F7", StyleIndex = (UInt32Value)26U };
            Cell cell356d = new Cell() { CellReference = "G7", StyleIndex = (UInt32Value)26U };
            Cell cell357d = new Cell() { CellReference = "H7", StyleIndex = (UInt32Value)26U };
            Cell cell358d = new Cell() { CellReference = "I7", StyleIndex = (UInt32Value)26U };
            Cell cell359d = new Cell() { CellReference = "J7", StyleIndex = (UInt32Value)21U };

            row49d.Append(cell350d);
            row49d.Append(cell351d);
            row49d.Append(cell352d);
            row49d.Append(cell353d);
            row49d.Append(cell354d);
            row49d.Append(cell355d);
            row49d.Append(cell356d);
            row49d.Append(cell357d);
            row49d.Append(cell358d);
            row49d.Append(cell359d);


            Row row49e = new Row() { RowIndex = (UInt32Value)8U, Spans = new ListValue<StringValue>() { InnerText = "1:11" }, DyDescent = 0.3D };
            Cell cell350e = new Cell() { CellReference = "A8", StyleIndex = (UInt32Value)127U };
            Cell cell351e = new Cell() { CellReference = "B8", StyleIndex = (UInt32Value)25U };
            Cell cell352e = new Cell() { CellReference = "C8", StyleIndex = (UInt32Value)26U };
            Cell cell353e = new Cell() { CellReference = "D8", StyleIndex = (UInt32Value)164U };
            CellValue RevisionSavings = new CellValue();
            RevisionSavings.Text = "Revision Savings";
            cell353e.Append(RevisionSavings);
            Cell cell354e = new Cell() { CellReference = "E8", StyleIndex = (UInt32Value)204U };
            CellFormula RevisionFormula = new CellFormula();
            RevisionFormula.Text = _revisionTotalString;
            RevisionFormula.CalculateCell = true;
            cell354e.Append(RevisionFormula);
            Cell cell355e = new Cell() { CellReference = "F8", StyleIndex = (UInt32Value)26U };
            Cell cell356e = new Cell() { CellReference = "G8", StyleIndex = (UInt32Value)26U };
            Cell cell357e = new Cell() { CellReference = "H8", StyleIndex = (UInt32Value)26U };
            Cell cell358e = new Cell() { CellReference = "I8", StyleIndex = (UInt32Value)26U };
 
            row49e.Append(cell350e);
            row49e.Append(cell351e);
            row49e.Append(cell352e);
            row49e.Append(cell353e);
            row49e.Append(cell354e);
            row49e.Append(cell355e);
            row49e.Append(cell356e);
            row49e.Append(cell357e);
            row49e.Append(cell358e);
 

            Row row49f = new Row() { RowIndex = (UInt32Value)9U, Spans = new ListValue<StringValue>() { InnerText = "1:11" }, DyDescent = 0.3D };
            Cell cell350f = new Cell() { CellReference = "A9", StyleIndex = (UInt32Value)127U };
            Cell cell351f = new Cell() { CellReference = "B9", StyleIndex = (UInt32Value)25U };
            Cell cell352f = new Cell() { CellReference = "C9", StyleIndex = (UInt32Value)26U };
            Cell cell353f = new Cell() { CellReference = "D9", StyleIndex = (UInt32Value)164U };
            CellValue HardDollarSavings = new CellValue();
            HardDollarSavings.Text = "Total Minimum Hard-Dollar Savings";
            cell353f.Append(HardDollarSavings);
            Cell cell354f = new Cell() { CellReference = "E9", StyleIndex = (UInt32Value)204U };
            CellFormula HardDollarFormula = new CellFormula();
            HardDollarFormula.Text = "=SUM(E5:E8)";
            HardDollarFormula.CalculateCell = true;
            cell354f.Append(HardDollarFormula);
            Cell cell355f = new Cell() { CellReference = "F9", StyleIndex = (UInt32Value)208U };
            CellFormula percentofsavings = new CellFormula();
            percentofsavings.Text = "=SUM(E9/'Revision Summary'!E" + _summaryindex + ")";
            percentofsavings.CalculateCell = true;
            cell355f.Append(percentofsavings);
            Cell cell356f = new Cell() { CellReference = "G9", StyleIndex = (UInt32Value)26U };
            Cell cell357f = new Cell() { CellReference = "H9", StyleIndex = (UInt32Value)26U };
            Cell cell358f = new Cell() { CellReference = "I9", StyleIndex = (UInt32Value)26U };
 
            row49f.Append(cell350f);
            row49f.Append(cell351f);
            row49f.Append(cell352f);
            row49f.Append(cell353f);
            row49f.Append(cell354f);
            row49f.Append(cell355f);
            row49f.Append(cell356f);
            row49f.Append(cell357f);
            row49f.Append(cell358f);
 

           


            sheetData3.Append(row48);
            sheetData3.Append(row49);
            sheetData3.Append(row49a);
            sheetData3.Append(row49b);
            sheetData3.Append(row49c);
            sheetData3.Append(row49d);
            sheetData3.Append(row49e);
            sheetData3.Append(row49f);

            UInt32Value _rowindex = 10;

            int _even = 1;
            int _invoiceID = 0;
            String PeriodChange = "";

            DateTime period = Convert.ToDateTime(_period);
            CoFreedomEntities db = new CoFreedomEntities();
            GlobalViewEntities gv = new GlobalViewEntities();
            
            var gvdata = gv.RevisionDataViews.Where(r => r.ContractID == _contractID && r.OverageToDate.Value == period).ToList().Distinct();
            var query = (from  r in gvdata   
                         orderby r.MeterGroup
                         select new RevisionDataView
                         {
                            OverageFromDate = r.OverageFromDate,
                            OverageToDate = r.OverageToDate,
                            MeterGroup = r.MeterGroup,
                            ContractVolume = r.ContractVolume,
                            ActualVolume = r.ActualVolume,
                            CPP = r.CPP,
                            
                            CreditAmount = r.CreditAmount,
                            Rollover = r.Rollover.Value
                         }).ToList();
 

           
            _rowindex++;
            foreach (var item in query)
            {

                _rowindex++;
                if (PeriodChange != item.OverageToDate.Value.ToShortDateString())
                {
                    Row row4 = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 25.5D, CustomHeight = true, DyDescent = 0.25D };
                    Cell cell31 = new Cell() { CellReference = "A" + _rowindex.ToString(), StyleIndex = (UInt32Value)19U };


                    Cell cell32 = new Cell() { CellReference = "B" + _rowindex.ToString() };
                    cell32.StyleIndex = _invoiceID == 0 ? 19U : 162U;
                    CellValue cellValuePeriod = new CellValue();
                    cellValuePeriod.Text = "From " + item.OverageFromDate.Value.ToString("MMM. yyyy") + " to " + item.OverageToDate.Value.ToString("MMM. yyyy");
                    cell32.Append(cellValuePeriod);
                    Cell cell33 = new Cell() { CellReference = "C" + _rowindex.ToString() };
                    cell33.StyleIndex = _invoiceID == 0 ? 19U : 160U;
                    Cell cell34 = new Cell() { CellReference = "D" + _rowindex.ToString() };
                    cell34.StyleIndex = _invoiceID == 0 ? 19U : 160U;
                    Cell cell35 = new Cell() { CellReference = "E" + _rowindex.ToString() };
                    cell35.StyleIndex = _invoiceID == 0 ? 19U : 160U;
                    Cell cell36 = new Cell() { CellReference = "F" + _rowindex.ToString() };
                    cell36.StyleIndex = _invoiceID == 0 ? 19U : 160U;
                    Cell cell37 = new Cell() { CellReference = "G" + _rowindex.ToString() };
                    cell37.StyleIndex = _invoiceID == 0 ? 19U : 160U;
                    Cell cell38 = new Cell() { CellReference = "H" + _rowindex.ToString() };
                    cell38.StyleIndex = _invoiceID == 0 ? 19U : 160U;
                    Cell cell39 = new Cell() { CellReference = "I" + _rowindex.ToString() };
                    cell39.StyleIndex = _invoiceID == 0 ? 19U : 160U;
                    Cell cell39b = new Cell() { CellReference = "J" + _rowindex.ToString() };
                    cell39b.StyleIndex = _invoiceID == 0 ? 19U : 160U;
                    row4.Append(cell31);
                    row4.Append(cell32);
                    row4.Append(cell33);
                    row4.Append(cell34);
                    row4.Append(cell35);
                    row4.Append(cell36);
                    row4.Append(cell37);
                    row4.Append(cell38);
                    row4.Append(cell39);
                    row4.Append(cell39b);
                    _rowindex++;

                    Row row6 = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, Height = 35.75D, CustomHeight = true, DyDescent = 0.25D };


                    Cell cell42 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)47U, DataType = CellValues.String };
                    CellValue cellValue6 = new CellValue();
                    cellValue6.Text = "Meter Group";

                    cell42.Append(cellValue6);

                    Cell cell43 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                    CellValue cellValue7 = new CellValue();
                    cellValue7.Text = "Quarterly Contracted Volume";

                    cell43.Append(cellValue7);

                    Cell cell44 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                    CellValue cellValue8 = new CellValue();
                    cellValue8.Text = "Actual Volume For Period";

                    cell44.Append(cellValue8);

                    Cell cell45 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                    CellValue cellValue9 = new CellValue();
                    cellValue9.Text = "Overage/ (Underage) For Period";

                    cell45.Append(cellValue9);

                    Cell cell46 = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                    CellValue cellValue10 = new CellValue();
                    cellValue10.Text = "Accrued Rollover Pages";

                    cell46.Append(cellValue10);

                    Cell cell47 = new Cell() { CellReference = "G" + _rowindex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                    CellValue cellValue11 = new CellValue();
                    cellValue11.Text = "Net Overage (Underage)";

                    cell47.Append(cellValue11);

                    Cell cell48 = new Cell() { CellReference = "H" + _rowindex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                    CellValue cellValue12 = new CellValue();
                    cellValue12.Text = "Excess CPP";

                    cell48.Append(cellValue12);

                    Cell cell48a = new Cell() { CellReference = "I" + _rowindex.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
                    CellValue cellValue12a = new CellValue();
                    cellValue12a.Text = "Credits";

                    cell48a.Append(cellValue12a);

                    Cell cell48b = new Cell() { CellReference = "J" + _rowindex.ToString(), StyleIndex = (UInt32Value)46U, DataType = CellValues.String };
                    CellValue cellValue12b = new CellValue();
                    cellValue12b.Text = "Overage Expense";

                    cell48b.Append(cellValue12b);

                    row6.Append(cell42);
                    row6.Append(cell43);
                    row6.Append(cell44);
                    row6.Append(cell45);
                    row6.Append(cell46);
                    row6.Append(cell47);
                    row6.Append(cell48);
                    row6.Append(cell48a);
                    row6.Append(cell48b);

                    sheetData3.Append(row4);
                    sheetData3.Append(row6);



                    PeriodChange = item.OverageToDate.Value.ToShortDateString();
                    _rowindex++;
                }

                if (_even <= 2)
                {
                    Row row7 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };
                    row7.RowIndex = _rowindex;
                    Cell cell49 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)60U, DataType = CellValues.String };
                    CellValue cellValue13 = new CellValue();
                    cellValue13.Text = item.MeterGroup;

                    cell49.Append(cellValue13);
                    Cell cell50 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)73U, DataType = CellValues.String };
                    CellValue SavingsTypeValue = new CellValue();
                    SavingsTypeValue.Text = item.ContractVolume.Value.ToString("#,##0");
                    cell50.Append(SavingsTypeValue);

                    Cell cell51 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)73U, DataType = CellValues.String };
                    CellValue MonthsValues = new CellValue();
                    MonthsValues.Text = item.ActualVolume.Value.ToString("#,##0");
                    cell51.Append(MonthsValues);

                    Cell cell52 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)177U, DataType = CellValues.Number };
                    CellFormula Overage = new CellFormula();
                    // Overage.Text = "=IF(SUM(C" + _rowindex.ToString() + " - D" + _rowindex.ToString() + ") < 0 ,0,SUM(C" + _rowindex.ToString() + " - D" + _rowindex.ToString() + "))";
                    Overage.Text = "=SUM(D" + _rowindex.ToString() + " - C" + _rowindex.ToString() + ")";
                    Overage.CalculateCell = true;
                    cell52.Append(Overage);

                    Cell cell53 = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)177U, DataType = CellValues.Number };
                    CellValue EndDateValue = new CellValue();
                    EndDateValue.Text = item.Rollover.ToString();
                    cell53.Append(EndDateValue);

                    Cell cell54 = new Cell() { CellReference = "G" + _rowindex.ToString(), StyleIndex = (UInt32Value)177U, DataType = CellValues.Number };
                    CellFormula CommentValue = new CellFormula();
                    CommentValue.Text = "=SUM(E"+_rowindex.ToString()+"-F"+_rowindex.ToString()+")";
                    CommentValue.CalculateCell = true;
                    cell54.Append(CommentValue);
                    Cell cell55 = new Cell() { CellReference = "H" + _rowindex.ToString(), StyleIndex = (UInt32Value)75U, DataType = CellValues.String };
                    CellValue CostSavingsAmountValue = new CellValue();
                    CostSavingsAmountValue.Text = item.CPP.Value.ToString("$ 0.0000");
                    cell55.Append(CostSavingsAmountValue);
                    Cell cell55a = new Cell() { CellReference = "I" + _rowindex.ToString(), StyleIndex = (UInt32Value)218U, DataType = CellValues.Number };
                    CellValue CreditsAmountValue = new CellValue();
                    CreditsAmountValue.Text = item.CreditAmount.ToString();
                    cell55a.Append(CreditsAmountValue);
                    Cell cell56 = new Cell() { CellReference = "J" + _rowindex.ToString(), StyleIndex = (UInt32Value)203U, DataType = CellValues.Number };
                    CellFormula OverageCharge = new CellFormula();
                    OverageCharge.Text = "=IF(I" + _rowindex.ToString() + "<>0,SUM(-1*I" + _rowindex.ToString() + "),IF(SUM(G" + _rowindex.ToString() + " * H" + _rowindex.ToString() + ") < 0, 0,SUM(G" + _rowindex.ToString() + " * H" + _rowindex.ToString() + ")))";
                    OverageCharge.CalculateCell = true;
                    cell56.Append(OverageCharge);

                    row7.Append(cell49);
                    row7.Append(cell50);
                    row7.Append(cell51);
                    row7.Append(cell52);
                    row7.Append(cell53);
                    row7.Append(cell54);
                    row7.Append(cell55);
                    row7.Append(cell55a);
                    row7.Append(cell56);
                    sheetData3.Append(row7);
                }

                if (_even > 2 && _even < 5)
                {

                    Row row7 = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };
                    row7.RowIndex = _rowindex;
                    Cell cell49 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)58U, DataType = CellValues.String };
                    CellValue cellValue13 = new CellValue();
                    cellValue13.Text = item.MeterGroup;

                    cell49.Append(cellValue13);
                    Cell cell50 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)6U, DataType = CellValues.String };
                    CellValue SavingsTypeValue = new CellValue();
                    SavingsTypeValue.Text = item.ContractVolume.Value.ToString("#,##0");
                    cell50.Append(SavingsTypeValue);

                    Cell cell51 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)6U, DataType = CellValues.String };
                    CellValue MonthsValues = new CellValue();
                    MonthsValues.Text = item.ActualVolume.Value.ToString("#,##0");
                    cell51.Append(MonthsValues);

                    Cell cell52 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)176U, DataType = CellValues.Number };
                    CellFormula Overage = new CellFormula();
                    // Overage.Text = "=IF(SUM(C" + _rowindex.ToString() + " - D" + _rowindex.ToString() + ") < 0 ,0,SUM(C" + _rowindex.ToString() + " - D" + _rowindex.ToString() + "))";
                    Overage.Text = "=SUM(D" + _rowindex.ToString() + " - C" + _rowindex.ToString() + ")";
                    Overage.CalculateCell = true;
                    cell52.Append(Overage);
                    Cell cell53 = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)176U, DataType = CellValues.Number };
                    CellValue EndDateValue = new CellValue();
                    EndDateValue.Text = item.Rollover.ToString();
                    cell53.Append(EndDateValue);

                    Cell cell54 = new Cell() { CellReference = "G" + _rowindex.ToString(), StyleIndex = (UInt32Value)176U, DataType = CellValues.Number };
                    CellFormula CommentValue = new CellFormula();
                    CommentValue.Text = "=SUM(E" + _rowindex.ToString() + " - F" + _rowindex.ToString() + ")";
                    CommentValue.CalculateCell = true;
                    cell54.Append(CommentValue);

                    Cell cell55 = new Cell() { CellReference = "H" + _rowindex.ToString(), StyleIndex = (UInt32Value)17U, DataType = CellValues.String };
                    CellValue CostSavingsAmountValue = new CellValue();
                    CostSavingsAmountValue.Text = item.CPP.Value.ToString("$ 0.0000");
                    cell55.Append(CostSavingsAmountValue);
                    
                    Cell cell55a = new Cell() { CellReference = "I" + _rowindex.ToString(), StyleIndex = (UInt32Value)217U, DataType = CellValues.Number };
                    CellValue CreditsAmountValue = new CellValue();
                    CreditsAmountValue.Text = item.CreditAmount.ToString();
                    cell55a.Append(CreditsAmountValue);

                    Cell cell56 = new Cell() { CellReference = "J" + _rowindex.ToString(), StyleIndex = (UInt32Value)202U, DataType = CellValues.Number };
                    CellFormula OverageCharge = new CellFormula();
                    OverageCharge.Text = "=IF(I" + _rowindex.ToString() + "<>0,SUM(-1*I" + _rowindex.ToString() + "),IF(SUM(G" + _rowindex.ToString() + " * H" + _rowindex.ToString() + ") < 0, 0,SUM(G" + _rowindex.ToString() + " * H" + _rowindex.ToString() + ")))";
                    OverageCharge.CalculateCell = true;
                    cell56.Append(OverageCharge);

                    row7.Append(cell49);
                    row7.Append(cell50);
                    row7.Append(cell51);
                    row7.Append(cell52);
                    row7.Append(cell53);
                    row7.Append(cell54);
                    row7.Append(cell55);
                    row7.Append(cell55a);
                    row7.Append(cell56);
                    sheetData3.Append(row7);


                }
                if (_even == 4) _even = 1; else _even++;

            }
            _rowindex++;
            Row row127 = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };
          
            Cell cell838 = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)92U };
            Cell cell839 = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell840 = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell841 = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell843 = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell844 = new Cell() { CellReference = "G" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell845 = new Cell() { CellReference = "H" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell846 = new Cell() { CellReference = "I" + _rowindex.ToString(), StyleIndex = (UInt32Value)120U };
            Cell cell847 = new Cell() { CellReference = "J" + _rowindex.ToString(), StyleIndex = (UInt32Value)121U };
            row127.Append(cell838);
            row127.Append(cell839);
            row127.Append(cell840);
            row127.Append(cell841);
            row127.Append(cell843);
            row127.Append(cell844);
            row127.Append(cell845);
            row127.Append(cell846);
            row127.Append(cell847);
            sheetData3.Append(row127);
            _rowindex++;
            Row row127a = new Row() { RowIndex = (UInt32Value)_rowindex, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, DyDescent = 0.25D };
            Cell cell837a = new Cell() { CellReference = "B" + _rowindex.ToString(), StyleIndex = (UInt32Value)164U };
            Cell cell838a = new Cell() { CellReference = "C" + _rowindex.ToString(), StyleIndex = (UInt32Value)164U };
            Cell cell839a = new Cell() { CellReference = "D" + _rowindex.ToString(), StyleIndex = (UInt32Value)164U };
            Cell cell840a = new Cell() { CellReference = "E" + _rowindex.ToString(), StyleIndex = (UInt32Value)164U };
            Cell cell841a = new Cell() { CellReference = "F" + _rowindex.ToString(), StyleIndex = (UInt32Value)164U };
            Cell cell842a = new Cell() { CellReference = "G" + _rowindex.ToString(), StyleIndex = (UInt32Value)164U };
            Cell cell843a = new Cell() { CellReference = "H" + _rowindex.ToString(), StyleIndex = (UInt32Value)164U };
            CellValue TotalLabel = new CellValue();
            TotalLabel.Text = "Current Overage Expense:";
            cell843a.Append(TotalLabel);
            Cell cell844a = new Cell() { CellReference = "J" + _rowindex.ToString(), StyleIndex = (UInt32Value)204U };
            CellFormula cellFormula52a = new CellFormula();
            UInt32Value EndRow = _rowindex - 2;
            cellFormula52a.CalculateCell = true;
            cellFormula52a.Text = "=Sum(J14:J" + EndRow + ")";
            cell844a.Append(cellFormula52a);
            row127a.Append(cell837a);
            row127a.Append(cell838a);
            row127a.Append(cell839a);
            row127a.Append(cell840a);
            row127a.Append(cell841a);
            row127a.Append(cell842a);
            row127a.Append(cell843a);
            row127a.Append(cell844a);
            sheetData3.Append(row127a);

            PrintOptions printOptions3 = new PrintOptions() { HorizontalCentered = true };
            PageMargins pageMargins3 = new PageMargins() { Left = 0.333D, Right = 0.333D, Top = 1.25D, Bottom = 0.75D, Header = 0.3D, Footer = 0.2D };
            PageSetup pageSetup3 = new PageSetup() { Orientation = OrientationValues.Landscape, HorizontalDpi = (UInt32Value)0U, VerticalDpi = (UInt32Value)0U, Id = "rId1" };


            MergeCells mergeCells105 = new MergeCells() { Count = (UInt32Value)3U };
            MergeCell mergeCell118a = new MergeCell() { Reference = "A1:J1" };
            MergeCell mergeCell119a = new MergeCell() { Reference = "A2:J2" };
            MergeCell mergeCell120a = new MergeCell() { Reference = "A3:J3" };


            mergeCells105.Append(mergeCell118a);
            mergeCells105.Append(mergeCell119a);
            mergeCells105.Append(mergeCell120a);



            HeaderFooter headerFooter3 = new HeaderFooter();
            OddHeader oddHeader3 = new OddHeader();
            oddHeader3.Text = "&C&G";
            OddFooter oddFooter3 = new OddFooter();
            oddFooter3.Text = "&C&G";

            headerFooter3.Append(oddHeader3);
            headerFooter3.Append(oddFooter3);
            LegacyDrawingHeaderFooter legacyDrawingHeaderFooter3 = new LegacyDrawingHeaderFooter() { Id = "rId2" };

            worksheet3.Append(sheetDimension3);
            worksheet3.Append(sheetViews3);
            worksheet3.Append(sheetFormatProperties3);
            worksheet3.Append(columns6);
            worksheet3.Append(sheetData3);
            worksheet3.Append(mergeCells105);  
            worksheet3.Append(printOptions3);
            worksheet3.Append(pageMargins3);
            worksheet3.Append(pageSetup3);
            worksheet3.Append(headerFooter3);
            worksheet3.Append(legacyDrawingHeaderFooter3);

            worksheetPart3.Worksheet = worksheet3;
        }
        //Revision History
        //
        private void GenerateWorksheetPartRevisionHistory(WorksheetPart worksheetPart3)
        {
            Worksheet worksheet3 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet3.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet3.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetDimension sheetDimension3 = new SheetDimension() { Reference = "A1:K34" };

            SheetViews sheetViews3 = new SheetViews();

            SheetView sheetView3 = new SheetView() { ShowGridLines = false, View = SheetViewValues.PageLayout, ZoomScaleNormal = (UInt32Value)100U, WorkbookViewId = (UInt32Value)0U };
            Selection selection3 = new Selection() { ActiveCell = "A6", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A6" } };

            sheetView3.Append(selection3);

            sheetViews3.Append(sheetView3);
            SheetFormatProperties sheetFormatProperties3 = new SheetFormatProperties() { DefaultRowHeight = 12.75D, DyDescent = 0.2D };

            Columns columns3 = new Columns();
            Columns columns6 = new Columns();
            Column column9 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 23.2D, BestFit = true, CustomWidth = true };
            Column column10 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 10.3D, BestFit = true, CustomWidth = true };
            Column column11 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 8.3D, BestFit = true, CustomWidth = true };
            Column column12 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 8.3D, BestFit = true, CustomWidth = true };
            Column column13 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 8.3D, BestFit = true, CustomWidth = true };
            Column column14 = new Column() { Min = (UInt32Value)6U, Max = (UInt32Value)6U, Width = 10.3D, BestFit = true, CustomWidth = true };
            Column column15 = new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 10.3D, BestFit = true, CustomWidth = true };
            Column column16 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 10.3D, BestFit = true, CustomWidth = true };
            columns3.Append(column9);
            columns3.Append(column10);
            columns6.Append(column11);
            columns6.Append(column12);
            columns6.Append(column13);
            columns6.Append(column14);
            columns6.Append(column15);
            columns6.Append(column16);

            SheetData sheetData3 = new SheetData();



            Row row103 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:7" }, Height = 15.75D, DyDescent = 0.25D };

            Cell cell674 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)125U, DataType = CellValues.String };
            CellFormula cellFormula76 = new CellFormula();
            cellFormula76.Text = "SETUP!$B$2";
            CellValue cellValue129 = new CellValue();
            cellValue129.Text = "?";
            cellFormula76.CalculateCell = true;
            cell674.Append(cellFormula76);
            cell674.Append(cellValue129);
            Cell cell675 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)132U };
            Cell cell676 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)132U };
            Cell cell677 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)132U };
            Cell cell678 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)132U };
            Cell cell679 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)132U };
            Cell cell680 = new Cell() { CellReference = "G1", StyleIndex = (UInt32Value)132U };
            Cell cell680a = new Cell() { CellReference = "H1", StyleIndex = (UInt32Value)132U };
            row103.Append(cell674);
            row103.Append(cell675);
            row103.Append(cell676);
            row103.Append(cell677);
            row103.Append(cell678);
            row103.Append(cell679);
            row103.Append(cell680);
            row103.Append(cell680a);
            sheetData3.Append(row103);

            //Report Header
            Row row48 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:11" }, Height = 20.5D, CustomHeight = true, DyDescent = 0.3D };

            Cell cell339 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)123U, DataType = CellValues.String };
            CellFormula cellFormula77a = new CellFormula();


            cellFormula77a.Text = "=\"REVision History Through Period Ending \"&TEXT(SETUP!$B$4,\"MMMMMMMM DD, YYYY\")&\"\"";
            //"=\"The summary below details the \"&SETUP!$B$2&\" service calls for the assets managed by FPR for the current period.  The assets are sorted by Device ID#, to highlight potentially problematic devices (in RED).  For the period reviewed, there were \"&TEXT($I$6,\"#,###\")&\" service calls.\"";
            CellValue cellValue130a = new CellValue();
            cellValue130a.Text = "=\"REVision Summary Through  Period Ending \"&TEXT(SETUP!$B$4,\"MMMMMMMM DD, YYYY\")&\"\"";
            cellFormula77a.CalculateCell = true;
            cell339.Append(cellFormula77a);
            cell339.Append(cellValue130a);
            Cell cell340 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)131U };
            Cell cell341 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)131U };
            Cell cell342 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)131U };
            Cell cell343 = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)131U };
            Cell cell344 = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)131U };
            Cell cell345 = new Cell() { CellReference = "G2", StyleIndex = (UInt32Value)131U };
            Cell cell346 = new Cell() { CellReference = "H2", StyleIndex = (UInt32Value)131U };
            Cell cell347 = new Cell() { CellReference = "I2", StyleIndex = (UInt32Value)131U };
            Cell cell348 = new Cell() { CellReference = "J2", StyleIndex = (UInt32Value)131U };
            Cell cell349 = new Cell() { CellReference = "K2", StyleIndex = (UInt32Value)131U };

            row48.Append(cell339);
            row48.Append(cell340);
            row48.Append(cell341);
            row48.Append(cell342);
            row48.Append(cell343);
            row48.Append(cell344);
            row48.Append(cell345);
            row48.Append(cell346);
            row48.Append(cell347);
            row48.Append(cell348);
            row48.Append(cell349);

            Row row49 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:11" }, Height = 15D, DyDescent = 0.3D };
            Cell cell350 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)19U };
            CellValue cellValue35 = new CellValue("Overage Volume History");
            cell350.Append(cellValue35);
            Cell cell351 = new Cell() { CellReference = "B3", StyleIndex = (UInt32Value)25U };
            Cell cell352 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)26U };
            Cell cell353 = new Cell() { CellReference = "D3", StyleIndex = (UInt32Value)26U };
            Cell cell354 = new Cell() { CellReference = "E3", StyleIndex = (UInt32Value)26U };
            Cell cell355 = new Cell() { CellReference = "F3", StyleIndex = (UInt32Value)26U };
            Cell cell356 = new Cell() { CellReference = "G3", StyleIndex = (UInt32Value)26U };
            Cell cell357 = new Cell() { CellReference = "H3", StyleIndex = (UInt32Value)26U };
            Cell cell358 = new Cell() { CellReference = "I3", StyleIndex = (UInt32Value)26U };
            Cell cell359 = new Cell() { CellReference = "J3", StyleIndex = (UInt32Value)21U };
            Cell cell360 = new Cell() { CellReference = "K3", StyleIndex = (UInt32Value)21U };

            row49.Append(cell350);
            row49.Append(cell351);
            row49.Append(cell352);
            row49.Append(cell353);
            row49.Append(cell354);
            row49.Append(cell355);
            row49.Append(cell356);
            row49.Append(cell357);
            row49.Append(cell358);
            row49.Append(cell359);
            row49.Append(cell360);

            sheetData3.Append(row48);
            //    sheetData3.Append(row49);

            ExcelRevisionExport export = new ExcelRevisionExport();

            DateTime perioddate = Convert.ToDateTime(_period);
            DateTime overridedate = Convert.ToDateTime(_overrideDate);
            var RevisionHistory = export.GetRevisionHistory(_contractID).Where(o => o.peroid >= overridedate && o.peroid <= perioddate).ToList();
            uint RowIndex = 3;
            int footerIndex = 1;
            int Count = 0;
            int _even = 0;
            DateTime CurrentPeriod =  new DateTime();
            foreach (var revision in RevisionHistory)
            {
                 var revisiondetail = revision.detail;
                
                RowIndex++;
                CurrentPeriod = revision.peroid.Value;
                sheetData3.Append(CreateHeader(RowIndex, CurrentPeriod.ToShortDateString()));
                RowIndex++;
                sheetData3.Append(CreateHeaderLabels(RowIndex));
                RowIndex++;
                var StartIndex = RowIndex;
                foreach (var detail in revisiondetail)
                {
                    if (_even == 0)
                    {
                        sheetData3.Append(CreateContentBLU(RowIndex, detail));
                        _even = 1;

                    }
                    else
                    {
                        sheetData3.Append(CreateContentWHT(RowIndex, detail));

                        _even = 0;
                    }
                    RowIndex++;
                }

               
                     
               // sheetData3.Append(Createfooter(RowIndex, revision.Notes));
               sheetData3.Append(CreateTotalRow(StartIndex,RowIndex -1,RowIndex, revision.Notes));
               
                
            }






            foreach (var cell in sheetData3.Descendants<CellFormula>())
                cell.CalculateCell = true;

            ConditionalFormatting conditionalFormatting2 = new ConditionalFormatting() { SequenceOfReferences = new ListValue<StringValue>() { InnerText = "E8:H8" } };

            ConditionalFormattingRule conditionalFormattingRule2 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)2U, Priority = 1, StopIfTrue = true, Operator = ConditionalFormattingOperatorValues.LessThan };
            Formula formula2 = new Formula();
            formula2.Text = "0";

            conditionalFormattingRule2.Append(formula2);

            conditionalFormatting2.Append(conditionalFormattingRule2);

            ConditionalFormatting conditionalFormatting3 = new ConditionalFormatting() { SequenceOfReferences = new ListValue<StringValue>() { InnerText = "E8:H8" } };

            ConditionalFormattingRule conditionalFormattingRule3 = new ConditionalFormattingRule() { Type = ConditionalFormatValues.CellIs, FormatId = (UInt32Value)1U, Priority = 2, StopIfTrue = true, Operator = ConditionalFormattingOperatorValues.LessThan };
            Formula formula3 = new Formula();
            formula3.Text = "0";

            conditionalFormattingRule3.Append(formula3);

            conditionalFormatting3.Append(conditionalFormattingRule3);
            PrintOptions printOptions3 = new PrintOptions() { HorizontalCentered = true };
            PageMargins pageMargins3 = new PageMargins() { Left = 0.553D, Right = 0.333D, Top = 1.25D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D,  };
            PageSetup pageSetup3 = new PageSetup() { Orientation = OrientationValues.Landscape, HorizontalDpi = (UInt32Value)0U, VerticalDpi = (UInt32Value)0U, Id = "rId1" };
            

            MergeCells mergeCells105 = new MergeCells() { Count = (UInt32Value)3U };
            MergeCell mergeCell118a = new MergeCell() { Reference = "A1:L1" };
            MergeCell mergeCell119a = new MergeCell() { Reference = "A2:L2" };
            MergeCell mergeCell120a = new MergeCell() { Reference = "A3:L3" };


            mergeCells105.Append(mergeCell118a);
            mergeCells105.Append(mergeCell119a);
            mergeCells105.Append(mergeCell120a);



            HeaderFooter headerFooter3 = new HeaderFooter();
            OddHeader oddHeader3 = new OddHeader();
            oddHeader3.Text = "&C&G";
            OddFooter oddFooter3 = new OddFooter();
            oddFooter3.Text = "&C&G";

            headerFooter3.Append(oddHeader3);
            headerFooter3.Append(oddFooter3);
            LegacyDrawingHeaderFooter legacyDrawingHeaderFooter3 = new LegacyDrawingHeaderFooter() { Id = "rId2" };

            worksheet3.Append(sheetDimension3);
            worksheet3.Append(sheetViews3);
            worksheet3.Append(sheetFormatProperties3);
            worksheet3.Append(columns3);
            worksheet3.Append(sheetData3);
            worksheet3.Append(mergeCells105);
            worksheet3.Append(printOptions3);
            worksheet3.Append(pageMargins3);
            worksheet3.Append(pageSetup3);
           
            worksheet3.Append(headerFooter3);
            worksheet3.Append(legacyDrawingHeaderFooter3);
          
            worksheetPart3.Worksheet = worksheet3;
        }
        public Row CreateContentBLU(UInt32Value index, RevisionDataModel VisionDataList)
        {
            Row r = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:11" }, Height = 14.25D, DyDescent = 0.3D };
            r.RowIndex = index;

            Cell cell383 = new Cell() { CellReference = "A" + index.ToString(), StyleIndex = (UInt32Value)169U, DataType = CellValues.String };
            CellValue cellValue47 = new CellValue();
            cellValue47.Text = VisionDataList.MeterGroup;

            cell383.Append(cellValue47);

            Cell cell384 = new Cell() { CellReference = "B" + index.ToString(), StyleIndex = (UInt32Value)73U };
            CellValue cellValue48 = new CellValue();
            cellValue48.Text = VisionDataList.ContractVolume.Value.ToString("N0");

            cell384.Append(cellValue48);

            Cell cell385 = new Cell() { CellReference = "C" + index.ToString(), StyleIndex = (UInt32Value)73U };
            CellValue cellValue49 = new CellValue();
            DecimalValue AdjustedVolume = VisionDataList.ActualVolume.Value;
            cellValue49.Text = AdjustedVolume.Value.ToString("N0");

            cell385.Append(cellValue49);

            Cell cell386 = new Cell() { CellReference = "D" + index.ToString(), StyleIndex = (UInt32Value) 165U, DataType = CellValues.String  };
            CellValue cellValue50 = new CellValue();

            cellValue50.Text = VisionDataList.CPP.Value.ToString("$0.#####");

            cell386.Append(cellValue50);

            Cell cell387 = new Cell() { CellReference = "E" + index.ToString(), StyleIndex = (UInt32Value)181U };
            CellFormula cellFormula11 = new CellFormula();
            cellFormula11.Text = "IF(B" + index + "=0,\"\",IF(C" + index + "=\"\",\"\",IF(C" + index + "<B" + index + ",SUM((C" + index + "-B" + index + ")/B" + index + "),SUM((C" + index + "-B" + index + ")/B" + index + "))))";
            CellValue cellValue51 = new CellValue();
            cellValue51.Text = "5";

            cell387.Append(cellFormula11);
            cell387.Append(cellValue51);

            Cell cell388 = new Cell() { CellReference = "F" + index.ToString(), StyleIndex = (UInt32Value)73U };
            CellFormula cellFormula12 = new CellFormula();
            cellFormula12.Text = "IF(C" + index + "=\"\",\"\",IF(B" + index + "=\"\",\"\",IF(SUM(C" + index + "-B" + index + ")<0,\"0\",SUM(C" + index + "-B" + index + "))))";
            CellValue cellValue52 = new CellValue();
            cellValue52.Text = "6";

            cell388.Append(cellFormula12);
            cell388.Append(cellValue52);

            Cell cell386b = new Cell() { CellReference = "G" + index.ToString(), StyleIndex = (UInt32Value)73U, DataType = CellValues.Number };
            CellValue cellValue50a = new CellValue();
            cellValue50a.Text = VisionDataList.Rollover.ToString();

            cell386b.Append(cellValue50a);

            Cell cell389 = new Cell() { CellReference = "H" + index.ToString(), StyleIndex = (UInt32Value)75U };
            CellFormula cellFormula13 = new CellFormula();
            cellFormula13.Text = "IF(SUM(F" + index + " - G" + index + ") <0,\"$0.00\",SUM(F" + index + " - G" + index + ")*D" + index + ")";
            // cellFormula13.Text = VisionDataList.OverageCost.ToString("F2");
            CellValue cellValue53 = new CellValue();
            cellValue53.Text = "7";

            cell389.Append(cellFormula13);
            cell389.Append(cellValue53);

            Cell cell386A = new Cell() { CellReference = "I" + index.ToString(), StyleIndex = (UInt32Value)75U };
            CellValue cellValue50A = new CellValue();
            cellValue50A.Text = VisionDataList.CreditAmount.Value.ToString("F2");

            cell386A.Append(cellValue50A);

            Cell cell390 = new Cell() { CellReference = "J" + index.ToString(), StyleIndex = (UInt32Value)75U };
            CellFormula cellFormula14 = new CellFormula();
            cellFormula14.Text = "SUM(F" + index + "* " + VisionDataList.CPP.Value + ")";
            CellValue cellValue54 = new CellValue();
            cellValue54.Text = "8";

            cell390.Append(cellFormula14);
            cell390.Append(cellValue54);


            Cell cell391 = new Cell() { CellReference = "K" + index.ToString(), StyleIndex = (UInt32Value)61U };
            CellFormula cellFormula15 = new CellFormula();
            cellFormula15.Text = "SUM((J" + index + "-H" + index + ")+ I" + index + ")";
            CellValue cellValue55 = new CellValue();
            cellValue55.Text = "9";

            cell391.Append(cellFormula15);
            cell391.Append(cellValue55);


            r.Append(cell383);
            r.Append(cell384);
            r.Append(cell385);
            r.Append(cell386);
            r.Append(cell387);
            r.Append(cell388);
            r.Append(cell386b);
            r.Append(cell389);
            r.Append(cell386A);
            r.Append(cell390);
            r.Append(cell391);






            return r;
        }
        public Row CreateContentWHT(UInt32Value index, RevisionDataModel VisionDataList)
        {
            Row r = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:11" }, Height = 14.25D, DyDescent = 0.3D };
            r.RowIndex = index;

            Cell cell383 = new Cell() { CellReference = "A" + index.ToString(), StyleIndex = (UInt32Value)58U, DataType = CellValues.String };
            CellValue cellValue47 = new CellValue();
            cellValue47.Text = VisionDataList.MeterGroup;

            cell383.Append(cellValue47);

            Cell cell384 = new Cell() { CellReference = "B" + index.ToString(), StyleIndex = (UInt32Value)6U };
            CellValue cellValue48 = new CellValue();
            cellValue48.Text = VisionDataList.ContractVolume.Value.ToString("N0");

            cell384.Append(cellValue48);

            Cell cell385 = new Cell() { CellReference = "C" + index.ToString(), StyleIndex = (UInt32Value)6U };
            CellValue cellValue49 = new CellValue();
            DecimalValue AdjustedVolume = VisionDataList.ActualVolume.Value;
            cellValue49.Text = AdjustedVolume.Value.ToString("N0");

            cell385.Append(cellValue49);

            Cell cell386 = new Cell() { CellReference = "D" + index.ToString(), StyleIndex = (UInt32Value)6U };
            CellValue cellValue50 = new CellValue();

            cellValue50.Text = VisionDataList.CPP.Value.ToString("$0.#####");

            cell386.Append(cellValue50);

            Cell cell387 = new Cell() { CellReference = "E" + index.ToString(), StyleIndex = (UInt32Value)185U };
            CellFormula cellFormula11 = new CellFormula();
            cellFormula11.Text = "IF(B" + index + "=0,\"\",IF(C" + index + "=\"\",\"\",IF(C" + index + "<B" + index + ",SUM((C" + index + "-B" + index + ")/B" + index + "),SUM((C" + index + "-B" + index + ")/B" + index + "))))";
            CellValue cellValue51 = new CellValue();
            cellValue51.Text = "5";

            cell387.Append(cellFormula11);
            cell387.Append(cellValue51);

            Cell cell388 = new Cell() { CellReference = "F" + index.ToString(), StyleIndex = (UInt32Value)6U };
            CellFormula cellFormula12 = new CellFormula();
            cellFormula12.Text = "IF(C" + index + "=\"\",\"\",IF(B" + index + "=\"\",\"\",IF(SUM(C" + index + "-B" + index + ")<0,\"0\",SUM(C" + index + "-B" + index + "))))";
            CellValue cellValue52 = new CellValue();
            cellValue52.Text = "6";

            cell388.Append(cellFormula12);
            cell388.Append(cellValue52);

            Cell cell386b = new Cell() { CellReference = "G" + index.ToString(), StyleIndex = (UInt32Value)6U };
            CellValue cellValue50a = new CellValue();
            cellValue50a.Text = VisionDataList.Rollover.ToString();

            cell386b.Append(cellValue50a);

            Cell cell389 = new Cell() { CellReference = "H" + index.ToString(), StyleIndex = (UInt32Value)17U };
            CellFormula cellFormula13 = new CellFormula();
            cellFormula13.Text = "IF(SUM(F" + index + " - G" + index + ") <0,\"$0.00\",SUM(F" + index + " - G" + index + ")*D" + index + ")";
            // cellFormula13.Text = VisionDataList.OverageCost.ToString("F2");
            CellValue cellValue53 = new CellValue();
            cellValue53.Text = "7";

            cell389.Append(cellFormula13);
            cell389.Append(cellValue53);

            Cell cell386A = new Cell() { CellReference = "I" + index.ToString(), StyleIndex = (UInt32Value)17U };
            CellValue cellValue50A = new CellValue();
            cellValue50A.Text = VisionDataList.CreditAmount.Value.ToString("F2");

            cell386A.Append(cellValue50A);

            Cell cell390 = new Cell() { CellReference = "J" + index.ToString(), StyleIndex = (UInt32Value)17U };
            CellFormula cellFormula14 = new CellFormula();
            cellFormula14.Text = "SUM(F" + index + "* " + VisionDataList.CPP.Value + ")";
            CellValue cellValue54 = new CellValue();
            cellValue54.Text = "8";

            cell390.Append(cellFormula14);
            cell390.Append(cellValue54);


            Cell cell391 = new Cell() { CellReference = "K" + index.ToString(), StyleIndex = (UInt32Value)59U };
            CellFormula cellFormula15 = new CellFormula();
            cellFormula15.Text = "SUM((J" + index + "-H" + index + ")+ I" + index + ")";
            CellValue cellValue55 = new CellValue();
            cellValue55.Text = "9";

            cell391.Append(cellFormula15);
            cell391.Append(cellValue55);


            r.Append(cell383);
            r.Append(cell384);
            r.Append(cell385);
            r.Append(cell386);
            r.Append(cell387);
            r.Append(cell388);
            r.Append(cell386b);
            r.Append(cell389);
            r.Append(cell386A);
            r.Append(cell390);
            r.Append(cell391);






            return r;
        }
        public Row CreateHeader(UInt32Value index, StringValue Header1)
        {
            Row r = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:11" }, Height = 15D, DyDescent = 0.3D };
            r.RowIndex = index;



            Cell cell361 = new Cell() { CellReference = "A" + index.ToString(), StyleIndex = (UInt32Value)194U, DataType = CellValues.String };
            CellValue cellValue37 = new CellValue();
            cellValue37.Text = Header1;

            cell361.Append(cellValue37);
            Cell cell362 = new Cell() { CellReference = "B" + index.ToString(), StyleIndex = (UInt32Value)194U };
            Cell cell363 = new Cell() { CellReference = "C" + index.ToString(), StyleIndex = (UInt32Value)194U };
            Cell cell364 = new Cell() { CellReference = "D" + index.ToString(), StyleIndex = (UInt32Value)194U };
            Cell cell365 = new Cell() { CellReference = "E" + index.ToString(), StyleIndex = (UInt32Value)194U };
            Cell cell366 = new Cell() { CellReference = "F" + index.ToString(), StyleIndex = (UInt32Value)194U };
            Cell cell366a = new Cell() { CellReference = "G" + index.ToString(), StyleIndex = (UInt32Value)194U };
            Cell cell367 = new Cell() { CellReference = "H" + index.ToString(), StyleIndex = (UInt32Value)194U };
            Cell cell368 = new Cell() { CellReference = "I" + index.ToString(), StyleIndex = (UInt32Value)194U };
            Cell cell369 = new Cell() { CellReference = "J" + index.ToString(), StyleIndex = (UInt32Value)194U };
            Cell cell369A = new Cell() { CellReference = "K" + index.ToString(), StyleIndex = (UInt32Value)194U };
            r.Append(cell361);
            r.Append(cell362);
            r.Append(cell363);
            r.Append(cell364);
            r.Append(cell365);
            r.Append(cell366);
            r.Append(cell366a);
            r.Append(cell367);
            r.Append(cell368);
            r.Append(cell369);
            r.Append(cell369A);


            return r;
        }
        public Row Createfooter(UInt32Value index, StringValue Footer1)
        {
            Row r = new Row() { Spans = new ListValue<StringValue>() { InnerText = "1:11" }, Height = 15D, DyDescent = 0.3D };
            r.RowIndex = index;



            Cell cell361 = new Cell() { CellReference = "A" + index.ToString(), StyleIndex = (UInt32Value)211U, DataType = CellValues.String };
            CellValue cellValue37 = new CellValue();
            cellValue37.Text = Footer1;

            cell361.Append(cellValue37);
            Cell cell362 = new Cell() { CellReference = "B" + index.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell363 = new Cell() { CellReference = "C" + index.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell364 = new Cell() { CellReference = "D" + index.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell365 = new Cell() { CellReference = "E" + index.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell366 = new Cell() { CellReference = "F" + index.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell367 = new Cell() { CellReference = "G" + index.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell368 = new Cell() { CellReference = "H" + index.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell369 = new Cell() { CellReference = "I" + index.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell370 = new Cell() { CellReference = "J" + index.ToString(), StyleIndex = (UInt32Value)174U };
            Cell cell371 = new Cell() { CellReference = "K" + index.ToString(), StyleIndex = (UInt32Value)174U };



            r.Append(cell361);
            r.Append(cell362);
            r.Append(cell363);
            r.Append(cell364);
            r.Append(cell365);
            r.Append(cell366);
            r.Append(cell367);
            r.Append(cell368);
            r.Append(cell369);
            r.Append(cell370);
            r.Append(cell371);

            MergeCells mergeCellsNote = new MergeCells() { Count = (UInt32Value)1U };
            MergeCell mergeCell118a = new MergeCell() { Reference = "A"+ index.ToString()+":I" + index.ToString() };



            mergeCellsNote.Append(mergeCell118a);
            r.Append(mergeCellsNote);


            return r;
        }
        public Row CreateTotalRow(UInt32Value StartRow, UInt32Value EndRow, UInt32Value index, StringValue Totals)
        {
            Row row4a = new Row() { RowIndex = (UInt32Value)index, Spans = new ListValue<StringValue>() { InnerText = "1:10" }, DyDescent = 0.25D };
            Cell cell31a = new Cell() { CellReference = "A" + index.ToString() };
            cell31a.StyleIndex = 174U;
            Cell cell32a = new Cell() { CellReference = "B" + index.ToString() };
            cell32a.StyleIndex = 174U;
            Cell cell33a = new Cell() { CellReference = "C" + index.ToString() };
            cell33a.StyleIndex = 174U;
            Cell cell34a = new Cell() { CellReference = "D" + index.ToString() };
            cell34a.StyleIndex = 174U;
            Cell cell35a = new Cell() { CellReference = "E" + index.ToString() };
            cell35a.StyleIndex = 174U;
            Cell cell36a = new Cell() { CellReference = "F" + index.ToString() };
            cell36a.StyleIndex = 174U;
            Cell cell37a = new Cell() { CellReference = "G" + index.ToString() };
            cell37a.StyleIndex = 174U;
            Cell cell38a = new Cell() { CellReference = "H" + index.ToString() };
            cell38a.StyleIndex = 174U;
            Cell cell39a = new Cell() { CellReference = "I" + index.ToString() };
            cell39a.StyleIndex = 174U;
            Cell cell40a = new Cell() { CellReference = "J" + index.ToString() };
            cell40a.StyleIndex = 174U;
            Cell cell41a = new Cell() { CellReference = "K" + index.ToString() };
            cell41a.StyleIndex = 174U;
            CellValue TotalLabel2a = new CellValue();
            TotalLabel2a.Text = "Revision Savings: ";
            cell40a.Append(TotalLabel2a);
             
            CellFormula RevisionTotal = new CellFormula();
          
            RevisionTotal.Text = "SUM(K" + StartRow.Value.ToString() + ":K" + EndRow.Value.ToString() + ")";
           
            RevisionTotal.CalculateCell = true;
            cell41a.Append(RevisionTotal);

            row4a.Append(cell31a);
            row4a.Append(cell32a);
            row4a.Append(cell33a);
            row4a.Append(cell34a);
            row4a.Append(cell35a);
            row4a.Append(cell36a);
            row4a.Append(cell37a);
            row4a.Append(cell38a);
            row4a.Append(cell39a);
            row4a.Append(cell40a);
            row4a.Append(cell41a);
          

            return row4a;
        }
        public Row CreateHeaderLabels(UInt32Value index) 
        {
            Row r = new Row();
            r.RowIndex = index;


            Cell cell372 = new Cell() { CellReference = "A" + index.ToString(), StyleIndex = (UInt32Value)47U, DataType = CellValues.String };
            CellValue cellValue38 = new CellValue();
            cellValue38.Text = "Meter Type Description";

            cell372.Append(cellValue38);

            Cell cell373 = new Cell() { CellReference = "B" + index.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue39 = new CellValue();
            cellValue39.Text = "Quarterly Contracted Volume";

            cell373.Append(cellValue39);

            Cell cell374 = new Cell() { CellReference = "C" + index.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue40 = new CellValue();
            cellValue40.Text = "Actual Quarterly Volume";

            cell374.Append(cellValue40);

            Cell cell375 = new Cell() { CellReference = "D" + index.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue41 = new CellValue();
            cellValue41.Text = "FPR Overage Rate";
            cell375.Append(cellValue41);

            Cell cell376 = new Cell() { CellReference = "E" + index.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue42 = new CellValue();
            cellValue42.Text = "% Volume Increase / Decrease";

            cell376.Append(cellValue42);

            Cell cell377 = new Cell() { CellReference = "F" + index.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue43 = new CellValue();
            cellValue43.Text = "Overage Volumes";

            cell377.Append(cellValue43);

            Cell cell377a = new Cell() { CellReference = "G" + index.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue43a = new CellValue();
            cellValue43a.Text = "Rollover Volume";

            cell377a.Append(cellValue43a);

            Cell cell378 = new Cell() { CellReference = "H" + index.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue44 = new CellValue();
            cellValue44.Text = "Current Overage Expense";

            cell378.Append(cellValue44);

            Cell cell378a = new Cell() { CellReference = "I" + index.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue44a = new CellValue();
            cellValue44a.Text = "Credits";

            cell378a.Append(cellValue44a);

            Cell cell379 = new Cell() { CellReference = "J" + index.ToString(), StyleIndex = (UInt32Value)45U, DataType = CellValues.String };
            CellValue cellValue45 = new CellValue();
            cellValue45.Text = "Overage Cost W/O FPR";

            cell379.Append(cellValue45);

            Cell cell380 = new Cell() { CellReference = "K" + index.ToString(), StyleIndex = (UInt32Value)46U, DataType = CellValues.String };
            CellValue cellValue46 = new CellValue();
            cellValue46.Text = "Net Overage Savings / Cost Increase";

            cell380.Append(cellValue46);


            r.Append(cell372);
            r.Append(cell373);
            r.Append(cell374);
            r.Append(cell375);
            r.Append(cell376);
            r.Append(cell377);
            r.Append(cell377a);
            r.Append(cell378);
            r.Append(cell378a);
            r.Append(cell379);
            r.Append(cell380);





            return r;
        }
        // Printer Page and Grapics for Headers and footers
        //
        // Generates content of vmlDrawingPart1.
        private void GenerateVmlDrawingPart1Content(VmlDrawingPart vmlDrawingPart1)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(vmlDrawingPart1.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\r\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\r\n <o:shapelayout v:ext=\"edit\">\r\n  <o:idmap v:ext=\"edit\" data=\"7\"/>\r\n </o:shapelayout><v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\"\r\n  o:preferrelative=\"t\" path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\">\r\n  <v:stroke joinstyle=\"miter\"/>\r\n  <v:formulas>\r\n   <v:f eqn=\"if lineDrawn pixelLineWidth 0\"/>\r\n   <v:f eqn=\"sum @0 1 0\"/>\r\n   <v:f eqn=\"sum 0 0 @1\"/>\r\n   <v:f eqn=\"prod @2 1 2\"/>\r\n   <v:f eqn=\"prod @3 21600 pixelWidth\"/>\r\n   <v:f eqn=\"prod @3 21600 pixelHeight\"/>\r\n   <v:f eqn=\"sum @0 0 1\"/>\r\n   <v:f eqn=\"prod @6 1 2\"/>\r\n   <v:f eqn=\"prod @7 21600 pixelWidth\"/>\r\n   <v:f eqn=\"sum @8 21600 0\"/>\r\n   <v:f eqn=\"prod @7 21600 pixelHeight\"/>\r\n   <v:f eqn=\"sum @10 21600 0\"/>\r\n  </v:formulas>\r\n  <v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/>\r\n  <o:lock v:ext=\"edit\" aspectratio=\"t\"/>\r\n </v:shapetype><v:shape id=\"CH\" o:spid=\"_x0000_s7169\" type=\"#_x0000_t75\"\r\n  style=\'position:absolute;margin-left:0;margin-top:0;width:180pt;height:64.5pt;\r\n  z-index:1\'>\r\n  <v:imagedata o:relid=\"rId1\" o:title=\"FPR-VISION\"/>\r\n  <o:lock v:ext=\"edit\" rotation=\"t\"/>\r\n </v:shape><v:shape id=\"CF\" o:spid=\"_x0000_s7172\" type=\"#_x0000_t75\" style=\'position:absolute;\r\n  margin-left:0;margin-top:0;width:678pt;height:36pt;z-index:2\'>\r\n  <v:imagedata o:relid=\"rId2\" o:title=\"Footer\"/>\r\n  <o:lock v:ext=\"edit\" rotation=\"t\"/>\r\n </v:shape></xml>");
            writer.Flush();
            writer.Close();
        }
        // Generates content of imagePart1.
        private void GenerateImagePart1Content(ImagePart imagePart1)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart4Data);
            imagePart1.FeedData(data);
            data.Close();
        }
        // Generates content of imagePart2.
        private void GenerateImagePart2Content(ImagePart imagePart2)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart2Data);
            imagePart2.FeedData(data);
            data.Close();
        }
        // Generates content of vmlDrawingPart2.
        private void GenerateVmlDrawingPart2Content(VmlDrawingPart vmlDrawingPart2)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(vmlDrawingPart2.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\r\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\r\n <o:shapelayout v:ext=\"edit\">\r\n  <o:idmap v:ext=\"edit\" data=\"13\"/>\r\n </o:shapelayout><v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\"\r\n  path=\"m,l,21600r21600,l21600,xe\">\r\n  <v:stroke joinstyle=\"miter\"/>\r\n  <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\r\n </v:shapetype><v:shape id=\"_x0000_s13316\" type=\"#_x0000_t202\" style=\'position:absolute;\r\n  margin-left:366.75pt;margin-top:70.5pt;width:104.25pt;height:57.75pt;\r\n  z-index:1;visibility:hidden\' fillcolor=\"#ffffe1\" o:insetmode=\"auto\">\r\n  <v:fill color2=\"#ffffe1\"/>\r\n  <v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>\r\n  <v:path o:connecttype=\"none\"/>\r\n  <v:textbox style=\'mso-direction-alt:auto\'>\r\n   <div style=\'text-align:left\'></div>\r\n  </v:textbox>\r\n  <x:ClientData ObjectType=\"Note\">\r\n   <x:MoveWithCells/>\r\n   <x:SizeWithCells/>\r\n   <x:Anchor>\r\n    5, 56, 5, 4, 6, 128, 8, 4</x:Anchor>\r\n   <x:AutoFill>False</x:AutoFill>\r\n   <x:Row>5</x:Row>\r\n   <x:Column>5</x:Column>\r\n  </x:ClientData>\r\n </v:shape></xml>");
            writer.Flush();
            writer.Close();
        }
        // Generates content of spreadsheetPrinterSettingsPart1.
        private void GenerateSpreadsheetPrinterSettingsPart1Content(SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart1)
        {
            System.IO.Stream data = GetBinaryDataStream(spreadsheetPrinterSettingsPart1Data);
            spreadsheetPrinterSettingsPart1.FeedData(data);

            data.Close();
        }
        // Generates content of worksheetCommentsPart1.
        private void GenerateWorksheetCommentsPart1Content(WorksheetCommentsPart worksheetCommentsPart1)
        {
            Comments comments1 = new Comments();

            Authors authors1 = new Authors();
            Author author1 = new Author();
            author1.Text = "Brent";

            authors1.Append(author1);

            CommentList commentList1 = new CommentList();

            Comment comment1 = new Comment() { Reference = "F6", AuthorId = (UInt32Value)0U };

            CommentText commentText1 = new CommentText();

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            Bold bold31 = new Bold();
            FontSize fontSize25 = new FontSize() { Val = 9D };
            Color color134 = new Color() { Indexed = (UInt32Value)81U };
            RunFont runFont1 = new RunFont() { Val = "Tahoma" };
            FontFamily fontFamily1 = new FontFamily() { Val = 2 };

            runProperties1.Append(bold31);
            runProperties1.Append(fontSize25);
            runProperties1.Append(color134);
            runProperties1.Append(runFont1);
            runProperties1.Append(fontFamily1);
            Text text1 = new Text();
            // text1.Text = "If this is an On-Going Cost Avoidance item, then DO NOT enter an End Date until it actually ends.";

            run1.Append(runProperties1);
            run1.Append(text1);

            commentText1.Append(run1);

            comment1.Append(commentText1);

            commentList1.Append(comment1);

            comments1.Append(authors1);
            //   comments1.Append(commentList1);

            worksheetCommentsPart1.Comments = comments1;
        }
        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "1F497D" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "EEECE1" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "4F81BD" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "C0504D" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "9BBB59" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "8064A2" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4BACC6" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "F79646" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0000FF" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "800080" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme6 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Cambria" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Arial" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);

            fontScheme6.Append(majorFont1);
            fontScheme6.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint1 = new A.Tint() { Val = 50000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 300000 };

            schemeColor2.Append(tint1);
            schemeColor2.Append(saturationModulation1);

            gradientStop3.Append(schemeColor2);

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 35000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint2 = new A.Tint() { Val = 37000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 300000 };

            schemeColor3.Append(tint2);
            schemeColor3.Append(saturationModulation2);

            gradientStop4.Append(schemeColor3);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint3 = new A.Tint() { Val = 15000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 350000 };

            schemeColor4.Append(tint3);
            schemeColor4.Append(saturationModulation3);

            gradientStop5.Append(schemeColor4);

            gradientStopList1.Append(gradientStop3);
            gradientStopList1.Append(gradientStop4);
            gradientStopList1.Append(gradientStop5);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 16200000, Scaled = true };

            gradientFill2.Append(gradientStopList1);
            gradientFill2.Append(linearGradientFill1);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade1 = new A.Shade() { Val = 51000 };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 130000 };

            schemeColor5.Append(shade1);
            schemeColor5.Append(saturationModulation4);

            gradientStop6.Append(schemeColor5);

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 80000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade2 = new A.Shade() { Val = 93000 };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 130000 };

            schemeColor6.Append(shade2);
            schemeColor6.Append(saturationModulation5);

            gradientStop7.Append(schemeColor6);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade3 = new A.Shade() { Val = 94000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 135000 };

            schemeColor7.Append(shade3);
            schemeColor7.Append(saturationModulation6);

            gradientStop8.Append(schemeColor7);

            gradientStopList2.Append(gradientStop6);
            gradientStopList2.Append(gradientStop7);
            gradientStopList2.Append(gradientStop8);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 16200000, Scaled = false };

            gradientFill3.Append(gradientStopList2);
            gradientFill3.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill2);
            fillStyleList1.Append(gradientFill3);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();

            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade4 = new A.Shade() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 105000 };

            schemeColor8.Append(shade4);
            schemeColor8.Append(saturationModulation7);

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);

            A.Outline outline2 = new A.Outline() { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);

            A.Outline outline3 = new A.Outline() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();

            A.EffectList effectList1 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 38000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList1.Append(outerShadow1);

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();

            A.EffectList effectList2 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha2 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex12.Append(alpha2);

            outerShadow2.Append(rgbColorModelHex12);

            effectList2.Append(outerShadow2);

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow3 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha3 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex13.Append(alpha3);

            outerShadow3.Append(rgbColorModelHex13);

            effectList3.Append(outerShadow3);

            A.Scene3DType scene3DType1 = new A.Scene3DType();

            A.Camera camera1 = new A.Camera() { Preset = A.PresetCameraValues.OrthographicFront };
            A.Rotation rotation1 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 0 };

            camera1.Append(rotation1);

            A.LightRig lightRig1 = new A.LightRig() { Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };
            A.Rotation rotation2 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 1200000 };

            lightRig1.Append(rotation2);

            scene3DType1.Append(camera1);
            scene3DType1.Append(lightRig1);

            A.Shape3DType shape3DType1 = new A.Shape3DType();
            A.BevelTop bevelTop1 = new A.BevelTop() { Width = 63500L, Height = 25400L };

            shape3DType1.Append(bevelTop1);

            effectStyle3.Append(effectList3);
            effectStyle3.Append(scene3DType1);
            effectStyle3.Append(shape3DType1);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.GradientFill gradientFill4 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint4 = new A.Tint() { Val = 40000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 350000 };

            schemeColor12.Append(tint4);
            schemeColor12.Append(saturationModulation8);

            gradientStop9.Append(schemeColor12);

            A.GradientStop gradientStop10 = new A.GradientStop() { Position = 40000 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 45000 };
            A.Shade shade5 = new A.Shade() { Val = 99000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 350000 };

            schemeColor13.Append(tint5);
            schemeColor13.Append(shade5);
            schemeColor13.Append(saturationModulation9);

            gradientStop10.Append(schemeColor13);

            A.GradientStop gradientStop11 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade6 = new A.Shade() { Val = 20000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 255000 };

            schemeColor14.Append(shade6);
            schemeColor14.Append(saturationModulation10);

            gradientStop11.Append(schemeColor14);

            gradientStopList3.Append(gradientStop9);
            gradientStopList3.Append(gradientStop10);
            gradientStopList3.Append(gradientStop11);

            A.PathGradientFill pathGradientFill1 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle1 = new A.FillToRectangle() { Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

            pathGradientFill1.Append(fillToRectangle1);

            gradientFill4.Append(gradientStopList3);
            gradientFill4.Append(pathGradientFill1);

            A.GradientFill gradientFill5 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList4 = new A.GradientStopList();

            A.GradientStop gradientStop12 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 80000 };
            A.SaturationModulation saturationModulation11 = new A.SaturationModulation() { Val = 300000 };

            schemeColor15.Append(tint6);
            schemeColor15.Append(saturationModulation11);

            gradientStop12.Append(schemeColor15);

            A.GradientStop gradientStop13 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade7 = new A.Shade() { Val = 30000 };
            A.SaturationModulation saturationModulation12 = new A.SaturationModulation() { Val = 200000 };

            schemeColor16.Append(shade7);
            schemeColor16.Append(saturationModulation12);

            gradientStop13.Append(schemeColor16);

            gradientStopList4.Append(gradientStop12);
            gradientStopList4.Append(gradientStop13);

            A.PathGradientFill pathGradientFill2 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle2 = new A.FillToRectangle() { Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 };

            pathGradientFill2.Append(fillToRectangle2);

            gradientFill5.Append(gradientStopList4);
            gradientFill5.Append(pathGradientFill2);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(gradientFill4);
            backgroundFillStyleList1.Append(gradientFill5);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme6);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);

            themePart1.Theme = theme1;
        }
        // Generates content of spreadsheetPrinterSettingsPart6.
        private void GenerateSpreadsheetPrinterSettingsPart6Content(SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart6)
        {
            System.IO.Stream data = GetBinaryDataStream(spreadsheetPrinterSettingsPart6Data);
            spreadsheetPrinterSettingsPart6.FeedData(data);
            data.Close();
        }
        // Generates content of sharedStringTablePart1.
        private void GenerateSharedStringTablePart1Content(SharedStringTablePart sharedStringTablePart1)
        {
            SharedStringTable sharedStringTable1 = new SharedStringTable() { Count = (UInt32Value)58U, UniqueCount = (UInt32Value)47U };

            SharedStringItem sharedStringItem1 = new SharedStringItem();
            Text text2 = new Text();
            text2.Text = "Meter Group";

            sharedStringItem1.Append(text2);

            SharedStringItem sharedStringItem2 = new SharedStringItem();
            Text text3 = new Text();
            text3.Text = "Start Date";

            sharedStringItem2.Append(text3);

            SharedStringItem sharedStringItem3 = new SharedStringItem();
            Text text4 = new Text();
            text4.Text = "End Date";

            sharedStringItem3.Append(text4);

            SharedStringItem sharedStringItem4 = new SharedStringItem();
            Text text5 = new Text();
            text5.Text = "Date";

            sharedStringItem4.Append(text5);

            SharedStringItem sharedStringItem5 = new SharedStringItem();
            Text text6 = new Text();
            text6.Text = "Asset Cost If Purchased";

            sharedStringItem5.Append(text6);

            SharedStringItem sharedStringItem6 = new SharedStringItem();
            Text text7 = new Text();
            text7.Text = "Comments";

            sharedStringItem6.Append(text7);

            SharedStringItem sharedStringItem7 = new SharedStringItem();
            Text text8 = new Text();
            text8.Text = "Client Name:";

            sharedStringItem7.Append(text8);

            SharedStringItem sharedStringItem8 = new SharedStringItem();
            Text text9 = new Text();
            text9.Text = "Period Ending:";

            sharedStringItem8.Append(text9);

            SharedStringItem sharedStringItem9 = new SharedStringItem();
            Text text10 = new Text();
            text10.Text = "Enter Data Below:";

            sharedStringItem9.Append(text10);

            SharedStringItem sharedStringItem10 = new SharedStringItem();
            Text text11 = new Text();
            text11.Text = "Contract Start Date:";

            sharedStringItem10.Append(text11);

            SharedStringItem sharedStringItem11 = new SharedStringItem();
            Text text12 = new Text();
            text12.Text = "Savings Term";

            sharedStringItem11.Append(text12);

            SharedStringItem sharedStringItem12 = new SharedStringItem();
            Text text13 = new Text();
            text13.Text = "Savings Type";

            sharedStringItem12.Append(text13);

            SharedStringItem sharedStringItem13 = new SharedStringItem();
            Text text14 = new Text();
            text14.Text = "Total Savings";

            sharedStringItem13.Append(text14);

            SharedStringItem sharedStringItem14 = new SharedStringItem();
            Text text15 = new Text();
            text15.Text = "Service Call Experience";

            sharedStringItem14.Append(text15);

            SharedStringItem sharedStringItem15 = new SharedStringItem();
            Text text16 = new Text();
            text16.Text = "Account Management Needs";

            sharedStringItem15.Append(text16);

            SharedStringItem sharedStringItem16 = new SharedStringItem();
            Text text17 = new Text();
            text17.Text = "Professionalism";

            sharedStringItem16.Append(text17);

            SharedStringItem sharedStringItem17 = new SharedStringItem();
            Text text18 = new Text();
            text18.Text = "Par Level Management";

            sharedStringItem17.Append(text18);

            SharedStringItem sharedStringItem18 = new SharedStringItem();
            Text text19 = new Text();
            text19.Text = "Servicing Companies Performance";

            sharedStringItem18.Append(text19);

            SharedStringItem sharedStringItem19 = new SharedStringItem();
            Text text20 = new Text();
            text20.Text = "Submitted by";

            sharedStringItem19.Append(text20);

            SharedStringItem sharedStringItem20 = new SharedStringItem();
            Text text21 = new Text();
            text21.Text = "E-Mail Address";

            sharedStringItem20.Append(text21);

            SharedStringItem sharedStringItem21 = new SharedStringItem();
            Text text22 = new Text();
            text22.Text = "Fax Number";

            sharedStringItem21.Append(text22);

            SharedStringItem sharedStringItem22 = new SharedStringItem();
            Text text23 = new Text();
            text23.Text = "Pages By User";

            sharedStringItem22.Append(text23);

            SharedStringItem sharedStringItem23 = new SharedStringItem();
            Text text24 = new Text();
            text24.Text = "Total Users:";

            sharedStringItem23.Append(text24);

            SharedStringItem sharedStringItem24 = new SharedStringItem();
            Text text25 = new Text();
            text25.Text = "Total Pages:";

            sharedStringItem24.Append(text25);

            SharedStringItem sharedStringItem25 = new SharedStringItem();
            Text text26 = new Text();
            text26.Text = "Amount";

            sharedStringItem25.Append(text26);

            SharedStringItem sharedStringItem26 = new SharedStringItem();
            Text text27 = new Text();
            text27.Text = "?";

            sharedStringItem26.Append(text27);

            SharedStringItem sharedStringItem27 = new SharedStringItem();
            Text text28 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text28.Text = "Below, you will find a summary of your Internet Faxing service usage for the current billing period.  ";

            sharedStringItem27.Append(text28);

            SharedStringItem sharedStringItem28 = new SharedStringItem();
            Text text29 = new Text();
            text29.Text = "Rollover Savings Calculator";

            sharedStringItem28.Append(text29);

            SharedStringItem sharedStringItem29 = new SharedStringItem();
            Text text30 = new Text();
            text30.Text = "Pre-FPR CPP";

            sharedStringItem29.Append(text30);

            SharedStringItem sharedStringItem30 = new SharedStringItem();
            Text text31 = new Text();
            text31.Text = "Date Submitted";

            sharedStringItem30.Append(text31);

            SharedStringItem sharedStringItem31 = new SharedStringItem();
            Text text32 = new Text();
            text32.Text = "Survey Completed by:";

            sharedStringItem31.Append(text32);

            SharedStringItem sharedStringItem32 = new SharedStringItem();
            Text text33 = new Text();
            text33.Text = "E-Mail:";

            sharedStringItem32.Append(text33);

            SharedStringItem sharedStringItem33 = new SharedStringItem();
            Text text34 = new Text();
            text34.Text = "Date:";

            sharedStringItem33.Append(text34);

            SharedStringItem sharedStringItem34 = new SharedStringItem();
            Text text35 = new Text();
            text35.Text = "Criteria";

            sharedStringItem34.Append(text35);

            SharedStringItem sharedStringItem35 = new SharedStringItem();
            Text text36 = new Text();
            text36.Text = "FPR Professionalism";

            sharedStringItem35.Append(text36);

            SharedStringItem sharedStringItem36 = new SharedStringItem();
            Text text37 = new Text();
            text37.Text = "Par-Level Management";

            sharedStringItem36.Append(text37);

            SharedStringItem sharedStringItem37 = new SharedStringItem();
            Text text38 = new Text();
            text38.Text = "Comments:";

            sharedStringItem37.Append(text38);

            SharedStringItem sharedStringItem38 = new SharedStringItem();
            Text text39 = new Text();
            text39.Text = "Below, you will find a summary of the different areas where FPR was able to deliver solutions resulting in cost-avoidance.  Some examples of cost-avoidance solutions include color reduction/elimination strategies, vendor re-negotiation opportunities, and many others.";

            sharedStringItem38.Append(text39);

            SharedStringItem sharedStringItem39 = new SharedStringItem();
            Text text40 = new Text();
            text40.Text = "Total Cost Avoidance:";

            sharedStringItem39.Append(text40);

            SharedStringItem sharedStringItem40 = new SharedStringItem();
            Text text41 = new Text();
            text41.Text = "New Model";

            sharedStringItem40.Append(text41);

            SharedStringItem sharedStringItem41 = new SharedStringItem();
            Text text42 = new Text();
            text42.Text = "Removed Model";

            sharedStringItem41.Append(text42);

            SharedStringItem sharedStringItem42 = new SharedStringItem();
            Text text43 = new Text();
            text43.Text = "New Serial Number";

            sharedStringItem42.Append(text43);

            SharedStringItem sharedStringItem43 = new SharedStringItem();
            Text text44 = new Text();
            text44.Text = "There are currently no cost-avoidance details to report.";

            sharedStringItem43.Append(text44);

            SharedStringItem sharedStringItem44 = new SharedStringItem();
            Text text45 = new Text();
            text45.Text = "There is currently no Easylink usage.";

            sharedStringItem44.Append(text45);

            SharedStringItem sharedStringItem45 = new SharedStringItem();
            Text text46 = new Text();
            text46.Text = "*This document is to be completed by FPR personnel in front of client.";

            sharedStringItem45.Append(text46);

            SharedStringItem sharedStringItem46 = new SharedStringItem();
            Text text47 = new Text();
            text47.Text = "Additional Comments:";

            sharedStringItem46.Append(text47);

            SharedStringItem sharedStringItem47 = new SharedStringItem();
            Text text48 = new Text();
            text48.Text = "Each period, FPR petitions your feedback regarding our overall delivery of service.  The historical results provided to us are detailed below.  On a scale from 1 (worst) to 10 (best), the results below indicate how you have rated us in the past, for the respective categories.";

            sharedStringItem47.Append(text48);

            sharedStringTable1.Append(sharedStringItem1);
            sharedStringTable1.Append(sharedStringItem2);
            sharedStringTable1.Append(sharedStringItem3);
            sharedStringTable1.Append(sharedStringItem4);
            sharedStringTable1.Append(sharedStringItem5);
            sharedStringTable1.Append(sharedStringItem6);
            sharedStringTable1.Append(sharedStringItem7);
            sharedStringTable1.Append(sharedStringItem8);
            sharedStringTable1.Append(sharedStringItem9);
            sharedStringTable1.Append(sharedStringItem10);
            sharedStringTable1.Append(sharedStringItem11);
            sharedStringTable1.Append(sharedStringItem12);
            sharedStringTable1.Append(sharedStringItem13);
            sharedStringTable1.Append(sharedStringItem14);
            sharedStringTable1.Append(sharedStringItem15);
            sharedStringTable1.Append(sharedStringItem16);
            sharedStringTable1.Append(sharedStringItem17);
            sharedStringTable1.Append(sharedStringItem18);
            sharedStringTable1.Append(sharedStringItem19);
            sharedStringTable1.Append(sharedStringItem20);
            sharedStringTable1.Append(sharedStringItem21);
            sharedStringTable1.Append(sharedStringItem22);
            sharedStringTable1.Append(sharedStringItem23);
            sharedStringTable1.Append(sharedStringItem24);
            sharedStringTable1.Append(sharedStringItem25);
            sharedStringTable1.Append(sharedStringItem26);
            sharedStringTable1.Append(sharedStringItem27);
            sharedStringTable1.Append(sharedStringItem28);
            sharedStringTable1.Append(sharedStringItem29);
            sharedStringTable1.Append(sharedStringItem30);
            sharedStringTable1.Append(sharedStringItem31);
            sharedStringTable1.Append(sharedStringItem32);
            sharedStringTable1.Append(sharedStringItem33);
            sharedStringTable1.Append(sharedStringItem34);
            sharedStringTable1.Append(sharedStringItem35);
            sharedStringTable1.Append(sharedStringItem36);
            sharedStringTable1.Append(sharedStringItem37);
            sharedStringTable1.Append(sharedStringItem38);
            sharedStringTable1.Append(sharedStringItem39);
            sharedStringTable1.Append(sharedStringItem40);
            sharedStringTable1.Append(sharedStringItem41);
            sharedStringTable1.Append(sharedStringItem42);
            sharedStringTable1.Append(sharedStringItem43);
            sharedStringTable1.Append(sharedStringItem44);
            sharedStringTable1.Append(sharedStringItem45);
            sharedStringTable1.Append(sharedStringItem46);
            sharedStringTable1.Append(sharedStringItem47);

            sharedStringTablePart1.SharedStringTable = sharedStringTable1;
        }
        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Brent";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2012-04-24T20:14:21Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2012-09-13T18:17:24Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "erhardte";
            document.PackageProperties.LastPrinted = System.Xml.XmlConvert.ToDateTime("2012-08-22T21:01:47Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }
        
        // Generates content of vmlDrawingPart4.
        private void GenerateVmlDrawingPart4Content(VmlDrawingPart vmlDrawingPart4)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(vmlDrawingPart4.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\r\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\r\n <o:shapelayout v:ext=\"edit\">\r\n  <o:idmap v:ext=\"edit\" data=\"13\"/>\r\n </o:shapelayout><v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\"\r\n  o:preferrelative=\"t\" path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\">\r\n  <v:stroke joinstyle=\"miter\"/>\r\n  <v:formulas>\r\n   <v:f eqn=\"if lineDrawn pixelLineWidth 0\"/>\r\n   <v:f eqn=\"sum @0 1 0\"/>\r\n   <v:f eqn=\"sum 0 0 @1\"/>\r\n   <v:f eqn=\"prod @2 1 2\"/>\r\n   <v:f eqn=\"prod @3 21600 pixelWidth\"/>\r\n   <v:f eqn=\"prod @3 21600 pixelHeight\"/>\r\n   <v:f eqn=\"sum @0 0 1\"/>\r\n   <v:f eqn=\"prod @6 1 2\"/>\r\n   <v:f eqn=\"prod @7 21600 pixelWidth\"/>\r\n   <v:f eqn=\"sum @8 21600 0\"/>\r\n   <v:f eqn=\"prod @7 21600 pixelHeight\"/>\r\n   <v:f eqn=\"sum @10 21600 0\"/>\r\n  </v:formulas>\r\n  <v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/>\r\n  <o:lock v:ext=\"edit\" aspectratio=\"t\"/>\r\n </v:shapetype><v:shape id=\"CH\" o:spid=\"_x0000_s13313\" type=\"#_x0000_t75\"\r\n  style=\'position:absolute;margin-left:0;margin-top:0;width:180pt;height:64.5pt;\r\n  z-index:1\'>\r\n  <v:imagedata o:relid=\"rId1\" o:title=\"FPR-VISION\"/>\r\n  <o:lock v:ext=\"edit\" rotation=\"t\"/>\r\n </v:shape><v:shape id=\"CF\" o:spid=\"_x0000_s13314\" type=\"#_x0000_t75\" style=\'position:absolute;\r\n  margin-left:0;margin-top:0;width:10in;height:36pt;z-index:2\'>\r\n  <v:imagedata o:relid=\"rId2\" o:title=\"FPR Letterhead Footer-Corp LeftBlock\"/>\r\n  <o:lock v:ext=\"edit\" rotation=\"t\"/>\r\n </v:shape></xml>");
            writer.Flush();
            writer.Close();
        }
        // Generates content of imagePart3.
        private void GenerateImagePart3Content(ImagePart imagePart3)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart4Data);
            imagePart3.FeedData(data);
            data.Close();
        }
        // Generates content of spreadsheetPrinterSettingsPart3.
        private void GenerateSpreadsheetPrinterSettingsPart3Content(SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart3)
        {
            System.IO.Stream data = GetBinaryDataStream(spreadsheetPrinterSettingsPart3Data);
            spreadsheetPrinterSettingsPart3.FeedData(data);
            data.Close();
        }
        // Generates content of spreadsheetPrinterSettingsPart5.
        private void GenerateSpreadsheetPrinterSettingsPart5Content(SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart5)
        {
            System.IO.Stream data = GetBinaryDataStream(spreadsheetPrinterSettingsPart5Data);
            spreadsheetPrinterSettingsPart5.FeedData(data);
            data.Close();
        }
        // Generates content of calculationChainPart1.
        private void GenerateCalculationChainPart1Content(CalculationChainPart calculationChainPart1)
        {
            CalculationChain calculationChain1 = new CalculationChain();
            CalculationCell calculationCell1 = new CalculationCell() { CellReference = "B4", SheetId = 5, NewLevel = true };
            CalculationCell calculationCell2 = new CalculationCell() { CellReference = "I27", SheetId = 7, NewLevel = true };
            CalculationCell calculationCell3 = new CalculationCell() { CellReference = "I26", SheetId = 7 };
            CalculationCell calculationCell4 = new CalculationCell() { CellReference = "I25", SheetId = 7 };
            CalculationCell calculationCell5 = new CalculationCell() { CellReference = "I24", SheetId = 7 };
            CalculationCell calculationCell6 = new CalculationCell() { CellReference = "I23", SheetId = 7 };
            CalculationCell calculationCell7 = new CalculationCell() { CellReference = "I22", SheetId = 7 };
            CalculationCell calculationCell8 = new CalculationCell() { CellReference = "I21", SheetId = 7 };
            CalculationCell calculationCell9 = new CalculationCell() { CellReference = "I20", SheetId = 7 };
            CalculationCell calculationCell10 = new CalculationCell() { CellReference = "I19", SheetId = 7 };
            CalculationCell calculationCell11 = new CalculationCell() { CellReference = "I18", SheetId = 7 };
            CalculationCell calculationCell12 = new CalculationCell() { CellReference = "I17", SheetId = 7 };
            CalculationCell calculationCell13 = new CalculationCell() { CellReference = "I16", SheetId = 7 };
            CalculationCell calculationCell14 = new CalculationCell() { CellReference = "I15", SheetId = 7 };
            CalculationCell calculationCell15 = new CalculationCell() { CellReference = "I14", SheetId = 7 };
            CalculationCell calculationCell16 = new CalculationCell() { CellReference = "I13", SheetId = 7 };
            CalculationCell calculationCell17 = new CalculationCell() { CellReference = "I12", SheetId = 7 };
            CalculationCell calculationCell18 = new CalculationCell() { CellReference = "I11", SheetId = 7 };
            CalculationCell calculationCell19 = new CalculationCell() { CellReference = "I10", SheetId = 7 };
            CalculationCell calculationCell20 = new CalculationCell() { CellReference = "I9", SheetId = 7 };
            CalculationCell calculationCell21 = new CalculationCell() { CellReference = "I8", SheetId = 7 };
            CalculationCell calculationCell22 = new CalculationCell() { CellReference = "I7", SheetId = 7 };
            CalculationCell calculationCell23 = new CalculationCell() { CellReference = "A29", SheetId = 9, NewLevel = true };
            CalculationCell calculationCell24 = new CalculationCell() { CellReference = "A28", SheetId = 9 };
            CalculationCell calculationCell25 = new CalculationCell() { CellReference = "A27", SheetId = 9 };
            CalculationCell calculationCell26 = new CalculationCell() { CellReference = "A26", SheetId = 9 };
            CalculationCell calculationCell27 = new CalculationCell() { CellReference = "A25", SheetId = 9 };
            CalculationCell calculationCell28 = new CalculationCell() { CellReference = "A24", SheetId = 9 };
            CalculationCell calculationCell29 = new CalculationCell() { CellReference = "A23", SheetId = 9 };
            CalculationCell calculationCell30 = new CalculationCell() { CellReference = "G30", SheetId = 5 };
            CalculationCell calculationCell31 = new CalculationCell() { CellReference = "G29", SheetId = 5 };
            CalculationCell calculationCell32 = new CalculationCell() { CellReference = "G28", SheetId = 5 };
            CalculationCell calculationCell33 = new CalculationCell() { CellReference = "G27", SheetId = 5 };
            CalculationCell calculationCell34 = new CalculationCell() { CellReference = "G26", SheetId = 5 };
            CalculationCell calculationCell35 = new CalculationCell() { CellReference = "A22", SheetId = 9, NewLevel = true };
            CalculationCell calculationCell36 = new CalculationCell() { CellReference = "A21", SheetId = 9 };
            CalculationCell calculationCell37 = new CalculationCell() { CellReference = "A20", SheetId = 9 };
            CalculationCell calculationCell38 = new CalculationCell() { CellReference = "A19", SheetId = 9 };
            CalculationCell calculationCell39 = new CalculationCell() { CellReference = "A18", SheetId = 9 };
            CalculationCell calculationCell40 = new CalculationCell() { CellReference = "A17", SheetId = 9 };
            CalculationCell calculationCell41 = new CalculationCell() { CellReference = "A16", SheetId = 9 };
            CalculationCell calculationCell42 = new CalculationCell() { CellReference = "A15", SheetId = 9 };
            CalculationCell calculationCell43 = new CalculationCell() { CellReference = "A14", SheetId = 9 };
            CalculationCell calculationCell44 = new CalculationCell() { CellReference = "A13", SheetId = 9, NewLevel = true };
            CalculationCell calculationCell45 = new CalculationCell() { CellReference = "C6", SheetId = 9, NewLevel = true };
            CalculationCell calculationCell46 = new CalculationCell() { CellReference = "G25", SheetId = 5, NewLevel = true };
            CalculationCell calculationCell47 = new CalculationCell() { CellReference = "G24", SheetId = 5 };
            CalculationCell calculationCell48 = new CalculationCell() { CellReference = "G23", SheetId = 5 };
            CalculationCell calculationCell49 = new CalculationCell() { CellReference = "G22", SheetId = 5 };
            CalculationCell calculationCell50 = new CalculationCell() { CellReference = "G21", SheetId = 5 };
            CalculationCell calculationCell51 = new CalculationCell() { CellReference = "G20", SheetId = 5 };
            CalculationCell calculationCell52 = new CalculationCell() { CellReference = "G19", SheetId = 5 };
            CalculationCell calculationCell53 = new CalculationCell() { CellReference = "G18", SheetId = 5 };
            CalculationCell calculationCell54 = new CalculationCell() { CellReference = "G17", SheetId = 5 };
            CalculationCell calculationCell55 = new CalculationCell() { CellReference = "G16", SheetId = 5 };
            CalculationCell calculationCell56 = new CalculationCell() { CellReference = "G15", SheetId = 5 };
            CalculationCell calculationCell57 = new CalculationCell() { CellReference = "G14", SheetId = 5 };
            CalculationCell calculationCell58 = new CalculationCell() { CellReference = "G13", SheetId = 5 };
            CalculationCell calculationCell59 = new CalculationCell() { CellReference = "G12", SheetId = 5 };
            CalculationCell calculationCell60 = new CalculationCell() { CellReference = "G11", SheetId = 5 };
            CalculationCell calculationCell61 = new CalculationCell() { CellReference = "G10", SheetId = 5 };
            CalculationCell calculationCell62 = new CalculationCell() { CellReference = "G9", SheetId = 5 };
            CalculationCell calculationCell63 = new CalculationCell() { CellReference = "G8", SheetId = 5 };
            CalculationCell calculationCell64 = new CalculationCell() { CellReference = "C9", SheetId = 12, NewLevel = true };
            CalculationCell calculationCell65 = new CalculationCell() { CellReference = "I28", SheetId = 7, NewLevel = true };
            CalculationCell calculationCell66 = new CalculationCell() { CellReference = "A2", SheetId = 10, NewLevel = true };
            CalculationCell calculationCell67 = new CalculationCell() { CellReference = "A2", SheetId = 5 };
            CalculationCell calculationCell68 = new CalculationCell() { CellReference = "A2", SheetId = 7, NewLevel = true };
            CalculationCell calculationCell69 = new CalculationCell() { CellReference = "C7", SheetId = 9, NewLevel = true };
            CalculationCell calculationCell70 = new CalculationCell() { CellReference = "A12", SheetId = 9 };
            CalculationCell calculationCell71 = new CalculationCell() { CellReference = "A11", SheetId = 9 };
            CalculationCell calculationCell72 = new CalculationCell() { CellReference = "A10", SheetId = 9 };
            CalculationCell calculationCell73 = new CalculationCell() { CellReference = "A9", SheetId = 9 };
            CalculationCell calculationCell74 = new CalculationCell() { CellReference = "A1", SheetId = 14, NewLevel = true };
            CalculationCell calculationCell75 = new CalculationCell() { CellReference = "A2", SheetId = 14, NewLevel = true };
            CalculationCell calculationCell76 = new CalculationCell() { CellReference = "G6", SheetId = 5, NewLevel = true };
            CalculationCell calculationCell77 = new CalculationCell() { CellReference = "F27", SheetId = 7, NewLevel = true };
            CalculationCell calculationCell78 = new CalculationCell() { CellReference = "F26", SheetId = 7 };
            CalculationCell calculationCell79 = new CalculationCell() { CellReference = "F25", SheetId = 7 };
            CalculationCell calculationCell80 = new CalculationCell() { CellReference = "F24", SheetId = 7 };
            CalculationCell calculationCell81 = new CalculationCell() { CellReference = "F23", SheetId = 7 };
            CalculationCell calculationCell82 = new CalculationCell() { CellReference = "F22", SheetId = 7 };
            CalculationCell calculationCell83 = new CalculationCell() { CellReference = "F21", SheetId = 7 };
            CalculationCell calculationCell84 = new CalculationCell() { CellReference = "F20", SheetId = 7 };
            CalculationCell calculationCell85 = new CalculationCell() { CellReference = "F19", SheetId = 7 };
            CalculationCell calculationCell86 = new CalculationCell() { CellReference = "F18", SheetId = 7 };
            CalculationCell calculationCell87 = new CalculationCell() { CellReference = "F17", SheetId = 7 };
            CalculationCell calculationCell88 = new CalculationCell() { CellReference = "F16", SheetId = 7 };
            CalculationCell calculationCell89 = new CalculationCell() { CellReference = "F15", SheetId = 7 };
            CalculationCell calculationCell90 = new CalculationCell() { CellReference = "F14", SheetId = 7 };
            CalculationCell calculationCell91 = new CalculationCell() { CellReference = "F13", SheetId = 7 };
            CalculationCell calculationCell92 = new CalculationCell() { CellReference = "F12", SheetId = 7 };
            CalculationCell calculationCell93 = new CalculationCell() { CellReference = "F11", SheetId = 7 };
            CalculationCell calculationCell94 = new CalculationCell() { CellReference = "F10", SheetId = 7 };
            CalculationCell calculationCell95 = new CalculationCell() { CellReference = "F9", SheetId = 7 };
            CalculationCell calculationCell96 = new CalculationCell() { CellReference = "F8", SheetId = 7 };
            CalculationCell calculationCell97 = new CalculationCell() { CellReference = "A2", SheetId = 9 };
            CalculationCell calculationCell98 = new CalculationCell() { CellReference = "A1", SheetId = 9, NewLevel = true };
            CalculationCell calculationCell99 = new CalculationCell() { CellReference = "A1", SheetId = 10 };
            CalculationCell calculationCell100 = new CalculationCell() { CellReference = "I5", SheetId = 7, NewLevel = true };
            CalculationCell calculationCell101 = new CalculationCell() { CellReference = "A1", SheetId = 7 };
            CalculationCell calculationCell102 = new CalculationCell() { CellReference = "A1", SheetId = 5 };

            calculationChain1.Append(calculationCell1);
            calculationChain1.Append(calculationCell2);
            calculationChain1.Append(calculationCell3);
            calculationChain1.Append(calculationCell4);
            calculationChain1.Append(calculationCell5);
            calculationChain1.Append(calculationCell6);
            calculationChain1.Append(calculationCell7);
            calculationChain1.Append(calculationCell8);
            calculationChain1.Append(calculationCell9);
            calculationChain1.Append(calculationCell10);
            calculationChain1.Append(calculationCell11);
            calculationChain1.Append(calculationCell12);
            calculationChain1.Append(calculationCell13);
            calculationChain1.Append(calculationCell14);
            calculationChain1.Append(calculationCell15);
            calculationChain1.Append(calculationCell16);
            calculationChain1.Append(calculationCell17);
            calculationChain1.Append(calculationCell18);
            calculationChain1.Append(calculationCell19);
            calculationChain1.Append(calculationCell20);
            calculationChain1.Append(calculationCell21);
            calculationChain1.Append(calculationCell22);
            calculationChain1.Append(calculationCell23);
            calculationChain1.Append(calculationCell24);
            calculationChain1.Append(calculationCell25);
            calculationChain1.Append(calculationCell26);
            calculationChain1.Append(calculationCell27);
            calculationChain1.Append(calculationCell28);
            calculationChain1.Append(calculationCell29);
            calculationChain1.Append(calculationCell30);
            calculationChain1.Append(calculationCell31);
            calculationChain1.Append(calculationCell32);
            calculationChain1.Append(calculationCell33);
            calculationChain1.Append(calculationCell34);
            calculationChain1.Append(calculationCell35);
            calculationChain1.Append(calculationCell36);
            calculationChain1.Append(calculationCell37);
            calculationChain1.Append(calculationCell38);
            calculationChain1.Append(calculationCell39);
            calculationChain1.Append(calculationCell40);
            calculationChain1.Append(calculationCell41);
            calculationChain1.Append(calculationCell42);
            calculationChain1.Append(calculationCell43);
            calculationChain1.Append(calculationCell44);
            calculationChain1.Append(calculationCell45);
            calculationChain1.Append(calculationCell46);
            calculationChain1.Append(calculationCell47);
            calculationChain1.Append(calculationCell48);
            calculationChain1.Append(calculationCell49);
            calculationChain1.Append(calculationCell50);
            calculationChain1.Append(calculationCell51);
            calculationChain1.Append(calculationCell52);
            calculationChain1.Append(calculationCell53);
            calculationChain1.Append(calculationCell54);
            calculationChain1.Append(calculationCell55);
            calculationChain1.Append(calculationCell56);
            calculationChain1.Append(calculationCell57);
            calculationChain1.Append(calculationCell58);
            calculationChain1.Append(calculationCell59);
            calculationChain1.Append(calculationCell60);
            calculationChain1.Append(calculationCell61);
            calculationChain1.Append(calculationCell62);
            calculationChain1.Append(calculationCell63);
            calculationChain1.Append(calculationCell64);
            calculationChain1.Append(calculationCell65);
            calculationChain1.Append(calculationCell66);
            calculationChain1.Append(calculationCell67);
            calculationChain1.Append(calculationCell68);
            calculationChain1.Append(calculationCell69);
            calculationChain1.Append(calculationCell70);
            calculationChain1.Append(calculationCell71);
            calculationChain1.Append(calculationCell72);
            calculationChain1.Append(calculationCell73);
            calculationChain1.Append(calculationCell74);
            calculationChain1.Append(calculationCell75);
            calculationChain1.Append(calculationCell76);
            calculationChain1.Append(calculationCell77);
            calculationChain1.Append(calculationCell78);
            calculationChain1.Append(calculationCell79);
            calculationChain1.Append(calculationCell80);
            calculationChain1.Append(calculationCell81);
            calculationChain1.Append(calculationCell82);
            calculationChain1.Append(calculationCell83);
            calculationChain1.Append(calculationCell84);
            calculationChain1.Append(calculationCell85);
            calculationChain1.Append(calculationCell86);
            calculationChain1.Append(calculationCell87);
            calculationChain1.Append(calculationCell88);
            calculationChain1.Append(calculationCell89);
            calculationChain1.Append(calculationCell90);
            calculationChain1.Append(calculationCell91);
            calculationChain1.Append(calculationCell92);
            calculationChain1.Append(calculationCell93);
            calculationChain1.Append(calculationCell94);
            calculationChain1.Append(calculationCell95);
            calculationChain1.Append(calculationCell96);
            calculationChain1.Append(calculationCell97);
            calculationChain1.Append(calculationCell98);
            calculationChain1.Append(calculationCell99);
            calculationChain1.Append(calculationCell100);
            calculationChain1.Append(calculationCell101);
            calculationChain1.Append(calculationCell102);

            calculationChainPart1.CalculationChain = calculationChain1;
        }
        // Generates content of vmlDrawingPart3.
        private void GenerateVmlDrawingPart3Content(VmlDrawingPart vmlDrawingPart3)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(vmlDrawingPart3.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\r\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\r\n <o:shapelayout v:ext=\"edit\">\r\n  <o:idmap v:ext=\"edit\" data=\"5\"/>\r\n </o:shapelayout><v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\"\r\n  o:preferrelative=\"t\" path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\">\r\n  <v:stroke joinstyle=\"miter\"/>\r\n  <v:formulas>\r\n   <v:f eqn=\"if lineDrawn pixelLineWidth 0\"/>\r\n   <v:f eqn=\"sum @0 1 0\"/>\r\n   <v:f eqn=\"sum 0 0 @1\"/>\r\n   <v:f eqn=\"prod @2 1 2\"/>\r\n   <v:f eqn=\"prod @3 21600 pixelWidth\"/>\r\n   <v:f eqn=\"prod @3 21600 pixelHeight\"/>\r\n   <v:f eqn=\"sum @0 0 1\"/>\r\n   <v:f eqn=\"prod @6 1 2\"/>\r\n   <v:f eqn=\"prod @7 21600 pixelWidth\"/>\r\n   <v:f eqn=\"sum @8 21600 0\"/>\r\n   <v:f eqn=\"prod @7 21600 pixelHeight\"/>\r\n   <v:f eqn=\"sum @10 21600 0\"/>\r\n  </v:formulas>\r\n  <v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/>\r\n  <o:lock v:ext=\"edit\" aspectratio=\"t\"/>\r\n </v:shapetype><v:shape id=\"CH\" o:spid=\"_x0000_s5121\" type=\"#_x0000_t75\"\r\n  style=\'position:absolute;margin-left:0;margin-top:0;width:180pt;height:64.5pt;\r\n  z-index:1\'>\r\n  <v:imagedata o:relid=\"rId1\" o:title=\"FPR-VISION\"/>\r\n  <o:lock v:ext=\"edit\" rotation=\"t\"/>\r\n </v:shape><v:shape id=\"CF\" o:spid=\"_x0000_s5124\" type=\"#_x0000_t75\" style=\'position:absolute;\r\n  margin-left:0;margin-top:0;width:678pt;height:36pt;z-index:2\'>\r\n  <v:imagedata o:relid=\"rId2\" o:title=\"Footer\"/>\r\n  <o:lock v:ext=\"edit\" rotation=\"t\"/>\r\n </v:shape></xml>");
            writer.Flush();
            writer.Close();
        }
        // Generates content of spreadsheetPrinterSettingsPart2.
        private void GenerateSpreadsheetPrinterSettingsPart2Content(SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart2)
        {
            System.IO.Stream data = GetBinaryDataStream(spreadsheetPrinterSettingsPart7Data);
            spreadsheetPrinterSettingsPart2.FeedData(data);
            data.Close();
        }
        // Generates content of vmlDrawingPart5.
        private void GenerateVmlDrawingPart5Content(VmlDrawingPart vmlDrawingPart5)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(vmlDrawingPart5.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\r\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\r\n <o:shapelayout v:ext=\"edit\">\r\n  <o:idmap v:ext=\"edit\" data=\"25\"/>\r\n </o:shapelayout><v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\"\r\n  o:preferrelative=\"t\" path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\">\r\n  <v:stroke joinstyle=\"miter\"/>\r\n  <v:formulas>\r\n   <v:f eqn=\"if lineDrawn pixelLineWidth 0\"/>\r\n   <v:f eqn=\"sum @0 1 0\"/>\r\n   <v:f eqn=\"sum 0 0 @1\"/>\r\n   <v:f eqn=\"prod @2 1 2\"/>\r\n   <v:f eqn=\"prod @3 21600 pixelWidth\"/>\r\n   <v:f eqn=\"prod @3 21600 pixelHeight\"/>\r\n   <v:f eqn=\"sum @0 0 1\"/>\r\n   <v:f eqn=\"prod @6 1 2\"/>\r\n   <v:f eqn=\"prod @7 21600 pixelWidth\"/>\r\n   <v:f eqn=\"sum @8 21600 0\"/>\r\n   <v:f eqn=\"prod @7 21600 pixelHeight\"/>\r\n   <v:f eqn=\"sum @10 21600 0\"/>\r\n  </v:formulas>\r\n  <v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/>\r\n  <o:lock v:ext=\"edit\" aspectratio=\"t\"/>\r\n </v:shapetype><v:shape id=\"CH\" o:spid=\"_x0000_s25601\" type=\"#_x0000_t75\"\r\n  style=\'position:absolute;margin-left:0;margin-top:0;width:180pt;height:64.5pt;\r\n  z-index:1\'>\r\n  <v:imagedata o:relid=\"rId1\" o:title=\"FPR-VISION\"/>\r\n  <o:lock v:ext=\"edit\" rotation=\"t\"/>\r\n </v:shape><v:shape id=\"CF\" o:spid=\"_x0000_s25603\" type=\"#_x0000_t75\" style=\'position:absolute;\r\n  margin-left:0;margin-top:0;width:678pt;height:36pt;z-index:2\'>\r\n  <v:imagedata o:relid=\"rId2\" o:title=\"Footer\"/>\r\n  <o:lock v:ext=\"edit\" rotation=\"t\"/>\r\n </v:shape></xml>");
            writer.Flush();
            writer.Close();
        }
        // Generates content of spreadsheetPrinterSettingsPart4.
        private void GenerateSpreadsheetPrinterSettingsPart4Content(SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart4)
        {
            System.IO.Stream data = GetBinaryDataStream(spreadsheetPrinterSettingsPart4Data);
            spreadsheetPrinterSettingsPart4.FeedData(data);
            data.Close();
        }
        // Generates content of vmlDrawingPart6.
        private void GenerateVmlDrawingPart6Content(VmlDrawingPart vmlDrawingPart6)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(vmlDrawingPart6.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\r\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\r\n <o:shapelayout v:ext=\"edit\">\r\n  <o:idmap v:ext=\"edit\" data=\"10\"/>\r\n </o:shapelayout><v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\"\r\n  o:preferrelative=\"t\" path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\">\r\n  <v:stroke joinstyle=\"miter\"/>\r\n  <v:formulas>\r\n   <v:f eqn=\"if lineDrawn pixelLineWidth 0\"/>\r\n   <v:f eqn=\"sum @0 1 0\"/>\r\n   <v:f eqn=\"sum 0 0 @1\"/>\r\n   <v:f eqn=\"prod @2 1 2\"/>\r\n   <v:f eqn=\"prod @3 21600 pixelWidth\"/>\r\n   <v:f eqn=\"prod @3 21600 pixelHeight\"/>\r\n   <v:f eqn=\"sum @0 0 1\"/>\r\n   <v:f eqn=\"prod @6 1 2\"/>\r\n   <v:f eqn=\"prod @7 21600 pixelWidth\"/>\r\n   <v:f eqn=\"sum @8 21600 0\"/>\r\n   <v:f eqn=\"prod @7 21600 pixelHeight\"/>\r\n   <v:f eqn=\"sum @10 21600 0\"/>\r\n  </v:formulas>\r\n  <v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/>\r\n  <o:lock v:ext=\"edit\" aspectratio=\"t\"/>\r\n </v:shapetype><v:shape id=\"CH\" o:spid=\"_x0000_s10242\" type=\"#_x0000_t75\"\r\n  style=\'position:absolute;margin-left:0;margin-top:0;width:180pt;height:64.5pt;\r\n  z-index:1\'>\r\n  <v:imagedata o:relid=\"rId1\" o:title=\"FPR-VISION\"/>\r\n  <o:lock v:ext=\"edit\" rotation=\"t\"/>\r\n </v:shape><v:shape id=\"CF\" o:spid=\"_x0000_s10244\" type=\"#_x0000_t75\" style=\'position:absolute;\r\n  margin-left:0;margin-top:0;width:678pt;height:36pt;z-index:2\'>\r\n  <v:imagedata o:relid=\"rId2\" o:title=\"Footer\"/>\r\n  <o:lock v:ext=\"edit\" rotation=\"t\"/>\r\n </v:shape></xml>");
            writer.Flush();
            writer.Close();
        }
        // Generates content of vmlDrawingPart9.
        private void GenerateVmlDrawingPart9Content(VmlDrawingPart vmlDrawingPart7)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(vmlDrawingPart7.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\r\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\r\n <o:shapelayout v:ext=\"edit\">\r\n  <o:idmap v:ext=\"edit\" data=\"10\"/>\r\n </o:shapelayout><v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\"\r\n  o:preferrelative=\"t\" path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\">\r\n  <v:stroke joinstyle=\"miter\"/>\r\n  <v:formulas>\r\n   <v:f eqn=\"if lineDrawn pixelLineWidth 0\"/>\r\n   <v:f eqn=\"sum @0 1 0\"/>\r\n   <v:f eqn=\"sum 0 0 @1\"/>\r\n   <v:f eqn=\"prod @2 1 2\"/>\r\n   <v:f eqn=\"prod @3 21600 pixelWidth\"/>\r\n   <v:f eqn=\"prod @3 21600 pixelHeight\"/>\r\n   <v:f eqn=\"sum @0 0 1\"/>\r\n   <v:f eqn=\"prod @6 1 2\"/>\r\n   <v:f eqn=\"prod @7 21600 pixelWidth\"/>\r\n   <v:f eqn=\"sum @8 21600 0\"/>\r\n   <v:f eqn=\"prod @7 21600 pixelHeight\"/>\r\n   <v:f eqn=\"sum @10 21600 0\"/>\r\n  </v:formulas>\r\n  <v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/>\r\n  <o:lock v:ext=\"edit\" aspectratio=\"t\"/>\r\n </v:shapetype><v:shape id=\"CF\" o:spid=\"_x0000_s10248\" type=\"#_x0000_t75\"\r\n  style=\'position:absolute;margin-left:0;margin-top:0;width:651pt;height:36pt;\r\n  z-index:1\'>\r\n  <v:imagedata o:relid=\"rId1\" o:title=\"Footer\"/>\r\n  <o:lock v:ext=\"edit\" rotation=\"t\"/>\r\n </v:shape><v:shape id=\"CH\" o:spid=\"_x0000_s10250\" type=\"#_x0000_t75\" style=\'position:absolute;\r\n  margin-left:0;margin-top:0;width:144.75pt;height:69pt;z-index:2\'>\r\n  <v:imagedata o:relid=\"rId2\" o:title=\"FPR-letterhead\"/>\r\n  <o:lock v:ext=\"edit\" rotation=\"t\"/>\r\n </v:shape></xml>");
            writer.Flush();
            writer.Close();
        }
        // Generates content of vmlDrawingPart1.
        // Generates content of vmlDrawingPart1.
        private void GenerateVmlDrawingPart10Content(VmlDrawingPart vmlDrawingPart10)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(vmlDrawingPart10.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\r\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\r\n <o:shapelayout v:ext=\"edit\">\r\n  <o:idmap v:ext=\"edit\" data=\"1\"/>\r\n </o:shapelayout><v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\"\r\n  o:preferrelative=\"t\" path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\">\r\n  <v:stroke joinstyle=\"miter\"/>\r\n  <v:formulas>\r\n   <v:f eqn=\"if lineDrawn pixelLineWidth 0\"/>\r\n   <v:f eqn=\"sum @0 1 0\"/>\r\n   <v:f eqn=\"sum 0 0 @1\"/>\r\n   <v:f eqn=\"prod @2 1 2\"/>\r\n   <v:f eqn=\"prod @3 21600 pixelWidth\"/>\r\n   <v:f eqn=\"prod @3 21600 pixelHeight\"/>\r\n   <v:f eqn=\"sum @0 0 1\"/>\r\n   <v:f eqn=\"prod @6 1 2\"/>\r\n   <v:f eqn=\"prod @7 21600 pixelWidth\"/>\r\n   <v:f eqn=\"sum @8 21600 0\"/>\r\n   <v:f eqn=\"prod @7 21600 pixelHeight\"/>\r\n   <v:f eqn=\"sum @10 21600 0\"/>\r\n  </v:formulas>\r\n  <v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/>\r\n  <o:lock v:ext=\"edit\" aspectratio=\"t\"/>\r\n </v:shapetype><v:shape id=\"CH\" o:spid=\"_x0000_s1025\" type=\"#_x0000_t75\"\r\n  style=\'position:absolute;margin-left:0;margin-top:0;width:202.5pt;height:80.25pt;\r\n  z-index:1\'>\r\n  <v:imagedata o:relid=\"rId1\" o:title=\"fpr-cs-portal-report\"/>\r\n  <o:lock v:ext=\"edit\" rotation=\"t\" aspectratio=\"f\"/>\r\n </v:shape><v:shape id=\"CF\" o:spid=\"_x0000_s1026\" type=\"#_x0000_t75\" style=\'position:absolute;\r\n  margin-left:0;margin-top:0;width:651pt;height:36pt;z-index:2\'>\r\n  <v:imagedata o:relid=\"rId2\" o:title=\"Footer\"/>\r\n  <o:lock v:ext=\"edit\" rotation=\"t\"/>\r\n </v:shape></xml>");
            writer.Flush();
            writer.Close();
        }
        // Generates content of vmlDrawingPart1.
        

        #region Binary Data

      //   private string imagePart2Data = "iVBORw0KGgoAAAANSUhEUgAAAMEAAABcCAYAAADXutEBAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAAyJpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADw/eHBhY2tldCBiZWdpbj0i77u/IiBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+IDx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDUuMy1jMDExIDY2LjE0NTY2MSwgMjAxMi8wMi8wNi0xNDo1NjoyNyAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iIHhtbG5zOnN0UmVmPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VSZWYjIiB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtcE1NOkRvY3VtZW50SUQ9InhtcC5kaWQ6RDE1NDk3RjFBRDFDMTFFMjhCMURCQUYzMDUzQ0UzODMiIHhtcE1NOkluc3RhbmNlSUQ9InhtcC5paWQ6RDE1NDk3RjBBRDFDMTFFMjhCMURCQUYzMDUzQ0UzODMiIHhtcDpDcmVhdG9yVG9vbD0iQWRvYmUgUGhvdG9zaG9wIENTNiAoV2luZG93cykiPiA8eG1wTU06RGVyaXZlZEZyb20gc3RSZWY6aW5zdGFuY2VJRD0ieG1wLmRpZDowRTI3MTY3MzFDQURFMjExOUJDRThEQjBDRDMwRUQ4MSIgc3RSZWY6ZG9jdW1lbnRJRD0ieG1wLmRpZDowRTI3MTY3MzFDQURFMjExOUJDRThEQjBDRDMwRUQ4MSIvPiA8L3JkZjpEZXNjcmlwdGlvbj4gPC9yZGY6UkRGPiA8L3g6eG1wbWV0YT4gPD94cGFja2V0IGVuZD0iciI/PvvJdzAAAAq8SURBVHja7F1Lcuu4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";
         
        private string imagePart2Data = "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ICAgICA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIgogICAgICAgICAgICB4bWxuczpkYz0iaHR0cDovL3B1cmwub3JnL2RjL2VsZW1lbnRzLzEuMS8iPgogICAgICAgICA8ZGM6Zm9ybWF0PmltYWdlL3BuZzwvZGM6Zm9ybWF0PgogICAgICA8L3JkZjpEZXNjcmlwdGlvbj4KICAgPC9yZGY6UkRGPgo8L3g6eG1wbWV0YT4KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAKPD94cGFja2V0IGVuZD0idyI/Pl+Dw5wAACAASURBVHic7d15nGNVmf/xT6p6p+mGXmi2JOzoBR0aWQQRcVcY15FlXHCZxMjihjIi47ihgzIy8lMHCImgIjAgiyKoIKKACtLK2l6gWSuB7rZ3unqvJb8/nnO5N+nqZumue6qS7/v1qld3Janck5vkPGd5zrmZRqOBiIiIpK/LdwFEREQ6lYKwiIiIJwrCIiIinigIi4iIeDLGdwFk6wmzgbLshvbZoB6el7whzAaHA5cA+/gp0lbVAOru5wngfuA+4ClgIbA2qIeDw3XwMBuMBz4LnD1cx/BkA3YOVwH/AJ4BHgfmAfOxc7sMWAMMBPVQ3z950RSERdpDFtgVeDVwPNCHBeT7gHvCbHAX8HhQDxf7K+KoMxbYy/2/kfjpAnqBGvAIFpgfDrPB/VigXhbUw4H0iyujkYKwyOiXafm3GwsgAbAv8D6sV/e3MBtcC9wW1MO5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";

       
        private string spreadsheetPrinterSettingsPart1Data = "UwBlAG4AZAAgAFQAbwAgAE8AbgBlAE4AbwB0AGUAIAAyADAAMQAwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEEAAbcAAwDAy8AAAIAAQDqCm8IZAABAA8AWAICAAEAWAIDAAEATABlAHQAdABlAHIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAABAAAAAgAAAAEAAAD/////AAAAAAAAAAAAAAAAAAAAAERJTlUiANAADAMAAMKskFEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABQAAAAAABwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAAAAAAA0AAAAFNNVEoAAAAAEADAAFMAZQBuAGQAIABUAG8AIABNAGkAYwByAG8AcwBvAGYAdAAgAE8AbgBlAE4AbwB0AGUAIAAyADAAMQAwACAARAByAGkAdgBlAHIAAABSRVNETEwAVW5pcmVzRExMAFBhcGVyU2l6ZQBMRVRURVIAT3JpZW50YXRpb24AUE9SVFJBSVQAUmVzb2x1dGlvbgBEUEk2MDAAQ29sb3JNb2RlADI0YnBwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==";

    //    private string spreadsheetPrinterSettingsPart2Data = "UwBlAG4AZAAgAFQAbwAgAE8AbgBlAE4AbwB0AGUAIAAyADAAMQAwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEEAAbcAAwDAy8AAAIAAQDqCm8IZAABAA8AWAICAAEAWAIDAAEATABlAHQAdABlAHIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAABAAAAAgAAAAEAAAD/////AAAAAAAAAAAAAAAAAAAAAERJTlUiANAADAMAAMKskFEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABQAAAAAABwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAAAAAAA0AAAAFNNVEoAAAAAEADAAFMAZQBuAGQAIABUAG8AIABNAGkAYwByAG8AcwBvAGYAdAAgAE8AbgBlAE4AbwB0AGUAIAAyADAAMQAwACAARAByAGkAdgBlAHIAAABSRVNETEwAVW5pcmVzRExMAFBhcGVyU2l6ZQBMRVRURVIAT3JpZW50YXRpb24AUE9SVFJBSVQAUmVzb2x1dGlvbgBEUEk2MDAAQ29sb3JNb2RlADI0YnBwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==";

        private string spreadsheetPrinterSettingsPart3Data = "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AAAAAAAAAAAAAAAAAAAAAACSAAAAQwBPAE4ARgBJAEQARQBOAFQASQBBAEwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEMATwBOAEYASQBEAEUATgBUAEkAQQBMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABIAAAAAAAAAJABAAAAQQByAGkAYQBsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAICAgAAAAAAAAAAAAMIBAAAAAAAHAAcABwAEAAcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAARABlAGYAYQB1AGwAdAAgAFMAZQB0AHQAaQBuAGcAcwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAEAAAAAAAQAAAAAAIFAAAAHAAABGAAAAAACAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAcAAgABAAAAAAAAAQAAAAAAAAAAAAAAAAAAAQAAAAAAAACUAQAAEQABAAEAAAAAAAAAAAAAAAkAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMADAAACAAIAAgAEAAIAAgACAAIABAABAAEAf38AAAEAAQACAAIAAAEBBAABAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAUAAAAWAAAAAAAAAAAAAAAAAAAAAAABAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==";

        private string spreadsheetPrinterSettingsPart4Data = "UwBlAG4AZAAgAFQAbwAgAE8AbgBlAE4AbwB0AGUAIAAyADAAMQAwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEEAAbcAAwDAy8AAAIAAQDqCm8IZAABAA8AWAICAAEAWAIDAAEATABlAHQAdABlAHIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAABAAAAAgAAAAEAAAD/////AAAAAAAAAAAAAAAAAAAAAERJTlUiANAADAMAAMKskFEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABQAAAAAABwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAAAAAAA0AAAAFNNVEoAAAAAEADAAFMAZQBuAGQAIABUAG8AIABNAGkAYwByAG8AcwBvAGYAdAAgAE8AbgBlAE4AbwB0AGUAIAAyADAAMQAwACAARAByAGkAdgBlAHIAAABSRVNETEwAVW5pcmVzRExMAFBhcGVyU2l6ZQBMRVRURVIAT3JpZW50YXRpb24AUE9SVFJBSVQAUmVzb2x1dGlvbgBEUEk2MDAAQ29sb3JNb2RlADI0YnBwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==";

        private string spreadsheetPrinterSettingsPart5Data = "UwBlAG4AZAAgAFQAbwAgAE8AbgBlAE4AbwB0AGUAIAAyADAAMQAwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEEAAbcAAwDAy8AAAIAAQDqCm8IZAABAA8AWAICAAEAWAIDAAEATABlAHQAdABlAHIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAABAAAAAgAAAAEAAAD/////AAAAAAAAAAAAAAAAAAAAAERJTlUiANAADAMAAMKskFEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABQAAAAAABwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAAAAAAA0AAAAFNNVEoAAAAAEADAAFMAZQBuAGQAIABUAG8AIABNAGkAYwByAG8AcwBvAGYAdAAgAE8AbgBlAE4AbwB0AGUAIAAyADAAMQAwACAARAByAGkAdgBlAHIAAABSRVNETEwAVW5pcmVzRExMAFBhcGVyU2l6ZQBMRVRURVIAT3JpZW50YXRpb24AUE9SVFJBSVQAUmVzb2x1dGlvbgBEUEk2MDAAQ29sb3JNb2RlADI0YnBwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==";

        private string spreadsheetPrinterSettingsPart6Data = "UwBlAG4AZAAgAFQAbwAgAE8AbgBlAE4AbwB0AGUAIAAyADAAMQAwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEEAAbcAAwDAy8AAAIAAQDqCm8IZAABAA8AWAICAAEAWAIDAAEATABlAHQAdABlAHIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAABAAAAAgAAAAEAAAD/////AAAAAAAAAAAAAAAAAAAAAERJTlUiANAADAMAAMKskFEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABQAAAAAABwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAAAAAAAA0AAAAFNNVEoAAAAAEADAAFMAZQBuAGQAIABUAG8AIABNAGkAYwByAG8AcwBvAGYAdAAgAE8AbgBlAE4AbwB0AGUAIAAyADAAMQAwACAARAByAGkAdgBlAHIAAABSRVNETEwAVW5pcmVzRExMAFBhcGVyU2l6ZQBMRVRURVIAT3JpZW50YXRpb24AUE9SVFJBSVQAUmVzb2x1dGlvbgBEUEk2MDAAQ29sb3JNb2RlADI0YnBwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==";

        private string spreadsheetPrinterSettingsPart7Data = "TgBQAEkAQwBBADQANwA4ADIAIAAoAEgAUAAgAEwAYQBzAGUAcgBKAGUAdAAgAFAAcgBvAGYAZQBzAHMAAAAAAAEEAwDcADQDD9sAAAIAAQDqCm8IAAABAAcAWAIBAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFNERE0ABgAAAAYAAEhQIExhc2VySmV0IFByb2Zlc3Npb25hbCBNMTIxMm4AAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAABAAAAAQAAAAEAAAABAAAAAQAAAAEAAAABAAAAAQAAAAEAAAABAAAAAQAAAAEAAAABAAAAAQAAAAEAAAABAAAAAQAAAAEAAAAAAAAAAQAAAAEAAAAaBDAAAAAAAAAAAAAAAAAADwAAAFoAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAICAgAD/AAAA//8AAAD/AAAA//8AAAD/AP8A/wAAAAAAAAAAAAAAAAAAAAAAKAAAAGQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA3gMAAN4DAAAAAAAAAAAAAACAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAK/TnjQDAAAAAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=";



        // private string imagePart2Data = "iVBORw0KGgoAAAANSUhEUgAAASwAAABrCAMAAAD6iU8nAAA0dGdJRnhYTVAgRGF0YVhNUD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4KPHg6eG1wbXRhIHhtbG5zOng9ImFkb2JlOm5zOm1ldGEvIiB4OnhtcHRrPSJBZG9iZSBYTVAgQ29yZSA1LjAtYzA2MCA2MS4xMzQ3NzcsIDIwMTAvMDIvMTItMTc6MzI6MDAgICAgICAgICI+ICAgPHJkZjpSRCB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPgogICAgICA8cmRmOkRzY3JpcHRpb24gcmRmOmFib3V0PSIiCiAgICAgICAgICAgIHhtbG5zOnhtcE1NPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvbW0vIgogICAgICAgICAgICB4bWxuczpzdGVmPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VSZWYjIj4KICAgICAgICAgPHhtcE1NOk9yaWdpbmFsRG9jdW1udElEPnhtcC5kaWQ6ODEwMDUzNDdGMTI2REYxMUFCQUE5QzNDN0EyMkUyRjM8L3htcE1NOk9yaWdpbmFsRG9jdW1lbnRJRD4KICAgICAgICAgPHhtcE1NOkRvY3VtZW50SUQ+eHAuZGlkOkYyRTIxNjhFQjE5NTExRTI4MTg3OENGREVDOUUzMjMwPC94bXBNTTpEb2N1bWVudElEPgogICAgICAgICA8eG1wTU06SW5zdGFuY2VJRD54bXAuaWlkOkYyRTIxNjhEQjE5NTExRTIxODc4Q0ZERUM5RTMyMzA8L3htcE1NOkluc3RhbmNlSUQ+CiAgICAgICAgIDx4bXBNTTpEZXJpdmRGcm9tIHJkZjpwYXJzZVR5cGU9IlJlc291cmNlIj4KICAgICAgICAgICAgPHN0UmVmOmluc3RhbmNlSUQ+eG1wLmlpZDo3NTFENTdEQzBBNDNERjExQjg5REQ5RjhGNjZCQUQxPC9zdFJlZjppbnN0YW5jZUlEPgogICAgICAgICAgICA8c3RSZWY6ZG9jdW1lbnRJRD54bXAuZGQ6ODEwMDUzNDdGMTI2REYxMUFCQUE5QzNDN0EyMkUyRjM8L3N0UmVmOmRvY3VtZW50SUQ+CiAgICAgICAgIDwveG1wTU06RGVyaXZlZEZyb20+CiAgICAgIDwvcmRmOkRlc2NyaXB0aW4+CiAgICAgIDxyZGY6RGVzY3JpcHRpb24gcmRmOmFib3V0PSIiCiAgICAgICAgICAgIHhtbG5zOnhtcD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLyI+CiAgICAgICAgIDx4bXA6Q3JlYW9yVG9vbD5BZG9iZSBGaXJld29ya3MgQ1M1IDExLjAuMC40ODQgV2luZG93czwveG1wOkNyZWF0b3JUb29sPgogICAgICAgICA8eG1wOkNyZWF0ZURhdGU+MjAxMy0wNS0wMVQwNDowNjowMFo8L3htcDpDZWF0ZURhdGU+CiAgICAgICAgIDx4bXA6TW9kaWZ5RGF0ZT4yMDEzLTA1LTAxVDA0OjA2OjAwWjwveG1wOk1vZGlmeURhdGU+CiAgICAgIDwvcmRmOkRlc2NyaXB0aW9uPgogICAgICA8cmRmOkRlc2NycHRpb24gcmRmOmFib3V0PSIiCiAgICAgICAgICAgIHhtbG5zOmRjPSJodHRwOi8vcHVybC5vcmcvZGMvZWxlbWVudHMvMS4xLyI+CiAgICAgICAgIDxkYzpmb3JtYXQ+aW1hZ2UvZ2lmL2RjOmZvcm1hdD4KICAgICAgPC9yZGY6RGVzY3JpcHRpb24+CiAgIDwvcmRmOlJERj4KPC94OnhtcG1lYT4KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgCj94cGFja2V0IGVuZD0idyI/PgH//v38+/r5+Pf29fTz8vHw7+7t7Ovq6ejn5uXk4+Lh4N/e3dzb2tnY19XU09LR0M/OzczLysnIx8bFxMPCwcC/vr28u7q5uLe2tbSzsrGwr66trKuqqainpqWko6KhoJ+enZybmpmYl5aVlJOSkZCPjo2Mi4qJiIeGhYSDgoGAf359fHt6eXh3dnV0c3JxcG9ubWxramloZ2ZlZGNiYWBfXl1cW1pZWFdWVVRTUlFQT05NTEtKSUhHRkVEQ0JBQD8+PTwAOjk4NzY1NDMyMTAvLi0sKyopKCcmJSQjIiEgHx4dHBsaGRgXFhUUExIREA8ODQwLCgkIBwYFBAMCAQATW5VDAAABgFBMVEX///////7//v/+///+//7+/v/+/v7//v73/////f7+/v39/v79/f3+/fz8/Pz++/zv///7+/v9+fn6+fr39/f+9vb98/Pz8/P87/D87O3g8/Ls7Oz75ufb7u364+Xn5+j53uDi4uP429342dv/1tje3t/319n31NfZ29z60NPY2Nn2zc/T09T1xsrQ0NHA1NPNzc7Ly8zzv8K7zszIycrIx8jJxsnFxcbyt7vCw8TAwMG+vr/wsLW6urzwq7G4uLnvp624tbe1tbbupqyzs7SwsLGsrK3smaCdr62mpqeioqSgn6Dph42bnJ2WlpiRkpPmd3+MjI2Hh4nja3OCgoR/f4B8fX/hXWZ6eXvhWGJ1dXdycnTfUltvcHJwbnBsbG5nZ2ndR1JhYWPbPEdcXF5XV1lTU1TYKzdQUFNPTE1LS07WHyzVHSlGRUfVHCnWHCnUHCjWGyvUGyjWGyjVGCXUEyHUEB06OTrTDBnSCRfRBRPQAhDQAA3OAAMrKCoTEBGI95p8AAAAAWJLR0QAiAUdSAAAAARnSUZnAAAABwLoke4AAAAMY21QUEpDbXAwNzEyAAAAA0gAc7wAAAocSURBVHhe7Zv7X9pYFsDvfnbbmd1PO62tU+zQFXFkxqJFUVxEa+lawQHlkQiIREwMCSCY0qkJhEeTf31vQCEJCNi6Tm97zw/24eGS++W87jkXALBgApgAJoAJYAKYACaACWACmMD3R+DJs3Hk0aNHQ9SemLENXfQRupBXTnOnoyR3NA2mj6/XOzrYX/9Nj2B92JqH8fjum4XnKCJbl2RplDQL02C2XL9eTW40Coe7vf2/awxbUpYbrfe5gxX0cK2XpcookXPTYKYgnUO5VlVqNnNvr7b/tjpEs72EJLfEQ4M1ooDutmBVKmK9Hv9nR95W30MZ8RlUWwXUjOv2YEFzacQfPnwITWQsWJWKXH6Fgj31nvE2YVWk1n575TFhVeTCLFK0bhEWDGlydeEmsCrNw+8VlhajGkc3glWtIuWIZsu6GCCDsqF4JVqO1ImkmZbZDbvKomhKqa2Dv2uCiH2ZYVUHSAvWWdeUDrCe6pC6Sn2yFrX6YtYlTVGsanlSlybl3AtEQGmPaYQlFg4GyGH8hQmWVFh5MQ1lYfdQlMvl8sXFFYDqcR+sxrGm2lZfj5frhhJMlFHyQyMsOXfd52y0LKnQNYjXlSqkVb6CJRWmzJZliOILuaqh/mp1K1kELOxzYf3a3dtbSX8GkLTayeiGnaB/JQvnop5Wo1NsoCFfDgucQs/qxizpfH04LHAk62HVvzNY+9WPHz9eARDF1wZYMJ4bLQvEm98zrHdaQdCFdW6E1a29uo4Wb+hhNXfR8MD2U96CG+5K3VwIz4d/joJ1bHDDFvRaZOQWYB3LPS+sSGXYShgW4F9XdWihOkp9GjOsZ8+mTPLsH5qYS4deNlyXJFj190qHmaGlw0qhqm+KNY/7etJfsZ2ZitJyX4s510lXZlgQSUfWC/KfULpFqVapmSzreKrzCcyu7x5Kkr6CFxsohSxTzNI6eLJRPnWqJDOs359AmXr15gBW8Npxsgsr3gerUsl1pHDRbOpygRb7cz9/xYbU92gjWzSNgwGwKpVCW8oXdannVVp7tK4FbNPZEH4CHelmgsuzoiTBbICQjITVHAzrciIhGsrxSqWaezbAsowNZo1pR5Aq3/tKhwFt82ssa3CDXWy90wxlvE4p7EL/gJBZjQPrGssaDKt1CpvwDx+OA0tsSiiddDrJbNQo7CaW1Sh0qqbhsDQHlFry6Ru0zGocWONaltaBL7Q78KNgQaOsFg7b/oqYjOyUfurMFEydUuNcUEtuYrWR+/1y8+ZsKEmGPCBW4jOIYeo8rhlWpyTQiTg4G0Lz6GU16FZwgB/vzu9NsGBvUGpq7eduoENyeG+GVc3NzJrktxftkYK5B6+7zCDLzeZ5bl83AjTCkk9XVnYPq/oOqSTLBwheDfm8g7RWlOYuzS93fBjffaWVV10xwuq0lV+XYftZl0NbR+jdPfo8WFcDi/YY4ue+XQ/sOqyIumIfUqvHkYtbnwer2htYDNrx4BbNfkPfnKlIdeQS4ufBkgq9Fs0AWmZYf2vLVM7QI63IF1fZExUTuxNYlzB25fZAupsUm8eIha27hAVO5fMPHz70onwDpZ5yX501/pD1Rm545WZvZX3rC17QOkWqnXUbA4v+iHNdD/7JqXEcXaki1Si9W1hgpS4a7oXUYfsLNvhxgO9eltEPWY+buksk2j1UpMqHOw3wAKzAO0eGqw45dKc7/+cAD93toNEbbmjU6ihFrTu2LPBKNwnSYEkFhA7Udw0LxBvGbx60rwoiIncOa6ZsHAhVc+iY1p3DMl05qogIzcMgLN395KEB/rKVrJ3uxjpId6eDhpt/YLpQP9c3WesFZHrM6+cSPK5dyVBYupQ/AtY7+EWnnjSOjFe397XK9P37rsLl1zIQiFuv5aYsw8l654LDp8J1jzx70dLdgWiKQ28KvWvpr0x8OjUu+rygXwo2pSVUrh1NLbycm3u5sDDX/vnq2g7TkwVNoC78OTczN/w7N8/bym31ubnZl7/eu2fA9QK+k/aO7T+grKAT4h+vLv0CwC9L2s+RsrT0IwA/rv7xx/KDBw962j5iDdy/f9/88sdLS6ur/zb/r2VVWwTMrz7u/saysWa1j3zzr0AhJJRiwEIVS97RD+MsMjZg5VVBUdN6bVaNDHrxRonn8vmw6VeLXMYCgJUpLvZg5dUaP/rt/3oNguETIMCfsFarx7MIrB671e0Ak641txUAx7LFsfzU4vY6AbB5bN48ce9eSD2z/EcQHPbt7WWXfx483fBFor5Np93vBBa/y7K1/JMnsAlfDEKllC1YytpdlmWrzeOyLS5OAvvyRtg96XXMc9kHdp8HsvdYHY6Nmpr661GMfAILTTIEoDJUIpTNprhQkM+kuZQvyaRoaj7AMWTqxJemKW7HmeUyqWIIgC1Vpf1WS0wRzmo1NQn2FKHEkMrZmVLy84rAKMkTgVM4SCtW3APbJYpkEhmCYdgkk7U+pZhMepvgMymO2GGyND0f5BJ0aNJht/R58chnv3MFRz5CEV4+xIfXIp6NIhkrpQMMsxkMOFN8iGV8kTzF8D4bzWTz4S02r/lqVFFVatLL/DcI/8JYarWUWkyrZ7GaIqjUdk2lVM4Z9kOboVmGLlFBlichHU+kmOEcfiFKsuRZbOOEpfiMdVOIposnISTiFdz5RskfYRjSXwy6Y2SaS5D8ooWhfWEywTBEPgi2eJajiRjDcdTEBM06gNVps4cFNRGg+bwiKBlSDUfVDK8uWhUoViAoQUatsWsAOHmOiIYdbj5piZTWwFqJYbaotDPDc7wNxOC6mShZpFnOBR9jQHq4c7sZ/YZ7JddWPu/cY1KljCvGZ7LMv9w8zRa3fCxDFCOuNJfk0m6vZ5vLAl+esgC/IlgAoypqcTOjJpiawltYhVRqP3jVkqIEoVOuuXyUKgDg56E6AIFSABAlH4gWSZrh5l3FLMNZbQxLF/c8PvcOl/pp9FN+HRoTbMkWqpEWnooK9DZdpIU0iNQSbDEcEShnhs2w3GYqHyLo7aSQZmokAHZBhe4mJFQhKqjRqApNSlUotfhDUiWjSk1QhJpCcCrUJDV1ADK1DbBR5DIlwsHXYnD1aKxEZYT0Fp8NZgiio4SETO6FJtyx+cnojjWU3NkJB6MesBmdd5PRzUjAueFdZmmLPZIgAhOLZGQn4oabspM8n55/SnJEOOv1ZMPAliYC1A4I0jvLkeWESvrpPBOEFwB3Yh6oPhGO2gDYJJOhSRCKWsFW1GmPxfzRNeAlklHnJqEpfQOSpLcJIaIrPkfsyauc7CkKjFZPv4HN33QLrjTLRLSwM6ZMkjAhBsZU/vbUJm9oI5MWZKL1F31YMHn3TneGf/Qt21EcrmN60Y2Uv2gf+MWYACaACWACmAAmgAlgApgAJoAJYAKYACaACWACmAAmgAlgApgAJoAJYAKYACaACWACmAAmgAlgApgAJoAJYAKYwNdI4H+YV448vfBK+gAAAABJRU5ErkJggg==";

        // private string spreadsheetPrinterSettingsPart8Data = "TgBQAEkAQwBBADQANwA4ADIAIAAoAEgAUAAgAEwAYQBzAGUAcgBKAGUAdAAgAFAAcgBvAGYAZQBzAHMAAAAAAAEEAwDcADQDD9sAAAIAAQDqCm8IAAABAAcAWAIBAAEAAAAAAAAAAAAAAAAAAACIa1nj/gcAAAAAAAAAAAAAyGtZ4/4HAAAAAAAAAAAAAIiZWOP+BwAAAAAAAAAAAAAQxVjj/gcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFNERE0ABgAAAAYAAEhQIExhc2VySmV0IFByb2Zlc3Npb25hbCBNMTIxMm4AAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAABAAAAAQAAAAEAAAABAAAAAQAAAAEAAAABAAAAAQAAAAEAAAABAAAAAQAAAAEAAAABAAAAAQAAAAEAAAABAAAAAQAAAAEAAAAAAAAAAQAAAAEAAAAaBDAAAAAAAAAAAAAAAAAADwAAAFoAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAICAgAD/AAAA//8AAAD/AAAA//8AAAD/AP8A/wAAAAAAAAAAAAAAAAAAAAAAKAAAAGQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA3gMAAN4DAAAAAAAAAAAAAACAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAK/TnjQDAAAAAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=";

        #endregion
        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

      
       
       
        

      
    }
}
using System.Configuration;
using System.Drawing;

using databaseAPI;

using EASendMail;

using GNAchartingtools;

using GNAgeneraltools;

using GNAspreadsheettools;

using GNAsurveytools;

using OfficeOpenXml;


namespace BBVTGR
{
    class Program
    {
        static void Main()
        {


#pragma warning disable CS0162

#pragma warning disable CS8600
#pragma warning disable CS8601
#pragma warning disable CS8602
#pragma warning disable CS8604
#pragma warning disable IDE0059


            gnaTools gnaT = new();
            GNAsurveycalcs gnaSurvey = new();
            dbAPI gnaDBAPI = new();
            spreadsheetAPI gnaSpreadsheetAPI = new();

            //==== System config variables

            string strFreezeScreen = ConfigurationManager.AppSettings["freezeScreen"];
            string strExcelPath = ConfigurationManager.AppSettings["ExcelPath"];
            string strExcelFile = ConfigurationManager.AppSettings["ExcelFile"];
            string strWorkbookFullPath = strExcelPath + strExcelFile;
            string strReferenceWorksheet = ConfigurationManager.AppSettings["ReferenceWorksheet"];
            string strFirstDataRow = ConfigurationManager.AppSettings["FirstDataRow"];
            string strFirstDataCol = ConfigurationManager.AppSettings["FirstDataCol"];
            string strFirstOutputRow = ConfigurationManager.AppSettings["FirstOutputRow"];

            //==== Set the EPPlus license
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            //====[ Main Program ]====================================================================================

            gnaT.WelcomeMessage("prismOffset 2023.03.03");

            // instantiate the lists
            var prism = new List<Prism>();

            // read the prism data from the reference worksheet into a list prismData
            // pointname, Esurvey, Nsurvey, Enow, Nnow, 
            // iPrismCount = prismData.Count()

            int iStartRow = Convert.ToInt16(strFirstDataRow);
            int iStartCol = Convert.ToInt16(strFirstDataCol);
            int iRow = iStartRow;
            int iCol = iStartCol;
            string strName, strReplacementName;
            double dblEsurvey, dblNsurvey, dblEnow, dblNnow;

            FileInfo newFile = new(strWorkbookFullPath);

            using (ExcelPackage package = new(newFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[strReferenceWorksheet];

                do
                {
                    strName = Convert.ToString(worksheet.Cells[iRow, iCol].Value).Trim();
                    strReplacementName = Convert.ToString(worksheet.Cells[iRow, iCol+15].Value).Trim();
                    dblEsurvey = Math.Round(Convert.ToDouble(worksheet.Cells[iRow, iCol + 1].Value),4);
                    dblNsurvey = Math.Round(Convert.ToDouble(worksheet.Cells[iRow, iCol + 2].Value),4);
                    dblEnow = Math.Round(Convert.ToDouble(worksheet.Cells[iRow, iCol + 10].Value),4);
                    dblNnow = Math.Round(Convert.ToDouble(worksheet.Cells[iRow, iCol + 11].Value),4);

                    if ((strName != "") || (strName != "None"))
                    {
                        prism.Add(new Prism() { Name = strName, Eref = dblEsurvey, Nref = dblNsurvey, Enow = dblEnow, Nnow= dblNnow, Note1 = "empty", ReplacementName = strReplacementName });
                    }
                    iRow++;
                    strName = Convert.ToString(worksheet.Cells[iRow, iCol].Value);
                } while (strName != "");
            }

            prism.Add(new Prism() { Name = "TheEnd", ReplacementName = "TheEnd" });






            goto ThatsAllFolks;




            // Loop i = 1 to i = iPrismCount-2 
            //  Ay = prismsData[i-1].Esurvey
            //  Ax = prismsData[i-1].Nsurvey
            //  ByRef = prismsData[i].Esurvey
            //  BxRef = prismsData[i].Nsurvey
            //  ByNow = prismsData[i].Enow
            //  BxNow = prismsData[i].Nsnow
            //  Cy = prismsData[i+1].Esurvey
            //  Cx = prismsData[i+1].Nsurvey
            //  compute base bearing AC
            //      bearingAC= Join(Ay,Ax,Cy,Cx)
            //  Compute transverse bearing
            //      bearingACperp = bearingAC-(pi/2)
            //      correct if <0
            // Compute bearing between reference location and current location of the prism being investigated
            //      bearingBrefToBnow, disBrefToBnow = Join(ByRef,BxRef,ByNow,BxNow)
            // Compute perpenducular displacement     
            //      slewAngle= bearingBrefToBnow - bearingACperp
            //      adjust if > (Pi/4)
            //      slewDistance = disBrefToBnow*cos(slewAngle)
            // Compute the sign




            {






                double dblSlew = 0.0;
                //
                // Purpose:
                //      To compute rail slew (horizontal displacement from reference position) for all prisms
                //      Write the slew to the reference worksheet
                // Input:
                //      Focus prism: Reference coordinates, current coordinates, rail bracket
                //      Next prism: Reference coordinates
                //      These are read off the reference worksheet: columns (3,4,12,13,20)
                // Algoriths:
                //      Intersection  
                //      The sign convention agreed will produce positive Horizontal Alignment faults in case of shifts
                //      towards the Left looking at High mileage.
                //      Sign convention set by strPositiveLeftFlag
                // Output:
                //      Current slew
                //      This is written to the Reference worksheet column 24
                // Useage:
                //     computeSlew(strExcelWorkbookFullPath, strReferenceWorksheet, strFirstDataRow, strPositiveLeftFlag)
                //

                // Read in all the prisms
                //Prism[] prism = readPointCoordinates(strExcelWorkbookFullPath, strReferenceWorksheet, strFirstDataRow, strCoordinateOrder);












                //string strName = "";
                //int iPrismCounter = 0;
                //strName = prism[0].Name;

                //double dblYA, dblXA, dblYB, dblXB, dblYC, dblXC;
                //double dblRailBearing = 0.0;
                //double dlbOffsetBearing = 0.0;
                //int iRow = Convert.ToInt32(strFirstDataRow);

                //FileInfo workingFile = new(strExcelWorkbookFullPath);

                //using (ExcelPackage package = new(workingFile))
                {

                    //using (ExcelWorksheet referenceWorksheet = package.Workbook.Worksheets[strReferenceWorksheet])
                    //{

                    //    do
                    //    {
                    //        dblYA = prism[iPrismCounter].Eref;  // reference location of prism
                    //        dblXA = prism[iPrismCounter].Nref;  // reference location of prism
                    //        dblYB = prism[iPrismCounter].E;     // current location of prism
                    //        dblXB = prism[iPrismCounter].N;     // current location of prism

                    //        if (prism[iPrismCounter].Track != "Rail Start")
                    //        {
                    //            dblYC = prism[iPrismCounter - 1].Eref;
                    //            dblXC = prism[iPrismCounter - 1].Nref;
                    //            var answer1 = gnaSurvey.Join(dblYC, dblXC, dblYA, dblXA);
                    //            dblRailBearing = answer1.Item1;             // the reference bearing of the rail
                    //            answer1 = gnaSurvey.Join(dblYC, dblXC, dblYB, dblXB);
                    //            dlbOffsetBearing = answer1.Item1;           // the bearing to the offset prism
                    //        }
                    //        else
                    //        {
                    //            dblYC = prism[iPrismCounter + 1].Eref;
                    //            dblXC = prism[iPrismCounter + 1].Nref;
                    //            var answer2 = gnaSurvey.Join(dblYA, dblXA, dblYC, dblXC);
                    //            dblRailBearing = answer2.Item1;             // the reference bearing of the rail
                    //            answer2 = gnaSurvey.Join(dblYB, dblXB, dblYC, dblXC);
                    //            dlbOffsetBearing = answer2.Item1;           // the bearing to the offset prism, but generating the wrong sign
                    //        }

                    //        dblSlew = Math.Round(Math.Pow(Math.Pow(dblYB - dblYA, 2) + Math.Pow(dblXB - dblXA, 2), 0.5), 3);



                    //        if (Math.Abs(dblRailBearing - dlbOffsetBearing) > 3.14)     // To catch where the track is lying almost exactly due north
                    //        {
                    //            if (dblRailBearing < dlbOffsetBearing)
                    //            {
                    //                dblRailBearing = dblRailBearing + 6.28318530717958;
                    //            }
                    //            else
                    //            {
                    //                dlbOffsetBearing = dlbOffsetBearing + 6.28318530717958;
                    //            }
                    //        }

                    //        if (dlbOffsetBearing > dblRailBearing)
                    //        {
                    //            dblSlew = -dblSlew;
                    //        }

                    //        if (prism[iPrismCounter].Track == "Rail Start")
                    //        {
                    //            dblSlew = -dblSlew;
                    //        }

                    //        referenceWorksheet.Cells[iRow, 24].Value = dblSlew;

                    //        iRow++;
                    //        iPrismCounter++;
                    //        strName = prism[iPrismCounter].Name;

                    //    } while (strName != "NoMore");

                    //    try
                    //    {
                    //        referenceWorksheet.Calculate();
                    //        package.Save();
                    //        package.Dispose();
                    //    }
                    //    catch (Exception ex)
                    //    {
                    //        Console.WriteLine("");
                    //        Console.WriteLine("Error:");
                    //        Console.WriteLine("computeSlew: " + strExcelWorkbookFullPath);
                    //        Console.WriteLine("Close the workbook and re-run the software.");
                    //        Console.WriteLine("\n" + ex);
                    //        Console.WriteLine("\nPress any key to exit..");
                    //        Console.ReadKey();
                    //        Environment.Exit(0);
                    //    }





                    //}




                }

            }







ThatsAllFolks:



            gnaT.freezeScreen(strFreezeScreen);
            Environment.Exit(0);
            Console.WriteLine("");
            Console.WriteLine("Task Complete....");
        }
    }


    public class Prism
    {
        public string? Name { get; set; }
        public string? ReplacementName { get; set; }
        public double Nref { get; set; }
        public double Eref { get; set; }
        public double Href { get; set; }
        public double Nnow { get; set; }
        public double Enow { get; set; }
        public double Hnow { get; set; }
        public double SlewBearing { get; set; }
        public double SlewDistance { get; set; }
        public string? Note1 { get; set; }
    }





}
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Drawing;
using System.Drawing.Imaging;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using FilePath = System.IO.Path;
using static System.Net.Mime.MediaTypeNames;
using Path = System.IO.Path;
using Image = System.Drawing.Image;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Drawing.Drawing2D;
using Rectangle = System.Drawing.Rectangle;

namespace GenerateTPSReport
{
    public partial class Form1 : Form
    {

        private string[] files = null;
        private string basedir = null;
        

        static string filePath = @"C:\Users\C74340\OneDrive - Microchip Technology Inc\Desktop\Final Report\DPT_TPS.xlsx";

        List<string> fileNameList = new List<string>();
        List<string> DPTImageFileList = new List<string>();
        List<string> FuncImageList = new List<string>();
        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// On Select Folder Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Select_Folder_Click(object sender, EventArgs e)
        {

            // get the selected folder
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    basedir = fbd.SelectedPath;
                    files = Directory.GetFiles(fbd.SelectedPath);
                }
            }


            // grab all the filenames and add it to the list
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < files.Length; i++)
            {
                string newname = FilePath.GetFileName(files[i]).Split(new char[] {'.'})[0];
                if (!newname.Contains("CH1") && !newname.Contains("CH2"))
                {
                    // check if the filename is repeated, if repeated do not add to the list.
                    if (!sb.ToString().Contains(newname))
                    {
                        sb.AppendLine(newname);
                    }
                }
            }


            fileNameList = sb.ToString().Split(new char[] { '\n' }).ToList();

            
           // File.WriteAllText(Directory.GetParent(basedir) + "\\Filename.txt", sb.ToString());
            sb.Clear();

        }


        /// <summary>
        /// Clone a given sheet into a new sheet
        /// </summary>
        static void CloneSheet(string clonedSheetName)
        {
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filePath, true))
            { 
                WorkbookPart workbookPart = spreadSheet.WorkbookPart;
                WorksheetPart sourceSheetPart = GetWorkSheetPart(workbookPart, "template");
                Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();

                SpreadsheetDocument tempSheet = SpreadsheetDocument.Create(new MemoryStream(), spreadSheet.DocumentType);
                WorkbookPart tempWorkbookPart = tempSheet.AddWorkbookPart();
                WorksheetPart tempWorksheetPart = tempWorkbookPart.AddPart(sourceSheetPart);
                WorksheetPart clonedSheet = workbookPart.AddPart(tempWorksheetPart);

                Sheet copiedSheet = new Sheet();
                copiedSheet.Name = clonedSheetName;
                copiedSheet.Id = workbookPart.GetIdOfPart(clonedSheet);
                copiedSheet.SheetId = (uint)sheets.ChildElements.Count + 1;
                sheets.Append(copiedSheet);
            }
        }

        /// <summary>
        /// Get Part Id
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        static WorksheetPart GetWorkSheetPart(WorkbookPart workbookPart, string sheetName)
        { 
            //Get the relationship id of the sheetname
            string relId = workbookPart.Workbook.Descendants<Sheet>() .Where(s => s.Name.Value.Equals(sheetName)) .First() .Id;
            return (WorksheetPart)workbookPart.GetPartById(relId); 
        }


        /// <summary>
        /// Generates Report for TPS
        /// </summary>
        private void GenerateExcelFile(string WorkSheetName, string[] ImageFilePath)
        {

            // Open the document and read all the lines into array of strings


            string SheetName = WorkSheetName.Trim('\r');

            // Add multiple images or apply further changes
            try
            {
                // Open spreadsheet
               SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, true);

                // Get WorksheetPart
                WorksheetPart worksheetPart = ExcelUtility.GetWorksheetPartByName(spreadsheetDocument, SheetName);

                // Add DPT image
                ExcelUtility.AddImage(worksheetPart, ImageFilePath[0], "DPT Image", 2, 7);  //B7


                // Add Functional Test Image

                ExcelUtility.AddImage(worksheetPart, ImageFilePath[1], "Func Image", 22, 7);
                ExcelUtility.AddImage(worksheetPart, ImageFilePath[2], "Func Image CH1", 29, 7);
                ExcelUtility.AddImage(worksheetPart, ImageFilePath[3], "Func Image CH2", 26, 25);

                // Other operations if needed

                worksheetPart.Worksheet.Save();

                spreadsheetDocument.Dispose();
            }
            catch (Exception ex) { }
        }


        /// <summary>
        /// Get all the DPT image files into a list for further use
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SelectDPTImageFileFolder_Click(object sender, EventArgs e)
        {
            // get the selected folder
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    DPTImageFileList = Directory.GetFiles(fbd.SelectedPath).ToList();
                }
            }
        }


        /// <summary>
        /// Get all the Func image files into a list for further use
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SelectFuncImageFolder_Click(object sender, EventArgs e)
        {
            // get the selected folder
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    FuncImageList = Directory.GetFiles(fbd.SelectedPath).ToList();
                }
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void GenerateReport_Click(object sender, EventArgs e)
        {


            // Add Sheets to the existing Excel File

            for (int i = 0; i < fileNameList.Count; ++i)
            {
                string filename = fileNameList[i].Trim('\r');

                if (!string.IsNullOrEmpty(filename))
                {
                    CloneSheet(filename);
                }
            }

                for (int i = 0; i < fileNameList.Count; ++i)
            {
                int index = i;

                string[] ImageFilePath = new string[4];

                // Add DPT image
                ImageFilePath[0] = DPTImageFileList[index];

                // Add Func Image
                ImageFilePath[1] = FuncImageList[index++];
                ImageFilePath[2] = FuncImageList[index++];
                ImageFilePath[3] = FuncImageList[index];

                string filename = fileNameList[i].Trim('\r');

                if (!string.IsNullOrEmpty(filename))
                {
                    GenerateExcelFile(filename, ImageFilePath);
                }

            }
            
        }


        /// <summary>
        /// Resize Images
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ResizeImage_Click(object sender, EventArgs e)
        {
            string filepath = "C:\\Users\\C74340\\OneDrive - Microchip Technology Inc\\Desktop\\Final Report\\DPT Edited";
            int width = 500;
            int height = 340;
            for (int i = 0; i < DPTImageFileList.Count; i++)
            {
                ImageUtility.ResizeImage(Image.FromFile(DPTImageFileList[i]), width, height).Save(Path.Combine(filepath, FilePath.GetFileName(DPTImageFileList[i])), System.Drawing.Imaging.ImageFormat.Jpeg);
            }
        }
    }

}
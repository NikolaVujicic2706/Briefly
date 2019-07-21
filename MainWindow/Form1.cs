using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Doc.Interface;
using Spire.Doc.Documents;
using Spire.Doc.Utilities;
using Spire.Doc.Collections;
using Spire.Doc.Fields;
using TextBox = Spire.Doc.Fields.TextBox;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace MainWindow
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		private void Button_is_clicked(object sender, EventArgs e)
		{
			OpenFileDialog openFileDialog = new OpenFileDialog();
			openFileDialog.Title = "Please select Word file";
            openFileDialog.Filter = "Word Files|*.docx;*.doc";
            openFileDialog.ShowDialog();
			word_path_textBox.Text = openFileDialog.FileName;
		}

		private void Attach_Text_File(object sender, EventArgs e)
		{
			OpenFileDialog openFileDialog = new OpenFileDialog();
			openFileDialog.Title = "Please select Excel file";
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.ShowDialog();
			excel_path_textBox.Text = openFileDialog.FileName;
		}


		private void Choose_Destionation_Folder(object sender, EventArgs e)
		{
			FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
			         if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
			         {
				     destination_textBox.Text = folderBrowserDialog.SelectedPath;
			         }
		}

		private async void Extract_Button_Clicked(object sender, EventArgs e)
		{
            if (word_path_textBox.Text == "" || excel_path_textBox.Text == "" || destination_textBox.Text =="" || fibre_ref_textBox.Text =="")
                {
                MessageBox.Show("Please, fill in all the fields!");
                }
               else {
                   try
                     {
                    status_label.Text = "Processing...";
                    await Task.Run(() => Extract_Images());
                    status_label.Text = "The images have been extracted!";
                }
                     catch (Exception excel_not_found_exeption)
                     {
                        MessageBox.Show(excel_not_found_exeption.Message);
                     }
               }
        }

        //method which fetch sorted array of chamber strings
        public string[] Fetch_Sorted_Chambers()
        {
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(excel_path_textBox.Text);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                // Find the number of rows
                int rowCount = xlWorksheet.Cells.Find("*", System.Reflection.Missing.Value,
                                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

                string[] chambers_id = new string[rowCount / 2];
                for (int i = 0; i < rowCount / 2; i++)
                {
                    chambers_id[i] = xlRange.Cells[2 * (i + 1), 2].Value.ToString();
                }
                xlWorkbook.Close();
                //this code snippet sorting chambers_id array
                if (rowCount / 2 != 2 || rowCount / 2 != 1)
                {
                    string temp = null;
                    for (int i = 1; i < (rowCount / 2) - 1; i += 4)
                    {
                        temp = chambers_id[i];
                        chambers_id[i] = chambers_id[i + 1];
                        chambers_id[i + 1] = temp;
                    }
                }

                return chambers_id;
        }

        //Method which clear all text box when delete buton is clicked
        private void Delete_button_Click(object sender, EventArgs e)
        {
            word_path_textBox.Clear();
            excel_path_textBox.Clear();
            destination_textBox.Clear();
            fibre_ref_textBox.Clear();
            status_label.Text = "";

        }
        //Main logic of application method
        public void Extract_Images()
        {

            String folderLocation = destination_textBox.Text;
            try
            {
                Document document = new Document(word_path_textBox.Text);
                int index = 1;
                int folder_number = 0;
                string[] sorted_chambers = Fetch_Sorted_Chambers();
                //Get Each Section of Document 
                foreach (Section section in document.Sections)
                {
                    //Get Each Paragraph of Section 
                    foreach (Paragraph paragraph in section.Paragraphs)
                    {
                        //Get Each Document Object of Paragraph Items 
                        foreach (DocumentObject docObject in paragraph.ChildObjects)
                        {
                            //If DocumentObjectType is TextBox
                            if (docObject.DocumentObjectType == DocumentObjectType.TextBox)
                            {
                                TextBox textbox = docObject as TextBox;
                                //Get each Document Object from textbox body
                                for (int i = 0; i < 2; i++)
                                {
                                    foreach (DocumentObject textboxdocObject in textbox.Body.ChildObjects)
                                    {
                                        foreach (DocumentObject childObject in textboxdocObject.ChildObjects)
                                        {
                                            //If Type is Picture.
                                            if (childObject.DocumentObjectType == DocumentObjectType.Picture)
                                            {
                                                //save the pictures to the folder and the subfolder
                                                if (i == 0)
                                                {
                                                    DocPicture textBoxPicture = childObject as DocPicture;
                                                    String fibre = fibre_ref_textBox.Text;
                                                    String main_folder_name = sorted_chambers[folder_number];
                                                    String subfolder_Name = String.Format(@"{0}\{1}\{2}", folderLocation, main_folder_name, fibre);
                                                    bool exist = Directory.Exists(main_folder_name);
                                                    if (exist)
                                                    {
                                                        Directory.Delete(main_folder_name);
                                                    }
                                                    Directory.CreateDirectory(main_folder_name);
                                                    Directory.CreateDirectory(subfolder_Name);
                                                    String imageName = String.Format(@"{0}\{1}\ClosedLid.jpg", folderLocation, main_folder_name);
                                                    String imageName1 = String.Format(@"{0}\{1}\{2}\ClosedLid.jpg", folderLocation, main_folder_name, fibre);
                                                    textBoxPicture.Image.Save(imageName, System.Drawing.Imaging.ImageFormat.Jpeg);
                                                    textBoxPicture.Image.Save(imageName1, System.Drawing.Imaging.ImageFormat.Jpeg);
                                                    index++;
                                                }
                                                else
                                                {
                                                    DocPicture textBoxPicture = childObject as DocPicture;
                                                    String fibre = fibre_ref_textBox.Text;
                                                    String main_folder_name = sorted_chambers[folder_number];
                                                    String imageName = String.Format(@"{0}\{1}\OpenLid.jpg", folderLocation, main_folder_name);
                                                    String imageName1 = String.Format(@"{0}\{1}\{2}\OpenLid.jpg", folderLocation, main_folder_name, fibre);
                                                    textBoxPicture.Image.Save(imageName, System.Drawing.Imaging.ImageFormat.Jpeg);
                                                    textBoxPicture.Image.Save(imageName1, System.Drawing.Imaging.ImageFormat.Jpeg);
                                                    index++;
                                                }
                                                i++;
                                            }
                                        }
                                    }
                                }
                                folder_number++;
                            }
                        }
                    }
                }
            }
            catch (System.Exception file_not_found_exeption)
                {
                MessageBox.Show(file_not_found_exeption.Message);
            }
           
        }

        //the method allows only number as valid entries in the fibre reference field
        private void Fibre_ref_textBox_TextChanged(object sender, EventArgs e)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(fibre_ref_textBox.Text, "[^0-9]"))
            {
                MessageBox.Show("Please, enter only numbers in the Fibre reference field.");
                fibre_ref_textBox.Text = fibre_ref_textBox.Text.Remove(fibre_ref_textBox.Text.Length - 1);
            }
        }

    }
}
	


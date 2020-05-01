using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML;
using System.Windows.Forms;
using System.IO;

namespace WindowsFormsApp4
{
    public partial class Form1 : Form
    {
        public Form1() //initialize Form and populate existing consignments into drop down on tab Hazourdous consignments
        {
            InitializeComponent();

            GetFiles();
            
        }


        public void GetFiles() //get existing consignments
        {
            // Put all file names in root directory into array.
            string[] array1 = Directory.GetFiles(@"C:\TEST CONSIGNMENTS", "*", SearchOption.AllDirectories);
            string[] array2 = Directory.GetFiles(@"S:\Haz Waste Notes In Out\Hazardous Waste General In", "*", SearchOption.AllDirectories);



            foreach (string name in array1)
            {
                
                ExistingConsignments.Items.Add(name);
            }

            foreach (string name in array2)
            {
                ExistingConsignments.Items.Add(name);
            }

        }

        public void Method(string a, DateTime date) //prints 2 consignments with date and consingment code generated
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process
            {
                StartInfo = new System.Diagnostics.ProcessStartInfo()
                {
                    CreateNoWindow = true,
                    Verb = "print",
                    FileName = a
                }
            };//create print process



            string consignment = File.ReadAllText(@"C:\consignment.txt");//read consignment codes
            int code = Convert.ToInt32(consignment);
            progressBar1.Visible = true;
            progressBar1.Value = 25;

            var workbook = new ClosedXML.Excel.XLWorkbook(a); // load the existing excel file
            var worksheet = workbook.Worksheets.Worksheet(1);


            worksheet.Cell("R6").SetValue(code);//set consignment in cell
            worksheet.Cell("D2").SetValue(date);
            workbook.Save();


            progressBar1.Value = 35;

            p.Start();//print file copy 1
            p.WaitForExit();




            worksheet.Cell("R6").SetValue(code);//set consignment in cell for copy2
            workbook.Save();

            progressBar1.Value = 55;


            System.IO.StreamWriter sw = new System.IO.StreamWriter(@"C:\consignment.txt");//save/update consignment
            code++;//code is consignment number
            
            sw.WriteLine(code);
            sw.Close();
            progressBar1.Value = 75;



            var workbook2 = new ClosedXML.Excel.XLWorkbook(@"C:\Users\karl\Desktop\note_numbers.xlsx"); // load the existing excel file
            var worksheet2 = workbook2.Worksheets.Worksheet(1);

            worksheet2.Cell("A2").SetValue(code);
            workbook2.Save();
           
            
            //for (int i = 0; i < number.Value; i++)
            //{
            //    code++;
            //    listBox1.Items.Add(code);

            //}


            progressBar1.Value = 100;


            if (numberCopies.Value == 2) {
                p.Start();//print file copy 2
                p.WaitForExit();
            }

           
            
            progressBar1.Value = 0;
            progressBar1.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e) //print consignment button on tab Weekly Runs
        {


            

            if (monthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Monday) {

                Method(@"c:\TEST CONSIGNMENTS\Monday\Hazardous Waste NEW Template selby.xlsX", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Monday\Hazardous Waste NEW Template catterick.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Monday\Hazardous Waste NEW Template leeming bar.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Monday\Hazardous Waste NEW Template harrogate stonefall.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Monday\Hazardous Waste NEW Template northallerton.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Monday\Hazardous Waste NEW Template ripon.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Monday\Hazardous Waste NEW Template whitby.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Monday\Hazardous Waste NEW Template burniston.xlsx", monthCalendar1.SelectionRange.Start.Date);

            }


            if (monthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Tuesday)
            {

                Method(@"c:\TEST CONSIGNMENTS\Tuesday\Hazardous Waste NEW Template leyburn.xlsX", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Tuesday\Hazardous Waste NEW Template catterick.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Tuesday\Hazardous Waste NEW Template leeming bar.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Tuesday\Hazardous Waste NEW Template harrogate stonefall.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Tuesday\Hazardous Waste NEW Template northallerton.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Tuesday\Hazardous Waste NEW Template ripon.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Tuesday\Hazardous Waste NEW Template seamer carr.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Tuesday\Hazardous Waste NEW Template harrogate west.xlsx", monthCalendar1.SelectionRange.Start.Date);

            }

            if (monthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Thursday)
            {

                Method(@"c:\TEST CONSIGNMENTS\Thursday\Hazardous Waste NEW Template THOLTHORPE.xlsX", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Thursday\Hazardous Waste NEW Template SOWERBY.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Thursday\Hazardous Waste NEW Template stokesley.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Thursday\Hazardous Waste NEW Template tadcaster.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Thursday\Hazardous Waste NEW Template malton and norton.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Thursday\Hazardous Waste NEW Template harrogate west.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Thursday\Hazardous Waste NEW Template ripon.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Thursday\Hazardous Waste NEW Template harrogate stonefall.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Thursday\Hazardous Waste NEW Template skipton.xlsx", monthCalendar1.SelectionRange.Start.Date);

                MessageBox.Show("Consignments Have Finished Printing");
            }

            if (monthCalendar1.SelectionStart.DayOfWeek == DayOfWeek.Friday)
            {

                Method(@"c:\TEST CONSIGNMENTS\Friday\Hazardous Waste NEW Template wombleton.xlsX", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Friday\Hazardous Waste NEW Template whitby.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Friday\Hazardous Waste NEW Template burniston.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Friday\Hazardous Waste NEW Template settle.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Friday\Hazardous Waste NEW Template selby.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Friday\Hazardous Waste NEW Template thornton-le-dale.xlsx", monthCalendar1.SelectionRange.Start.Date);
                Method(@"c:\TEST CONSIGNMENTS\Friday\Hazardous Waste NEW Template skipton.xlsx", monthCalendar1.SelectionRange.Start.Date);

                MessageBox.Show("Consignments Have Finished Printing");
            }
            else {

                MessageBox.Show("No Runs on Wednesday and Weekends, Please Select Another Date");

                }

        }

        private void PrintExisting_Click(object sender, EventArgs e) //print an existing consignment from tab Hazourdous consignment
        {

            // FileDialog browser = new FileDialog();
            // string tempPath = "";

            //if (browser.ShowDialog() == DialogResult.OK)
            //{
            //   tempPath = browser.SelectedPath; // prints path

            //}

            var openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = @"S:\Haz Waste Notes In Out";

            openFileDialog1.RestoreDirectory = true;

            if(openFileDialog1.ShowDialog() == DialogResult.OK) {
                Console.WriteLine(openFileDialog1.FileName);
            }




            Method(openFileDialog1.FileName, dateTimePicker1.Value);

            MessageBox.Show("Consignments Have Finished Printing");

        }

        private void button2_Click(object sender, EventArgs e) //creates and saves a new consignment and refreshes list of  consignments in drop down
        {

            ExistingConsignments.Items.Clear();

            GetFiles();
            
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            
            System.Diagnostics.Process p = new System.Diagnostics.Process
            {
                StartInfo = new System.Diagnostics.ProcessStartInfo()
                {
                    CreateNoWindow = true,
                    Verb = "print",
                    FileName = "S:\\Cardboard Files\\Cardboard Season Tickets\\Waste Transfer Season Ticket" + " " + name.Text
                }
            };//create print process



            string consignment = File.ReadAllText(@"C:\wtn.txt");//read consignment codes
            int code = Convert.ToInt32(consignment);


            var workbook = new ClosedXML.Excel.XLWorkbook(@"S:\Cardboard Files\Cardboard Season Tickets\blank.xlsx"); // load the existing excel file
            var worksheet = workbook.Worksheets.Worksheet(1);


            worksheet.Cell("b5").SetValue(code);//set consignment in cell
            worksheet.Cell("b13").SetValue(frequency.Text);
            worksheet.Cell("b18").SetValue(name.Text);
            worksheet.Cell("e19").SetValue(post.Text);
            worksheet.Cell("e20").SetValue(add1.Text);
            worksheet.Cell("e21").SetValue(add2.Text);
            worksheet.Cell("e22").SetValue(add3.Text);
            workbook.SaveAs("S:\\Cardboard Files\\Cardboard Season Tickets\\Waste Transfer Season Ticket" + " " + name.Text + ".xlsx");

                    


           // p.Start();//print file copy 1
           //p.WaitForExit();



            code++;//code is consignment number

            System.IO.StreamWriter sw = new System.IO.StreamWriter(@"C:\wtn.txt");//save/update consignment
            


            sw.WriteLine(code);
            sw.Close();









        } //create new season ticket

        private void button3_Click(object sender, EventArgs e)
        {


            var workbook = new ClosedXML.Excel.XLWorkbook(@"S:\Cardboard Files\cardboard weights 2019.xlsx"); // load the existing excel file
            var worksheet = workbook.Worksheets.Worksheet(1);


            int weight = Convert.ToInt32(cardWeight.Text) / checkedListBox1.CheckedItems.Count;


            string consignment = File.ReadAllText(@"C:\cellNumber.txt");//read consignment codes
            int cellNumber = Convert.ToInt32(consignment);

            foreach (object itemChecked in checkedListBox1.CheckedItems)
            {
                
                worksheet.Cell("A" + cellNumber).SetValue(dateTimePicker2.Value.Date);//set consignment in cell
                worksheet.Cell("B" + cellNumber).SetValue(itemChecked.ToString());
                worksheet.Cell("C" + cellNumber).SetValue(weight);
                workbook.Save();

                cellNumber++;
                

      
            }

            System.IO.StreamWriter sw = new System.IO.StreamWriter(@"C:\cellNumber.txt");//save/update consignment
            

            sw.WriteLine(cellNumber);
            sw.Close();
        }
    }
}

using System;
using System.Activities;
using System.Collections.Generic;
using System.Data;
using System.Security;
using System.Text.RegularExpressions;
using UiPath.Activities.Contracts;
using UiPath.CodedWorkflows;
using UiPath.CodedWorkflows.Utils;
using UiPath.Core;
using UiPath.Core.Activities;
using UiPath.Core.Activities.Storage;
using UiPath.Orchestrator.Client.Models;
using UiPath.Platform.ResourceHandling;
using UiPath.Testing;
using UiPath.Testing.Activities.TestDataQueues.Enums;
using UiPath.Testing.Enums;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using UiPath.PDF.Activities;
using System.Linq;


namespace Sales_Contract
{
    public class Testing : CodedWorkflow
    {
        [Workflow]
        public void Execute (string salesContactPageIN, string salesQuote1TxtIN)
    
        {
            
            //return(inputFile: salesContactPageIN, salesQuote1Input: salesQuote1TxtIN);
          
            string inputFile = salesContactPageIN;
            string salesQuote1Input = salesQuote1TxtIN;
            
            
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = false;

            //open the input document
            object inputPathObj = inputFile;
            Microsoft.Office.Interop.Word.Document inputDoc = wordApp.Documents.Open(ref inputPathObj, ref missing);

            //create a new DataTable
            System.Data.DataTable dataTable = new System.Data.DataTable();

            //copy the first table
            if (inputDoc.Tables.Count >= 1)
                {
	                Microsoft.Office.Interop.Word.Table table = inputDoc.Tables[1];
	
	//loop through the table rows and columns to read the text from each cell
	for (int rowIndex = 1; rowIndex <= table.Rows.Count; rowIndex++)
	{
		DataRow dataRow = dataTable.NewRow();
		
		for (int columnIndex = 1; columnIndex <= table.Columns.Count; columnIndex++)
		{
			Microsoft.Office.Interop.Word.Cell cell = table.Cell(rowIndex, columnIndex);
			string cellText = cell.Range.Text.TrimEnd('\r', '\a'); //remove extra characters
			
			//create the DataTable columns on the first row iteration
			if (rowIndex == 1)
			{
				System.Data.DataColumn dataColumn = new System.Data.DataColumn(cellText);
				dataTable.Columns.Add(dataColumn);
			}
			else
			{
				dataRow[columnIndex - 1] = cellText;
			}
		}
		
		if(rowIndex > 1)
		{
			dataTable.Rows.Add(dataRow);
		}
	}
}

 
//****************VARIABLES*************************//

string dateTime = DateTime.Now.AddDays(45).ToString("MM/dd/yyyy");

//***************************CUSTOMER VARIABLES**********************
//Customer name
string customerName = dataTable.Rows[0][0].ToString();
//outCustomerName = customerName;


//Customer part of datatable
string customerAddress = dataTable.Rows[2][0].ToString();

//Customer Street Variable
//string customerStreet_init = dataTable.Rows[2][0].ToString();
string customerStreet_edit = System.Text.RegularExpressions.Regex.Match(customerAddress, @"Address:(.*?)City").Value.Trim();
string customerStreet = System.Text.RegularExpressions.Regex.Replace(customerStreet_edit, @"Address:|City", "");
//outCustomerStreet = customerStreet;

//Customer City
string customerCity_edit = System.Text.RegularExpressions.Regex.Match(customerAddress, @"City:(.*?)State").Value.Trim();
string customerCity = System.Text.RegularExpressions.Regex.Replace(customerCity_edit, @"City:|State", "");
//outCustomerCity = customerCity;

//Customer State
string customerState = System.Text.RegularExpressions.Regex.Match(customerAddress, @"AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|
MI|MN|MS|MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY").Value.Trim();
//string customerState = System.Text.RegularExpressions.Regex.Replace(customerState_edit, @"Province:|Country", "");
//outCustomerState = customerState;

//Customer Zip Code
string customerZip = System.Text.RegularExpressions.Regex.Match(customerAddress, @"\d{5}(-\d{4})?").Value.Trim();
//string customerZip = System.Text.RegularExpressions.Regex.Replace(customerZip_edit, @"Zip Code:", "");
//outCustomerZip = customerZip;


//Customer contact part of datatable
string customerContact = dataTable.Rows[4][0].ToString();

//Customer Contact Name
string customerContactName_edit = System.Text.RegularExpressions.Regex.Match(customerContact, @"Name:(.*?)Title").Value.Trim();
string customerContactName = System.Text.RegularExpressions.Regex.Replace(customerContactName_edit, @"Name:|Title", "");
//outCustomerContactName = customerContactName;

//Customer Title
string customerTitle_edit = System.Text.RegularExpressions.Regex.Match(customerContact, @"Title:(.*?)Telephone").Value.Trim();
string customerTitle = System.Text.RegularExpressions.Regex.Replace(customerTitle_edit, @"Title:|Telephone", "");
//outCustomerTitle = customerTitle;

//Customer Phone
string customerPhone_edit = System.Text.RegularExpressions.Regex.Match(customerContact, @"Telephone:(.*?)Fax").Value.Trim();
string customerPhone = System.Text.RegularExpressions.Regex.Replace(customerPhone_edit, @"Telephone:|Fax", "");
//outCustomerPhone = customerPhone;

//Customer Email
string customerEmail = System.Text.RegularExpressions.Regex.Match(customerContact, @"(?("")(""[^""]+?""@)|(([0-9a-zA-Z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-zA-Z])@))" +  
    @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-zA-Z][-0-9a-zA-Z]*[0-9a-zA-Z]*\.)+[a-zA-Z0-9][\-a-zA-Z0-9]{0,22}[a-zA-Z0-9]))").Value.Trim();
//outCustomerEmail = customerEmail;


//Customer Billing Address part of datatable
string customerBilling = dataTable.Rows[6][0].ToString();

//Customer Billing Street
string customerBillingStreet_edit = System.Text.RegularExpressions.Regex.Match(customerBilling, @"Address:(.*?)City").Value.Trim();
string customerBillingStreet = System.Text.RegularExpressions.Regex.Replace(customerBillingStreet_edit, @"Address:|City", "");
//outCustomerBillingStreet = customerBillingStreet;

//Customer Billing City
string customerBillingCity_edit = System.Text.RegularExpressions.Regex.Match(customerBilling, @"City:(.*?)State").Value.Trim();
string customerBillingCity = System.Text.RegularExpressions.Regex.Replace(customerBillingCity_edit, @"City:|State", "");
//outCustomerBillingCity = customerBillingCity;

//Customer Billing State
string customerBillingState = System.Text.RegularExpressions.Regex.Match(customerBilling, @"AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|
MI|MN|MS|MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY").Value.Trim();
//string customerState = System.Text.RegularExpressions.Regex.Replace(customerState_edit, @"Province:|Country", "");
//outCustomerBillingState = customerBillingState;

//Customer Billing Zip
string customerBillingZip = System.Text.RegularExpressions.Regex.Match(customerBilling, @"\d{5}(-\d{4})?").Value.Trim();
//string customerZip = System.Text.RegularExpressions.Regex.Replace(customerZip_edit, @"Zip Code:", "");
//outCustomerBillingZip = customerBillingZip;



//**********SALES VARIABLES*****************

//Sales Contact Name
string salesContactName = dataTable.Rows[0][2].ToString();
//outSalesContactName = salesContactName;

//Sales part of datatable
string salesAddress = dataTable.Rows[2][2].ToString();

//Sales Street Address
string salesStreet_edit = System.Text.RegularExpressions.Regex.Match(salesAddress, @"Address:(.*?)City").Value.Trim();
string salesStreet = System.Text.RegularExpressions.Regex.Replace(salesStreet_edit, @"Address:|City", "");
//outSalesStreet = salesStreet;

//Sales City
string salesCity_edit = System.Text.RegularExpressions.Regex.Match(salesAddress, @"City:(.*?)State").Value.Trim();
string salesCity = System.Text.RegularExpressions.Regex.Replace(salesCity_edit, @"City:|State", "");
//outSalesCity = salesCity;

//Sales State
string salesState_edit = System.Text.RegularExpressions.Regex.Match(salesAddress, @"AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|
MI|MN|MS|MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY").Value.Trim();
//string salesState = System.Text.RegularExpressions.Regex.Replace(salesState_edit, @"Province:|Country", "");
//outSalesState = salesState_edit;

//Sales Zip
string salesZip_edit = System.Text.RegularExpressions.Regex.Match(salesAddress, @"\d{5}(-\d{4})?").Value.Trim();
//string salesZip = System.Text.RegularExpressions.Regex.Replace(salesZip_edit, @"Zip Code:|Fax", "");
//outSalesZip = salesZip_edit;

//Sales Email
string salesEmail = System.Text.RegularExpressions.Regex.Match(salesAddress, @"(?("")(""[^""]+?""@)|(([0-9a-zA-Z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-zA-Z])@))" +  
    @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-zA-Z][-0-9a-zA-Z]*[0-9a-zA-Z]*\.)+[a-zA-Z0-9][\-a-zA-Z0-9]{0,22}[a-zA-Z0-9]))").Value.Trim();
//outSalesEmail = salesEmail;

//Sales Manager
string salesManager_edit = System.Text.RegularExpressions.Regex.Match(salesAddress, @"Mgr:(.*?)SCVP").Value.Trim();
string salesManager = System.Text.RegularExpressions.Regex.Replace(salesManager_edit, @"Mgr:|SCVP", "");
//outSalesManager = salesManager;

//Sales SCVP
string salesSCVP_edit = System.Text.RegularExpressions.Regex.Match(salesAddress, @"Name:\s*(.+)").Value.Trim();
string salesSCVP = System.Text.RegularExpressions.Regex.Replace(salesSCVP_edit, @"Name:", "");
//outSalesSCVP = salesSCVP;


//Engagement manager info
string engagementManager = dataTable.Rows[4][2].ToString();

//Engagement Manager Name
string engagementManagerName_edit = System.Text.RegularExpressions.Regex.Match(engagementManager, @"Name:(.*?)Address").Value.Trim();
string engagementManagerName = System.Text.RegularExpressions.Regex.Replace(engagementManagerName_edit, @"Name:|Address", "");
//outEngagementManager = engagementManagerName;

//Engagement Manager email
string engagementManagerEmail = System.Text.RegularExpressions.Regex.Match(engagementManager, @"(?("")(""[^""]+?""@)|(([0-9a-zA-Z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-zA-Z])@))" +  
    @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-zA-Z][-0-9a-zA-Z]*[0-9a-zA-Z]*\.)+[a-zA-Z0-9][\-a-zA-Z0-9]{0,22}[a-zA-Z0-9]))").Value.Trim();
//outEngagementManagerEmail = engagementManagerEmail;

                


inputDoc.Close(ref missing);
wordApp.Quit(ref missing);

System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);



//First broad capture of the sales quote using regex, starting at the line that contains item description
string regex1 = System.Text.RegularExpressions.Regex.Match(salesQuote1Input, "Item Description[\\s\\S]+?Final Quote").Value;

// Removing pricing details from regex1 to create regex2
string regex2 = System.Text.RegularExpressions.Regex.Replace(
    regex1,
    @"\d+\s*\d{1,3},\d{3}\.\d{2}\s*\s*\d{1,3},\d{3}\.\d{2}\s*\d{1,3}\.\d{2}|\d+\s*\d{3}\.\d{2}\s*\s*\d{3}\.\d{2}\s*\d{1,3}\.\d{2}|\d+\s*\d{3}\.\d{2}\s*\s*\d{1,3},\d{3}\.\d{2}\s*\d{1,3}\.\d{2}|\[[^\]]*\]",
    "");



// Build data tables
var sitesStatesDT = new System.Data.DataTable();
sitesStatesDT.Columns.Add("Column1", typeof(string));
//outSitesStatesDT = sitesStatesDT;

var sitesStatesFinal = new System.Data.DataTable();
sitesStatesFinal.Columns.Add("Column1", typeof(string));
//outSitesStatesFinal = sitesStatesFinal;

var dtSKU = new System.Data.DataTable(); 
dtSKU.Columns.Add("SKU Column", typeof(string));
//outdtSKU = dtSKU;

var dtPrices = new System.Data.DataTable();
dtPrices.Columns.Add("Qty", typeof(decimal));
dtPrices.Columns.Add("List Price", typeof(decimal)); 
dtPrices.Columns.Add("Total Price", typeof(decimal));
dtPrices.Columns.Add("MRC", typeof(decimal));
//outdtPrices = dtPrices;

var dtSKU2 = new System.Data.DataTable();
dtSKU2.Columns.Add("Column1", typeof(string));  
dtSKU2.Columns.Add("Column2", typeof(int));
//outdtSKU2 = dtSKU2;

var dtJoined3 = new System.Data.DataTable();
dtJoined3.Columns.Add("Description", typeof(string));
dtJoined3.Columns.Add("Item Cost", typeof(string));
dtJoined3.Columns.Add("Quantity", typeof(string));  
dtJoined3.Columns.Add("MRC", typeof(string));
//outdtJoined3 = dtJoined3;

var dtPrices2 = new System.Data.DataTable();
dtPrices2.Columns.Add("MRC Unit Price", typeof(string));
dtPrices2.Columns.Add("Units", typeof(double));
dtPrices2.Columns.Add("MRC Total ", typeof(string));
//outdtPrices2 = dtPrices2;

//************************************************************************************************************************************************//



// Line by line capture for descriptions, quantity and prices
var sitesStates = Regex.Matches(regex1, @"^.*", RegexOptions.IgnoreCase | RegexOptions.Multiline)
    .Cast<Match>()
    .Select(m => m.Value)
    .ToList();

// Line by Line capture to build datatable with descriptions only
var sitesStatesRegex2 = Regex.Matches(regex2, @"^.*", RegexOptions.IgnoreCase | RegexOptions.Multiline)
    .Cast<Match>()
    .Select(m => m.Value)
    .ToList();

// Row counter and counter variables for datatable construction via line by line regex
int rowCount = sitesStatesRegex2.Count;
int counter = 0;


while (counter <= rowCount - 1)
{
    
    sitesStatesDT.Rows.Add(sitesStatesRegex2[counter]);
    counter++;
}

// Equivalent to FilterDataTable activity - Removing rows
var sitesStatesFiltered = sitesStatesDT.AsEnumerable()
    .Where(row => !(row.Field<string>(0).StartsWith("Sub Total") ||
                    row.Field<string>(0).Contains("Shipping") ||
                    row.Field<string>(0).Contains("Item Description") ||
                    string.IsNullOrEmpty(row.Field<string>(0)) ||
                    row.Field<string>(0).Contains("Final") ||
                    row.Field<string>(0).StartsWith("Total") ||
                    row.Field<string>(0).Contains("Price")))
    .CopyToDataTable();

// Equivalent to Multiple Assign for Indices Array - Getting indices to be able to create spaces when
//doing the descriptions and pricing datatable joins
string strSplitTrigger = "Bundle SubTotal $";

int [] indexArray1 = Enumerable.Range(0, sitesStatesFiltered.Rows.Count)
    .Where(i => sitesStatesFiltered.Rows[i][0].ToString().Contains(strSplitTrigger))
    .Select(i => i).ToArray();

sitesStatesFinal = sitesStatesFiltered.AsEnumerable()
    .Select(d => {
        var newRow = sitesStatesFinal.NewRow();
        newRow["Column1"] = System.Text.RegularExpressions.Regex.Replace(d["Column1"].ToString(), "Bundle SubTotal|\\$|\\d{1,2},?\\d{3}\\.\\d{2}", "");
        return newRow;
    }).CopyToDataTable();

// Assuming regex1 is a string containing the input text for regex matches
var regexSKU = new System.Text.RegularExpressions.Regex(@"\[[^\]]*\]");
List<System.Text.RegularExpressions.Match> matchSKU = regexSKU.Matches(regex1).Cast<System.Text.RegularExpressions.Match>().ToList();

var regexPrices = new System.Text.RegularExpressions.Regex(@"\d+\s*\d{1,3},\d{3}\.\d{2}\s*\s*\d{1,3},\d{3}\.\d{2}\s*\d{1,3}\.\d{2}|\d+\s*\d{3}\.\d{2}\s*\s*\d{3}\.\d{2}\s*\d{1,3}\.\d{2}|\d+\s*\d{3}\.\d{2}\s*\s*\d{1,3},\d{3}\.\d{2}\s*\d{1,3}\.\d{2}");
List<System.Text.RegularExpressions.Match> matchesPrice = regexPrices.Matches(regex1).Cast<System.Text.RegularExpressions.Match>().ToList();

// For DataTable dtSKU and dtPrice, you should initialize them before or within this code,
// depending on whether they are being passed in as arguments or not.
// Then you would fill them with the matched values from matchSKU and matchesPrice.


// Building SKU Datatable
foreach (Match currentMatch in matchSKU)
{
    dtSKU.Rows.Add(currentMatch.Value.Split('_'));
}

// Building Price Datatable
foreach (Match currentMatch in matchesPrice)
{
    string currentMatchResults = Regex.Replace(currentMatch.Value, @"\s+", " ");
    dtPrices.Rows.Add(currentMatch.Value.Replace(" ", "_").Split('_'));
}

// Regex Removal for dtSKU2 to take care of any overlapping
System.Data.DataTable tempTable = dtSKU2.Clone();
foreach (DataRow d in dtSKU.AsEnumerable())
{
    string c = Regex.Replace(d["SKU Column"].ToString(), @"\d+\s+\d{1,3},\d{3}.\d{2}\s+\s+\d{1,3},\d{3}.\d{2}\s+\d{1,3}.\d{2}", "");
    DataRow newRow = tempTable.NewRow();
    newRow.ItemArray = new object[] { c };
    tempTable.Rows.Add(newRow);
}
dtSKU2 = tempTable.AsEnumerable().CopyToDataTable();

// Filter dtSKU2 directly, thus creating dtSKU3
for (int i = dtSKU2.Rows.Count - 1; i >= 0; i--)
{
    var row = dtSKU2.Rows[i];
    if (!row[0].ToString().Contains("[SKU]") || string.IsNullOrEmpty(row[0].ToString()))
    {
        dtSKU2.Rows.RemoveAt(i);
    }
}
var dtSKU3 = new System.Data.DataTable();
dtSKU3 = dtSKU2;

// Filter dtPrices directly, thus creating dtPrices1
for (int i = dtPrices.Rows.Count - 1; i >= 0; i--)
{
    var row = dtPrices.Rows[i];
    if (string.IsNullOrEmpty(row[0].ToString()))
    {
        dtPrices.Rows.RemoveAt(i);
    }
}

var dtPrices1 = new System.Data.DataTable();
dtPrices1 = dtPrices;





// Doing math for the Totals in Prices
foreach (DataRow currentRow in dtPrices1.Rows)
{
    // Perform the calculation and division as per the UiPath [Add Data Row] activity
    double value = Convert.ToDouble(currentRow[3]) / 0.74;
    int quantity = Convert.ToInt32(currentRow[0]);
    double result = Math.Round(value / quantity, 2);

    // Prepare the new row data
    object[] newRowData = new object[]
    {
         
        result.ToString(),
        currentRow[0].ToString(),
        Math.Round(value, 2).ToString(),
    };

    // Add the new DataRow to dtPrices2
    DataRow newRow = dtPrices2.NewRow();
    newRow.ItemArray = newRowData;
    dtPrices2.Rows.Add(newRow);
    
   
}


 //add a blank row at the top of dtPrices2 so it lines up with sitesStatesFjnal
    object[] initialRowData = new object[] { /* Your initial row data here */ };
    DataRow initialRow = dtPrices2.NewRow();
    initialRow.ItemArray = initialRowData;
    dtPrices2.Rows.InsertAt(initialRow, 0);



// This inserts blank rows based on the indexing from the "Bundle Subtotal $" text
foreach (int currentItem in indexArray1)
{
    DataRow newRow = dtPrices2.NewRow();
    dtPrices2.Rows.InsertAt(newRow, currentItem);
    dtPrices2.Rows.InsertAt(dtPrices2.NewRow(), currentItem + 1);
}

// This joins the 'sitesStatesFinal' datatable with the 'dtPrices2' - and also makes their rows line up
//DataTable dtJoined3 = new DataTable(); // Assuming this DataTable is already set up correctly

foreach (DataRow currentRow1 in sitesStatesFinal.Rows)
{
    foreach (DataRow currentRow2 in dtPrices2.Rows)
    {
        if (sitesStatesFinal.Rows.IndexOf(currentRow1) == dtPrices2.Rows.IndexOf(currentRow2))
        {
            // Assuming both rows have the same schema and can be concatenated directly
            DataRow joinedRow = dtJoined3.NewRow();
            joinedRow.ItemArray = currentRow1.ItemArray.Concat(currentRow2.ItemArray).ToArray();
            dtJoined3.Rows.Add(joinedRow);
        }
    }
}

// Assuming mrcSUM is a variable already declared
string mrcSUM = Math.Round(dtJoined3.AsEnumerable().Sum(row => 
    Convert.ToDouble(string.IsNullOrEmpty(row["MRC"].ToString()) ? "0" : row["MRC"].ToString())), 2).ToString();

// Adding a new row to 'dtJoined3'
DataRow newRowForTotal = dtJoined3.NewRow();
newRowForTotal[3] = "Total MRC: " + mrcSUM; // Replace 3 with the actual index or column name
dtJoined3.Rows.Add(newRowForTotal);

 //add a blank row at the top of dtPrices2 so it lines up with sitesStatesFjnal
object[] initialRowData_final = new object[] { /* Your initial row data here */ };
DataRow initialRow_final = dtJoined3.NewRow();
initialRow_final.ItemArray = initialRowData_final;
dtJoined3.Rows.InsertAt(initialRow_final, 0);





//*****Inserting tables and variables into the contract template************************************//


//Initializing/opening Microsoft Word via Interop
missing = System.Reflection.Missing.Value;

Microsoft.Office.Interop.Word.Application wordApp_1 = new Microsoft.Office.Interop.Word.Application();
wordApp_1.Visible = true;

//Opening cover page again
inputPathObj = inputFile;
inputDoc = wordApp_1.Documents.Open(ref inputPathObj, ref missing, ref missing, ref missing);

//copying the table in the cover page
if (inputDoc.Tables.Count >= 1)
{
	Microsoft.Office.Interop.Word.Table table = inputDoc.Tables[1];
	table.Range.Copy();
}


//Closing the coverpage
inputDoc.Close(ref missing, ref missing, ref missing);


// Get the current directory of the executing script
string currentDirectory = System.IO.Directory.GetCurrentDirectory();

// Construct the full path to the template file
string templateFile = System.IO.Path.Combine(currentDirectory, "sales_contract.docx");
//Opening the template document
//string templateFile = @"sales_contract.docx";

object templatePathObj = templateFile;
Microsoft.Office.Interop.Word.Document templateDoc = wordApp_1.Documents.Open(ref templatePathObj, ref missing, ref missing, ref missing);



//Inserting the cover page table from the cover page document, and fixing formatting
Microsoft.Office.Interop.Word.Range placeholderRange = templateDoc.Content;
object findText = "<<coverPage>>";

Microsoft.Office.Interop.Word.Find find = placeholderRange.Find;
find.ClearFormatting();

if(find.Execute(ref findText, ref missing, ref missing))
{
	placeholderRange.Paste();
	
	
	Microsoft.Office.Interop.Word.Table wordTable = templateDoc.Tables[1];
	placeholderRange.Font.Name = "Arial";
	placeholderRange.Font.Size = 8;
	placeholderRange.HighlightColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdNoHighlight;
	
	
	var specificCells = new[]
	{
		//new Tuple<int, int>(1,1),
		//new Tuple<int, int>(1,2),
		//new Tuple<int, int>(1,3),	
		new Tuple<int, int>(2,1),
		new Tuple<int, int>(2,2),
		new Tuple<int, int>(2,3),
		//new Tuple<int, int>(3,1),
		//new Tuple<int, int>(3,2),
		//new Tuple<int, int>(3,3),
		new Tuple<int, int>(4,1),
		new Tuple<int, int>(4,2),
		new Tuple<int, int>(4,3),
		//new Tuple<int, int>(5,1),
		//new Tuple<int, int>(5,2),
		//new Tuple<int, int>(5,3),
		new Tuple<int, int>(6,1),
		new Tuple<int, int>(6,2),
		new Tuple<int, int>(6,3),
		//new Tuple<int, int>(7,1),
		//new Tuple<int, int>(7,2),
		//new Tuple<int, int>(7,3),
		new Tuple<int, int>(8,1),
		new Tuple<int, int>(8,2),
		new Tuple<int, int>(8,3),	
	};
	
	//state abbreviations hash - Used to make sure state abbreviations are proper case
	HashSet<string> stateAbbreviations = new HashSet<string>
	{
		"AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA",
		"HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", 
		"MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
		"NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC",
		"SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY"
	};
	
	
	//loop through the cells and apply the case change
	foreach (Microsoft.Office.Interop.Word.Row row in wordTable.Rows)
	{
		foreach (Microsoft.Office.Interop.Word.Cell cell in row.Cells)
		{
			//check if this cell is in the list of specific cells
			if (specificCells.Contains(new Tuple<int, int>(row.Index, cell.ColumnIndex)))
			{
				
				//get the text in the cell
				string text = cell.Range.Text.Trim();
				
				//ensure there's a space after each colon
				text = text.Replace(": ", ":");
				text = text.Replace(":", ": ");
				
				//split the text into words
				string[] words = text.Split(' ');
				
				//process each word
				for (int i = 0; i < words.Length; i++)
				{
					string upperWord = words[i].ToUpper();
					
					//if the word is a state abbreviation, make it uppercase
					if (stateAbbreviations.Contains(upperWord))
					{
						words[i] = upperWord;
					}
					else if (words[i].Contains("@"))
					{
						words[i] = words[i].ToLower();
					}
					else
					{
						//otherwise, change the case to title case
						words[i] = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(words[i].ToLower());
					}
				}
				
				//join the words back together
				string newText = string.Join(" ", words);
				
				//regex for state abbreviatons
				foreach (string abbreviation in stateAbbreviations)
				{
					string pattern = @"\b" + abbreviation + @"\b";
					newText = Regex.Replace(newText, pattern, abbreviation.ToUpper(), RegexOptions.IgnoreCase);
				}
				
				//set the cell's text to the new text
				cell.Range.Text = newText;
			
		}
	}
        
        
}
//*********************************END COVER PAGE TABLE***********************************************//
	
	


//find the placeholder and replace it with the customerName variable content
Microsoft.Office.Interop.Word.Range placeholderRange1 = templateDoc.Content;

object findText1 = "<<customerName>>";
object replaceWith1 = customerName;
object replace1 = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
object findWrap1 = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
placeholderRange1.Find.Execute(ref findText1, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith1, ref replace1, ref missing, ref findWrap1, ref missing, ref missing);



//find the placeholder and replace it with the dateTime variable content
Microsoft.Office.Interop.Word.Range placeholderRange3 = templateDoc.Content;

object findText3 = "<<dateTime>>";
object replaceWith3 = dateTime;
object replace3 = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
object findWrap3 = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
placeholderRange3.Find.Execute(ref findText3, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith3, ref replace3, ref missing, ref findWrap3, ref missing, ref missing);





	
//*********************PRICING TABLE******************************************************************************************************************//	

//find the "<<dataTable>>" placeholder and insert the pricing table - dtJoined3
findText = "<<pricingTable>>";
find = templateDoc.Content.Find;
find.ClearFormatting();

if (find.Execute(ref findText))
{
	Microsoft.Office.Interop.Word.Range dataTableRange = (Microsoft.Office.Interop.Word.Range) find.Parent;
	dataTableRange.Select();
	
	rowCount = dtJoined3.Rows.Count;
	int columnCount = dtJoined3.Columns.Count;
	Microsoft.Office.Interop.Word.Table wordTable1 = templateDoc.Tables.Add(dataTableRange, rowCount+1, columnCount);

//center the entire table horizontally
wordTable1.Rows.Alignment = Microsoft.Office.Interop.Word.WdRowAlignment.wdAlignRowCenter;

//set the cell alignment to center
foreach (Microsoft.Office.Interop.Word.Cell cell in wordTable1.Range.Cells)
{
	cell.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
	cell.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
	
	//set the font and size
	cell.Range.Font.Name = "Arial";
	cell.Range.Font.Size = 8; 
}

//add column headers
for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
{
	wordTable1.Cell(1, columnIndex +1).Range.Text = dtJoined3.Columns[columnIndex].ColumnName;
}

//add data rows
for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
{
	for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
	{
		wordTable1.Cell(rowIndex+2, columnIndex+1).Range.Text = dtJoined3.Rows[rowIndex][columnIndex].ToString();
	}
	wordTable1.AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitContent);
 }

//set the table's border style

Microsoft.Office.Interop.Word.Border[] borders = new Microsoft.Office.Interop.Word.Border[6]
{
	

wordTable1.Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft],
wordTable1.Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight],
wordTable1.Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop],
wordTable1.Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom],
wordTable1.Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal],
wordTable1.Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical]
	
};

foreach (Microsoft.Office.Interop.Word.Border border in borders)
{
	border.LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
	border.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;
}

//set the preferred width of the table (value is in 
wordTable1.PreferredWidth = 700;

//set the width of the first column
wordTable1.Columns[1].Width = 100;
wordTable1.Columns[2].Width = 100;
wordTable1.Columns[3].Width = 100;
wordTable1.Columns[4].Width = 100;

foreach (Microsoft.Office.Interop.Word.Row row in wordTable1.Rows)
{
	Microsoft.Office.Interop.Word.Cell cell = row.Cells[1];
	cell.Range.Text = "\r\n" + cell.Range.Text;
}

}

//***************************************END OF PRICING TABLE***************************************//



string outputFile = System.IO.Path.Combine(currentDirectory, "sales_contract_endresult.docx");

//save the modified template document
object outputPathObj = outputFile;
templateDoc.SaveAs2(ref outputPathObj, ref missing);

templateDoc.Close(ref missing);
wordApp_1.Quit(ref missing);

//return(inputFile: salesContactPageIN, salesQuote1Input: salesQuote1TxtIN);

System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp_1);


	
        }
    }
}
    
    
}


using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Business_Intelligence___Assignment_2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private int GetDateID(string date)
        {
            // Remove the time from the date
            var dateSplit = date.Split(new char[0], StringSplitOptions.RemoveEmptyEntries);
            // Overwrite the original date value
            date = dateSplit[0];

            // Set array variable to store split version of formatted dates
            string[] arrayDate = date.Split('/');
            // Creates variables for day, month and year
            Int32 year = Convert.ToInt32(arrayDate[2]);
            Int32 month = Convert.ToInt32(arrayDate[1]);
            Int32 day = Convert.ToInt32(arrayDate[0]);

            // Use DateTime function to find out what day of the week it is
            DateTime MyDate = new DateTime(year, month, day);

            // Convert to database friendly format
            string dbDate = MyDate.ToString("M/dd/yyyy");

            // Create connection to destination database string
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            using (SqlConnection myconnection = new SqlConnection(connectionStringDestination))
            {
                // Open connection
                myconnection.Open();
                // Check if the date already exists in the database
                SqlCommand command = new SqlCommand("SELECT id FROM Time where date = @date", myconnection);
                command.Parameters.Add(new SqlParameter("date", dbDate));


                // Run the command and read the results
                using (SqlDataReader reader = command.ExecuteReader())
                {


                    // If there are rules then the date already exist so set exist to true
                    if (reader.HasRows)
                    {
                        if (reader.Read())
                        {
                            return Convert.ToInt32(reader["id"]);
                        }
                    }
                }
            }

            return 0;
        }
        private int GetCustomerID(string name)
        {

            // Create connection to destination database string
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            using (SqlConnection myconnection = new SqlConnection(connectionStringDestination))
            {
                // Open connection
                myconnection.Open();
                // Check if the date already exists in the database
                SqlCommand command = new SqlCommand("SELECT id FROM Customer where name = @name", myconnection);
                command.Parameters.Add(new SqlParameter("name", name));

                // Run the command and read the results
                using (SqlDataReader reader = command.ExecuteReader())
                {


                    // If there are rules then the date already exist so set exist to true
                    if (reader.HasRows)
                    {
                        if (reader.Read())
                        {
                            return Convert.ToInt32(reader["id"]);
                        }
                    }
                }

                return 0;
            }
        }
        private int GetProductID(string reference)
        {

            // Create connection to destination database string
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            using (SqlConnection myconnection = new SqlConnection(connectionStringDestination))
            {
                // Open connection
                myconnection.Open();
                // Check if the date already exists in the database
                SqlCommand command = new SqlCommand("SELECT id FROM Product where reference = @reference", myconnection);
                command.Parameters.Add(new SqlParameter("reference", reference));

                // Run the command and read the results
                using (SqlDataReader reader = command.ExecuteReader())
                {


                    // If there are rules then the date already exist so set exist to true
                    if (reader.HasRows)
                    {
                        if (reader.Read())
                        {
                            return Convert.ToInt32(reader["id"]);
                        }
                    }
                }
            }

            return 0;
        }

        private void SplitDates(string RawDate)
        {
            // Set array variable to store split version of formatted dates
            string[] arrayDate = RawDate.Split('/');
            // Creates variables for day, month and year
            Int32 year = Convert.ToInt32(arrayDate[2]);
            Int32 month = Convert.ToInt32(arrayDate[1]);
            Int32 day = Convert.ToInt32(arrayDate[0]);

            // Use DateTime function to find out what day of the week it is
            DateTime MyDate = new DateTime(year, month, day);
            // Set string for day of the week
            string DayOfWeek = MyDate.DayOfWeek.ToString();
            // Set Int32 variable to day of the year
            Int32 DayOfYear = MyDate.DayOfYear;
            // Set month name variable by converting my date into string and only using the month
            String MonthName = MyDate.ToString("MMM");
            // Set number of week in the year
            Int32 WeekNumber = DayOfYear / 7;
            // Default set weekend to false
            Boolean Weekend = false;
            // If day of the week is equal to saturday or sunday set weekend to true
            if (DayOfWeek == "Saturday" || DayOfWeek == "Sunday") { Weekend = true; }

            // Convert to database friendly format
            string dbDate = MyDate.ToString("M/dd/yyyy");

            // Call InsertTimeDimension function
            InsertTimeDimension(dbDate, DayOfWeek, DayOfYear, MonthName, month, WeekNumber, year, Weekend, DayOfYear);
        }
        private void SplitCustomers(string RawCustomer)
        {
            // set array variable to store split data of customers
            string[] arrayCustomer = RawCustomer.Split('/');

            // Customer variable
            // Name
            String Name = arrayCustomer[1];
            // Country
            string Country = arrayCustomer[2];
            //City
            string City = arrayCustomer[3];
            // State
            string State = arrayCustomer[4];
            // Postal Code
            string PostCode = arrayCustomer[5];
            // Region
            string Region = arrayCustomer[6];
            // Reference
            string Reference = arrayCustomer[0];

            // Call InsertCustomerDimension
            InsertCustomerDimension(Name, Country, City, State, PostCode, Region, Reference);
        }

        private void SplitProducts(string RawProduct)
        {
            // set array variable to store split data of customers
            string[] arrayProduct = RawProduct.Split('/');

            // Product variable

            // Product Name
            string ProductName = arrayProduct[3];
            // Category
            string Category = arrayProduct[1];
            // Sub Categoey
            string SubCategory = arrayProduct[2];
            // Reference 
            string Reference = arrayProduct[0];

            // Call InsertCustomerDimension
            InsertProductDimension(ProductName, Category, SubCategory, Reference);
        }

        private void InsertProductDimension(string productname, string category, string subcategory, string reference)
        {
            // Create connection to destination database string
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            using (SqlConnection myconnection = new SqlConnection(connectionStringDestination))
            {
                // Open connection
                myconnection.Open();
                // Check if the date already exists in the database
                SqlCommand command = new SqlCommand("SELECT id FROM Product where reference = @reference", myconnection);
                command.Parameters.Add(new SqlParameter("reference", reference));

                // Create a variable and assign it as false by defualt
                Boolean exists = false;

                // Run the command and read the results
                using (SqlDataReader reader = command.ExecuteReader())
                {


                    // If there are rules then the date already exist so set exist to true
                    if (reader.HasRows) { exists = true; }
                }

                if (exists == false)
                {
                    SqlCommand InsertCommand = new SqlCommand(
                        "INSERT INTO Product (category, subcategory, name, reference)" +
                        "VALUES (@category, @subcategory, @name, @reference)", myconnection);
                    InsertCommand.Parameters.Add(new SqlParameter("category", category));
                    InsertCommand.Parameters.Add(new SqlParameter("subcategory", subcategory));
                    InsertCommand.Parameters.Add(new SqlParameter("name", productname));
                    InsertCommand.Parameters.Add(new SqlParameter("reference", reference));

                    // Insert the line
                    int recordsAffected = InsertCommand.ExecuteNonQuery();
                }
            }
        }

        private void InsertCustomerDimension(string name, string country, string city, string state, string postcode, string region, string reference)
        {
            // Create connection to destination database string
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            using (SqlConnection myconnection = new SqlConnection(connectionStringDestination))
            {
                // Open connection
                myconnection.Open();
                // Check if the date already exists in the database
                SqlCommand command = new SqlCommand("SELECT id FROM Customer where reference = @reference", myconnection);
                command.Parameters.Add(new SqlParameter("reference", reference));

                // Create a variable and assign it as false by defualt
                Boolean exists = false;

                // Run the command and read the results
                using (SqlDataReader reader = command.ExecuteReader())
                {


                    // If there are rules then the date already exist so set exist to true
                    if (reader.HasRows) { exists = true; }
                }

                if (exists == false)
                {
                    SqlCommand InsertCommand = new SqlCommand(
                        " INSERT INTO Customer (name, country, city, state, postalcode, region, reference) " +
                        " VALUES (@name, @country, @city, @state, @postalcode, @region, @reference) ", myconnection);
                    InsertCommand.Parameters.Add(new SqlParameter("name", name));
                    InsertCommand.Parameters.Add(new SqlParameter("country", country));
                    InsertCommand.Parameters.Add(new SqlParameter("city", city));
                    InsertCommand.Parameters.Add(new SqlParameter("state", state));
                    InsertCommand.Parameters.Add(new SqlParameter("postalcode", postcode));
                    InsertCommand.Parameters.Add(new SqlParameter("region", region));
                    InsertCommand.Parameters.Add(new SqlParameter("reference", reference));

                    // Insert the line
                    int recordsAffected = InsertCommand.ExecuteNonQuery();
                }
            }
        }

        private void InsertTimeDimension(string date, string dayName, Int32 dayNumber, string monthName, Int32 monthNumber, Int32 weekNumber, Int32 year, Boolean weekend, Int32 dayofyear)
        {
            // Create connection to destination database string
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            using (SqlConnection myconnection = new SqlConnection(connectionStringDestination))
            {
                // Open connection
                myconnection.Open();
                // Check if the date already exists in the database
                SqlCommand command = new SqlCommand("SELECT id FROM Time where date = @date", myconnection);
                command.Parameters.Add(new SqlParameter("date", date));

                // Create a variable and assign it as false by defualt
                Boolean exists = false;

                // Run the command and read the results
                using (SqlDataReader reader = command.ExecuteReader())
                {


                    // If there are rules then the date already exist so set exist to true
                    if (reader.HasRows) { exists = true; }
                }

                if (exists == false)
                {
                    SqlCommand InsertCommand = new SqlCommand(
                        " INSERT INTO Time (dayName, dayNumber, monthName, monthNumber, weekNumber, year, weekend, date, dayofyear) " +
                        " VALUES (@dayName, @dayNumber, @monthName, @monthNumber, @weekNumber, @year, @weekend, @date, @dayofyear) ", myconnection);
                    InsertCommand.Parameters.Add(new SqlParameter("dayName", dayName));
                    InsertCommand.Parameters.Add(new SqlParameter("dayNumber", dayNumber));
                    InsertCommand.Parameters.Add(new SqlParameter("monthName", monthName));
                    InsertCommand.Parameters.Add(new SqlParameter("monthNumber", monthNumber));
                    InsertCommand.Parameters.Add(new SqlParameter("weekNumber", weekNumber));
                    InsertCommand.Parameters.Add(new SqlParameter("year", year));
                    InsertCommand.Parameters.Add(new SqlParameter("weekend", weekend));
                    InsertCommand.Parameters.Add(new SqlParameter("date", date));
                    InsertCommand.Parameters.Add(new SqlParameter("dayofyear", dayofyear));

                    // Insert the line
                    int recordsAffected = InsertCommand.ExecuteNonQuery();
                }
            }
        }

        private void InsertFactTable(Int32 productid, Int32 customerid, Int32 timeid, Int32 quantity, decimal sales, decimal profit, decimal discount)
        {
            // Create connection to destination database string
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            using (SqlConnection myconnection = new SqlConnection(connectionStringDestination))
            {
                // Open connection
                myconnection.Open();
                // Check if the date already exists in the database
                SqlCommand command = new SqlCommand("SELECT timeId FROM FactTable1 where timeId = @timeid", myconnection);
                command.Parameters.Add(new SqlParameter("timeid", timeid));

                // Create a variable and assign it as false by defualt
                Boolean exists = false;

                // Run the command and read the results
                using (SqlDataReader reader = command.ExecuteReader())
                {


                    // If there are rules then the date already exist so set exist to true
                    if (reader.HasRows) { exists = true; }
                }

                if (exists == false)
                {
                    SqlCommand InsertCommand = new SqlCommand(
                        " INSERT INTO FactTable1 (productId, timeId, customerId, value, discount, profit, quantity) " +
                        " VALUES (@productId, @timeId, @customerId, @value, @discount, @profit, @quantity) ", myconnection);
                    InsertCommand.Parameters.Add(new SqlParameter("productId", productid));
                    InsertCommand.Parameters.Add(new SqlParameter("timeId", timeid));
                    InsertCommand.Parameters.Add(new SqlParameter("customerId", customerid));
                    InsertCommand.Parameters.Add(new SqlParameter("value", sales));
                    InsertCommand.Parameters.Add(new SqlParameter("discount", discount));
                    InsertCommand.Parameters.Add(new SqlParameter("profit", profit));
                    InsertCommand.Parameters.Add(new SqlParameter("quantity", quantity));

                    // Insert the line
                    int recordsAffected = InsertCommand.ExecuteNonQuery();
                }
            }
        }

        private void btnGetDates_Click(object sender, EventArgs e)
        {
            // Creates Lists to store the dates in
            List<string> Dates = new List<string>();
            // Clear the listboxes to ensure no old data is in there
            listBoxDates.Items.Clear();

            // Create connection to data set string
            string connectionstring = Properties.Settings.Default.Data_set_1ConnectionString;

            using (OleDbConnection connection = new OleDbConnection(connectionstring))
            {
                // Open Connection
                connection.Open();
                // Set reader to null
                OleDbDataReader reader = null;
                // Perform sql query using connection variable
                OleDbCommand getDates = new OleDbCommand("SELECT [Order Date], [Ship Date] FROM Sheet1", connection);

                // Execute reader
                reader = getDates.ExecuteReader();
                // While the reader is reading
                while (reader.Read())
                {
                    // Add the dates found in the query by the reader into a string
                    Dates.Add(reader[0].ToString());
                    Dates.Add(reader[1].ToString());
                }
            }

            // Create list for formatted dates
            List<string> FormattedDates = new List<string>();

            // For each loop to change the date format for every date found in the query
            foreach (string date in Dates)
            {
                var dates = date.Split(new char[0], StringSplitOptions.RemoveEmptyEntries);
                FormattedDates.Add(dates[0]);
            }

            // Insert list box data into the dates list variable
            listBoxDates.DataSource = FormattedDates;

            // Loop through all data in database 
            foreach (string date in FormattedDates)
            {
                // Call function SplitDates
                SplitDates(date);
            }
        }

        private void btnGetCustomers_Click(object sender, EventArgs e)
        {
            // Creates Lists to store the dates in
            List<string> Customers = new List<string>();
            // Clear the listboxes to ensure no old data is in there
            listBoxCustomers.Items.Clear();

            // Create connection to data set string
            string connectionstring = Properties.Settings.Default.Data_set_1ConnectionString;

            using (OleDbConnection connection = new OleDbConnection(connectionstring))
            {
                // Open Connection
                connection.Open();
                // Set reader to null
                OleDbDataReader reader = null;
                // Perform sql query using connection variable
                OleDbCommand getCustomers = new OleDbCommand("SELECT [Customer ID], [Customer Name], Country, City, State, [Postal Code], Region FROM Sheet1", connection);

                // Execute reader
                reader = getCustomers.ExecuteReader();
                // While the reader is reading
                while (reader.Read())
                {
                    // Add the dates found in the query by the reader into a string
                    Customers.Add(reader[0].ToString() + "/" +
                        reader[1].ToString() + "/" +
                        reader[2].ToString() + "/" +
                        reader[3].ToString() + "/" +
                        reader[4].ToString() + "/" +
                        reader[5].ToString() + "/" +
                        reader[6].ToString());
                }
            }

            // Insert list box data into the dates list variable
            listBoxCustomers.DataSource = Customers;

            // Loop through all data in database 
            foreach (string name in Customers)
            {
                // Call function SplitCustomers
                SplitCustomers(name);
            }
        }

        private void btnGetProducts_Click(object sender, EventArgs e)
        {
            // Creates Lists to store the dates in
            List<string> Products = new List<string>();
            // Clear the listboxes to ensure no old data is in there
            listBoxProducts.Items.Clear();

            // Create connection to data set string
            string connectionstring = Properties.Settings.Default.Data_set_1ConnectionString;

            using (OleDbConnection connection = new OleDbConnection(connectionstring))
            {
                // Open Connection
                connection.Open();
                // Set reader to null
                OleDbDataReader reader = null;
                // Perform sql query using connection variable
                OleDbCommand getProducts = new OleDbCommand("SELECT [Product ID], Category, [Sub-Category], [Product Name] FROM Sheet1", connection);

                // Execute reader
                reader = getProducts.ExecuteReader();
                // While the reader is reading
                while (reader.Read())
                {
                    // Add the dates found in the query by the reader into a string
                    Products.Add(reader[0].ToString() + "/" +
                        reader[1].ToString() + "/" +
                        reader[2].ToString() + "/" +
                        reader[3].ToString());
                }
            }

            // Insert list box data into the dates list variable
            listBoxProducts.DataSource = Products;

            // Loop through all data in database 
            foreach (string name in Products)
            {
                // Call function SplitProducts
                SplitProducts(name);
            }

        }

        private void btnDatesDestination_Click(object sender, EventArgs e)
        {
            // Create List to store the data in
            List<String> DestinationDate = new List<string>();

            // Clear the listbox to ensure no old data is in there
            listBoxDatesDestination.DataSource = null;
            listBoxDatesDestination.Items.Clear();

            // Create connection to the MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // Open the connnection
                myConnection.Open();
                // Check if the date already exists in the database
                SqlCommand command = new SqlCommand("SELECT id, dayName, dayNumber, monthName, monthNumber, weekNumber, year, weekend, date, dayofyear FROM Time", myConnection);



                using (SqlDataReader reader = command.ExecuteReader())
                {
                    // Check if there is data in the table
                    if (reader.HasRows)
                    {
                        // Retrieve the data
                        while (reader.Read())
                        {
                            string id = reader["id"].ToString();
                            string dayName = reader["dayName"].ToString();
                            string dayNumber = reader["dayNumber"].ToString();
                            string monthName = reader["monthName"].ToString();
                            string monthNumber = reader["monthNumber"].ToString();
                            string weekNumber = reader["weekNumber"].ToString();
                            string year = reader["year"].ToString();
                            string weekend = reader["weekend"].ToString();
                            string date = reader["date"].ToString();
                            string dayOfYear = reader["dayofyear"].ToString();

                            string text;

                            text = "ID: " + id + " Day Name: " + dayName + " Day Number: " + dayNumber + " Month Name: " + monthName + " Month Number: " + monthNumber +
                                " Week Number: " + weekNumber + " Year: " + year + " Weekend: " + weekend + "Date: " + date + " Day of the Year: " + dayOfYear;

                            DestinationDate.Add(text);
                        }
                    }
                    else // If there is no data available
                    {
                        DestinationDate.Add("No Data available in the time dimension");
                    }
                }
            }

            // Bind the listbox to the list
            listBoxDatesDestination.DataSource = DestinationDate;

        }

        private void btnCustomerDestination_Click(object sender, EventArgs e)
        {
            // Create List to store the data in
            List<String> DestinationCustomer = new List<string>();

            // Clear the listbox to ensure no old data is in there
            listBoxCustomersDestination.DataSource = null;
            listBoxCustomersDestination.Items.Clear();

            // Create connection to the MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // Open the connnection
                myConnection.Open();
                // Check if the date already exists in the database
                SqlCommand command = new SqlCommand("SELECT id, name, country, city, state, postalCode, region, reference FROM Customer", myConnection);



                using (SqlDataReader reader = command.ExecuteReader())
                {
                    // Check if there is data in the table
                    if (reader.HasRows)
                    {
                        // Retrieve the data
                        while (reader.Read())
                        {
                            string id = reader["id"].ToString();
                            string customerName = reader["name"].ToString();
                            string country = reader["country"].ToString();
                            string city = reader["city"].ToString();
                            string state = reader["state"].ToString();
                            string postCode = reader["postalCode"].ToString();
                            string region = reader["region"].ToString();
                            string reference = reader["reference"].ToString();

                            string text;

                            text = "ID: " + id + " Name: " + customerName + " Country: " + country + " City: " + city + " State: " + state +
                                " Postal Code: " + postCode + " Region: " + region + " Reference: " + reference;

                            DestinationCustomer.Add(text);
                        }
                    }
                    else // If there is no data available
                    {
                        DestinationCustomer.Add("No Data available in the time dimension");
                    }
                }
            }

            // Bind the listbox to the list
            listBoxCustomersDestination.DataSource = DestinationCustomer;
        }

        private void btnProductsDestination_Click(object sender, EventArgs e)
        {
            // Create List to store the data in
            List<String> DestinationProduct = new List<string>();

            // Clear the listbox to ensure no old data is in there
            listBoxProductsDestination.DataSource = null;
            listBoxProductsDestination.Items.Clear();

            // Create connection to the MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // Open the connnection
                myConnection.Open();
                // Check if the date already exists in the database
                SqlCommand command = new SqlCommand("SELECT id, category, subcategory, name, reference FROM Product", myConnection);



                using (SqlDataReader reader = command.ExecuteReader())
                {
                    // Check if there is data in the table
                    if (reader.HasRows)
                    {
                        // Retrieve the data
                        while (reader.Read())
                        {
                            string id = reader["id"].ToString();
                            string productname = reader["name"].ToString();
                            string category = reader["category"].ToString();
                            string subcategory = reader["subcategory"].ToString();
                            string reference = reader["reference"].ToString();

                            string text;

                            text = "ID: " + id + " Name: " + productname + " Category: " + category + " Sub-Category: " + subcategory + " Reference: " + reference;

                            DestinationProduct.Add(text);
                        }
                    }
                    else // If there is no data available
                    {
                        DestinationProduct.Add("No Data available in the time dimension");
                    }
                }
            }

            // Bind the listbox to the list
            listBoxProductsDestination.DataSource = DestinationProduct;
        }

        private void btnFactTable_Click(object sender, EventArgs e)
        {
            // Create connection to data set string
            string connectionstring = Properties.Settings.Default.Data_set_1ConnectionString;

            using (OleDbConnection connection = new OleDbConnection(connectionstring))
            {
                // Open Connection
                connection.Open();
                // Set reader to null
                OleDbDataReader reader = null;
                // Perform sql query using connection variable
                OleDbCommand getAll = new OleDbCommand("SELECT ID, [Row ID], [Order ID], [Order Date], [Ship Date], [Ship Mode], [Customer ID], [Customer Name], Segment, Country, City, State, [Postal Code], Region, [Product ID], Category, [Sub-Category], [Product Name], Sales, Quantity," +
                         "Discount, Profit FROM Sheet1", connection);

                // Execute reader
                reader = getAll.ExecuteReader();
                // While the reader is reading
                while (reader.Read())
                {
                    // Get a line of source data 

                    // Get the numerical values
                    Decimal sales = Convert.ToDecimal(reader["sales"]);
                    Int32 quantity = Convert.ToInt32(reader["quantity"]);
                    Decimal profit = Convert.ToDecimal(reader["profit"]);
                    Decimal discount = Convert.ToDecimal(reader["discount"]);
                    // Get the dimension ID's
                    Int32 timeId = GetDateID(reader["Order Date"].ToString());
                    Int32 productId = GetProductID(reader["Product ID"].ToString());
                    Int32 customerId = GetCustomerID(reader["Customer Name"].ToString());

                    // Insert it into database
                    InsertFactTable(productId, customerId, timeId, quantity, sales, profit, discount);


                }

            }
            // Build table
            // Create List to store the data in
            List<String> FactTableList = new List<string>();

            // Clear the listbox to ensure no old data is in there
            listBoxFactTable.DataSource = null;
            listBoxFactTable.Items.Clear();

            // Create connection to the MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                // Open the connnection
                myConnection.Open();
                // Check if the date already exists in the database
                SqlCommand command = new SqlCommand("SELECT productId, timeId, customerId, value, discount, profit, quantity FROM FactTable1", myConnection);



                using (SqlDataReader reader = command.ExecuteReader())
                {
                    // Check if there is data in the table
                    if (reader.HasRows)
                    {
                        // Retrieve the data
                        while (reader.Read())
                        {
                            string productid = reader["productId"].ToString();
                            string timeid = reader["timeId"].ToString();
                            string customerid = reader["customerId"].ToString();
                            string value = reader["value"].ToString();
                            string discount = reader["discount"].ToString();
                            string profit = reader["profit"].ToString();
                            string quantity = reader["quantity"].ToString();

                            string text;

                            text = "Product ID: " + productid + " Customer ID: " + customerid + " Time ID: " + timeid + " Value: " + value + " Discount: " + discount + " Proft: " + profit + " Quantity: " + quantity; ;

                            FactTableList.Add(text);
                        }
                    }
                    else // If there is no data available
                    {
                        FactTableList.Add("No Data available in the time dimension");
                    }
                }
            }

            // Bind the listbox to the list
            listBoxFactTable.DataSource = FactTableList;

            // Combo boxes
            // Key Combo Boxes
            // Box 1
            comboBoxKey1.Items.Add("Product Name");
            comboBoxKey1.Items.Add("Category");
            comboBoxKey1.Items.Add("Sub-Category");
            // Box 2
            comboBoxKey2.Items.Add("Product Name");
            comboBoxKey2.Items.Add("Category");
            comboBoxKey2.Items.Add("Sub-Category");

            // Value Combo Boxes
            // Box 1
            comboBoxValue1.Items.Add("Profit");
            comboBoxValue1.Items.Add("Value");
            comboBoxValue1.Items.Add("Quantity");
            comboBoxValue1.Items.Add("Discount");
            // Box 2
            comboBoxValue2.Items.Add("Profit");
            comboBoxValue2.Items.Add("Value");
            comboBoxValue2.Items.Add("Quantity");
            comboBoxValue2.Items.Add("Discount");
        }

        private void btnLoadData_Click(object sender, EventArgs e)
        {
            // Create strings for dates
            List<string> Dates = new List<string>();

            // Create a connection to MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                //Open connection
                myConnection.Open();
                // Get dates from fact table
                SqlCommand command = new SqlCommand("SELECT date From FactTable1 JOIN Time ON FactTable1.timeId = Time.id", myConnection);

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    // Check if there is data in the table
                    if (reader.HasRows)
                    {
                        // Retrieve the data
                        while (reader.Read())
                        {
                            // Add Date to date list
                            Dates.Add(reader["date"].ToString());
                        }
                    }
                }
            }
            // Create list for newly formatted dates
            List<string> dateList = new List<string>();
            foreach (string date in Dates)
            {
                var dates = date.Split(new Char[0], StringSplitOptions.RemoveEmptyEntries);
                dateList.Add(dates[0]);
            }

            // Create Dictionaries
            // Create Sales Count
            Dictionary<String, Int32> SalesCount = new Dictionary<string, Int32>();
            // Create profit
            Dictionary<String, Decimal> Profit = new Dictionary<string, Decimal>();
            // Create Quantity
            Dictionary<String, Int32> Quantity = new Dictionary<string, Int32>();
            // Create Value
            Dictionary<String, Decimal> Value = new Dictionary<string, Decimal>();

            // Run through a look 7 times
            foreach (string date in dateList)
            {
                String[] arrayDate = date.Split('/');
                // Creates variables for day, month and year
                Int32 year = Convert.ToInt32(arrayDate[2]);
                Int32 month = Convert.ToInt32(arrayDate[1]);
                Int32 day = Convert.ToInt32(arrayDate[0]);

                DateTime myDate = new DateTime(year, month, day);

                // Convert to database friendly format
                string dbDate = myDate.ToString("M/dd/yyyy");

                using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
                {
                    // Open the connnection
                    myConnection.Open();
                    // Check if the date already exists in the database
                    SqlCommand command = new SqlCommand("SELECT COUNT(*) as SalesNumber From FactTable1 JOIN Time ON FactTable1.timeId = Time.id WHERE date = @date", myConnection);
                    command.Parameters.Add(new SqlParameter("@date", dbDate));



                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        // Check if there is data in the table
                        if (reader.HasRows)
                        {
                            // Retrieve the data
                            while (reader.Read())
                            {
                                SalesCount.Add(date, Int32.Parse(reader["SalesNumber"].ToString()));
                            }
                        }
                        else // If there is no data available
                        {
                            SalesCount.Add(date, 0);
                        }
                    }
                }

                using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
                {
                    // Open the connnection
                    myConnection.Open();
                    // Check if the date already exists in the database
                    SqlCommand command = new SqlCommand("SELECT profit From FactTable1 JOIN Time ON FactTable1.timeId = Time.id WHERE date = @date", myConnection);
                    command.Parameters.Add(new SqlParameter("@date", dbDate));



                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        // Check if there is data in the table
                        if (reader.HasRows)
                        {
                            // Retrieve the data
                            while (reader.Read())
                            {
                                Profit.Add(date, Decimal.Parse(reader["profit"].ToString()));
                            }
                        }
                        else // If there is no data available
                        {
                            Profit.Add(date, 0);
                        }
                    }
                }

                using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
                {
                    // Open the connnection
                    myConnection.Open();
                    // Check if the date already exists in the database
                    SqlCommand command = new SqlCommand("SELECT quantity From FactTable1 JOIN Time ON FactTable1.timeId = Time.id WHERE date = @date", myConnection);
                    command.Parameters.Add(new SqlParameter("@date", dbDate));



                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        // Check if there is data in the table
                        if (reader.HasRows)
                        {
                            // Retrieve the data
                            while (reader.Read())
                            {
                                Quantity.Add(date, Int32.Parse(reader["quantity"].ToString()));
                            }
                        }
                        else // If there is no data available
                        {
                            Quantity.Add(date, 0);
                        }
                    }
                }

                using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
                {
                    // Open the connnection
                    myConnection.Open();
                    // Check if the date already exists in the database
                    SqlCommand command = new SqlCommand("SELECT value From FactTable1 JOIN Time ON FactTable1.timeId = Time.id WHERE date = @date", myConnection);
                    command.Parameters.Add(new SqlParameter("@date", dbDate));



                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        // Check if there is data in the table
                        if (reader.HasRows)
                        {
                            // Retrieve the data
                            while (reader.Read())
                            {
                                Value.Add(date, Decimal.Parse(reader["value"].ToString()));
                            }
                        }
                        else // If there is no data available
                        {
                            Value.Add(date, 0);
                        }
                    }
                }

            }
            // End of for loop, array should be filled
            // Build Bar Chart
            chart1.DataSource = SalesCount;
            chart1.Series[0].XValueMember = "Key";
            chart1.Series[0].YValueMembers = "Value";
            chart1.DataBind();

            // Build Spiline Chart
            chart2.DataSource = Profit;
            chart2.Series[0].XValueMember = "Key";
            chart2.Series[0].YValueMembers = "Value";
            chart2.DataBind();

            // Build Stacked Chart
            chart3.DataSource = Quantity;
            chart3.Series[0].XValueMember = "Key";
            chart3.Series[0].YValueMembers = "Value";
            chart3.DataBind();

            // Build Line Chart
            chart4.DataSource = Value;
            chart4.Series[0].XValueMember = "Key";
            chart4.Series[0].YValueMembers = "Value";
            chart4.DataBind();
        }

        private void btnLoadData2_Click(object sender, EventArgs e)
        {
            // Create strings for Customer
            List<string> Customers = new List<string>();
            // Create list for states
            List<string> States = new List<string>();
            // Create list for postcodes
            List<string> PostCode = new List<string>();
            //Create list for Cities
            List<string> Cities = new List<string>();

            // Create a connection to MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                //Open connection
                myConnection.Open();
                // Get dates from fact table
                SqlCommand command = new SqlCommand("SELECT DISTINCT postalcode, name, state, city From FactTable1 JOIN Customer ON FactTable1.customerId = Customer.id", myConnection);

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    // Check if there is data in the table
                    if (reader.HasRows)
                    {
                        // Retrieve the data
                        while (reader.Read())
                        {
                            // Add region to customer list
                            Customers.Add(reader["name"].ToString());
                            PostCode.Add(reader["postalcode"].ToString());
                            States.Add(reader["state"].ToString());
                            Cities.Add(reader["city"].ToString());
                        }
                    }
                }
            }

            // Create Dictionaries
            // Create profit
            Dictionary<String, Decimal> ProfitCustomer = new Dictionary<string, Decimal>();
            // Create Quantity
            Dictionary<String, Int32> QuantityCustomer = new Dictionary<string, Int32>();
            // Create Value
            Dictionary<String, Decimal> ValueCustomer = new Dictionary<string, Decimal>();
            // Create Discount
            Dictionary<String, Decimal> DiscountCustomer = new Dictionary<string, Decimal>();

            // Loop through names
            foreach (string name in Customers)
            {
                using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
                {
                    // Open the connnection
                    myConnection.Open();
                    // Check if the quantity already exists in the database
                    SqlCommand command = new SqlCommand("SELECT quantity From FactTable1 JOIN Customer ON FactTable1.customerId = Customer.id WHERE name = @name", myConnection);
                    command.Parameters.Add(new SqlParameter("@name", name));

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        // Check if there is data in the table
                        if (reader.HasRows)
                        {
                            // Retrieve the data
                            while (reader.Read())
                            {
                                QuantityCustomer.Add(name, Int32.Parse(reader["quantity"].ToString()));
                            }
                        }
                        else // If there is no data available
                        {
                            QuantityCustomer.Add(name, 0);
                        }
                    }
                }
            }

            foreach (string state in States)
            {
                using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
                {
                    // Open the connnection
                    myConnection.Open();
                    // Check if the date already exists in the database
                    SqlCommand command = new SqlCommand("SELECT value From FactTable1 JOIN Customer ON FactTable1.customerId = Customer.id WHERE state = @state", myConnection);
                    command.Parameters.Add(new SqlParameter("@state", state));



                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        // Check if there is data in the table
                        if (reader.HasRows)
                        {
                            // Retrieve the data
                            while (reader.Read())
                            {
                                if (!ValueCustomer.ContainsKey(state))
                                {
                                    ValueCustomer.Add(state, Decimal.Parse(reader["value"].ToString()));
                                }
                            }
                        }
                        else // If there is no data available
                        {
                            ValueCustomer.Add(state, 0);
                        }
                    }
                }

                foreach (string postalcode in PostCode)
                {
                    using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
                    {
                        // Open the connnection
                        myConnection.Open();
                        // Check if the date already exists in the database
                        SqlCommand command = new SqlCommand("SELECT profit From FactTable1 JOIN Customer ON FactTable1.customerId = Customer.id WHERE postalcode = @postalcode", myConnection);
                        command.Parameters.Add(new SqlParameter("@postalcode", postalcode));



                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            // Check if there is data in the table
                            if (reader.HasRows)
                            {
                                // Retrieve the data
                                while (reader.Read())
                                {
                                    if (!ProfitCustomer.ContainsKey(postalcode))
                                    {
                                        ProfitCustomer.Add(postalcode, Decimal.Parse(reader["profit"].ToString()));
                                    }
                                }
                            }
                            else // If there is no data available
                            {
                                ProfitCustomer.Add(postalcode, 0);
                            }
                        }
                    }
                }

                foreach (string city in Cities)
                {
                    using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
                    {
                        // Open the connnection
                        myConnection.Open();
                        // Check if the date already exists in the database
                        SqlCommand command = new SqlCommand("SELECT discount From FactTable1 JOIN Customer ON FactTable1.customerId = Customer.id WHERE city = @city", myConnection);
                        command.Parameters.Add(new SqlParameter("@city", city));



                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            // Check if there is data in the table
                            if (reader.HasRows)
                            {
                                // Retrieve the data
                                while (reader.Read())
                                {
                                    if (!DiscountCustomer.ContainsKey(city))
                                    {
                                        DiscountCustomer.Add(city, Decimal.Parse(reader["discount"].ToString()));
                                    }
                                }
                            }
                            else // If there is no data available
                            {
                                DiscountCustomer.Add(city, 0);
                            }
                        }
                    }
                }

                // End of for loop, array should be filled
                // Build Bar Chart
                chart8.DataSource = QuantityCustomer;
                chart8.Series[0].XValueMember = "Key";
                chart8.Series[0].YValueMembers = "Value";
                chart8.DataBind();

                // Build Spiline Chart
                chart7.DataSource = ProfitCustomer;
                chart7.Series[0].XValueMember = "Key";
                chart7.Series[0].YValueMembers = "Value";
                chart7.DataBind();

                // Build Stacked Chart
                chart6.DataSource = ValueCustomer;
                chart6.Series[0].XValueMember = "Key";
                chart6.Series[0].YValueMembers = "Value";
                chart6.DataBind();

                // Build Line Chart
                chart5.DataSource = DiscountCustomer;
                chart5.Series[0].XValueMember = "Key";
                chart5.Series[0].YValueMembers = "Value";
                chart5.DataBind();
            }
        }

        private void btnLoadData3_Click(object sender, EventArgs e)
        {
            // Create strings for Customer
            List<string> Name = new List<string>();
            // Create list for states
            List<string> Category = new List<string>();
            // Create list for postcodes
            List<string> SubCategory = new List<string>();

            // Create a connection to MDF file
            string connectionStringDestination = Properties.Settings.Default.DestinationDatabaseConnectionString;

            using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
            {
                //Open connection
                myConnection.Open();
                // Get dates from fact table
                SqlCommand command = new SqlCommand("SELECT DISTINCT name, category, subcategory From FactTable1 JOIN Product ON FactTable1.productId = Product.id", myConnection);

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    // Check if there is data in the table
                    if (reader.HasRows)
                    {
                        // Retrieve the data
                        while (reader.Read())
                        {
                            // Add region to customer list
                            Name.Add(reader["name"].ToString());
                            Category.Add(reader["category"].ToString());
                            SubCategory.Add(reader["subcategory"].ToString());
                        }
                    }
                }
            }

            // Create Dictionaries
            // Create profit
            Dictionary<String, Decimal> ProfitProduct = new Dictionary<string, Decimal>();
            // Create Quantity
            Dictionary<String, Int32> QuantityProduct = new Dictionary<string, Int32>();
            // Create Value
            Dictionary<String, Decimal> ValueProduct = new Dictionary<string, Decimal>();
            // Create Discount
            Dictionary<String, Decimal> DiscountProduct = new Dictionary<string, Decimal>();

            if (comboBoxKey1.Text == "Product Name")
            {
                // Loop through names
                foreach (string name in Name)
                {
                    using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
                    {
                        // Open the connnection
                        myConnection.Open();
                        // Check if the quantity already exists in the database
                        if (comboBoxValue1.Text == "Profit")
                        {
                            SqlCommand command = new SqlCommand("SELECT profit From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE name = @name", myConnection);
                            command.Parameters.Add(new SqlParameter("@name", name));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        ProfitProduct.Add(name, Decimal.Parse(reader["profit"].ToString()));
                                    }
                                }
                                else // If there is no data available
                                {
                                    ProfitProduct.Add(name, 0);
                                }
                            }
                        }
                        else if (comboBoxValue1.Text == "Discount")
                        {
                            SqlCommand command = new SqlCommand("SELECT discount From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE name = @name", myConnection);
                            command.Parameters.Add(new SqlParameter("@name", name));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        DiscountProduct.Add(name, decimal.Parse(reader["discount"].ToString()));
                                    }
                                }
                                else // If there is no data available
                                {
                                    DiscountProduct.Add(name, 0);
                                }
                            }
                        }
                        else if (comboBoxValue1.Text == "Value")
                        {
                            SqlCommand command = new SqlCommand("SELECT value From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE name = @name", myConnection);
                            command.Parameters.Add(new SqlParameter("@name", name));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        ValueProduct.Add(name, Decimal.Parse(reader["value"].ToString()));
                                    }
                                }
                                else // If there is no data available
                                {
                                    ValueProduct.Add(name, 0);
                                }
                            }
                        }
                        else if (comboBoxValue1.Text == "Quantity")
                        {
                            SqlCommand command = new SqlCommand("SELECT quantity From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE name = @name", myConnection);
                            command.Parameters.Add(new SqlParameter("@name", name));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        QuantityProduct.Add(name, Int32.Parse(reader["quantity"].ToString()));
                                    }
                                }
                                else // If there is no data available
                                {
                                    QuantityProduct.Add(name, 0);
                                }
                            }
                        }
                    }
                }
            }

            if (comboBoxKey1.Text == "Category")
            {
                // Loop through names
                foreach (string category in Category)
                {
                    using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
                    {
                        // Open the connnection
                        myConnection.Open();
                        // Check if the quantity already exists in the database
                        if (comboBoxValue1.Text == "Profit")
                        {
                            SqlCommand command = new SqlCommand("SELECT profit From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE category = @category", myConnection);
                            command.Parameters.Add(new SqlParameter("@category", category));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        if (!ProfitProduct.ContainsKey(category))
                                        {
                                            ProfitProduct.Add(category, Decimal.Parse(reader["profit"].ToString()));
                                        }
                                    }
                                }
                                else // If there is no data available
                                {
                                    ProfitProduct.Add(category, 0);
                                }
                            }
                        }
                        else if (comboBoxValue1.Text == "Discount")
                        {
                            SqlCommand command = new SqlCommand("SELECT discount From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE category = @category", myConnection);
                            command.Parameters.Add(new SqlParameter("@category", category));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        if (!DiscountProduct.ContainsKey(category))
                                        {
                                            DiscountProduct.Add(category, Decimal.Parse(reader["discount"].ToString()));
                                        }
                                    }
                                }
                                else // If there is no data available
                                {
                                    DiscountProduct.Add(category, 0);
                                }
                            }
                        }
                        else if (comboBoxValue1.Text == "Value")
                        {
                            SqlCommand command = new SqlCommand("SELECT value From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE category = @category", myConnection);
                            command.Parameters.Add(new SqlParameter("@category", category));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        if (!ValueProduct.ContainsKey(category))
                                        {
                                            ValueProduct.Add(category, Decimal.Parse(reader["value"].ToString()));
                                        }
                                    }
                                }
                                else // If there is no data available
                                {
                                    ValueProduct.Add(category, 0);
                                }
                            }
                        }
                        else if (comboBoxValue1.Text == "Quantity")
                        {
                            SqlCommand command = new SqlCommand("SELECT quantity From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE category = @category", myConnection);
                            command.Parameters.Add(new SqlParameter("@category", category));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        if (!QuantityProduct.ContainsKey(category))
                                        {
                                            QuantityProduct.Add(category, Int32.Parse(reader["Quantity"].ToString()));
                                        }
                                    }
                                }
                                else // If there is no data available
                                {
                                    QuantityProduct.Add(category, 0);
                                }
                            }
                        }
                    }
                }
            }
            if (comboBoxKey1.Text == "Sub-Category")
            {
                // Loop through names
                foreach (string subcategory in SubCategory)
                {
                    using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
                    {
                        // Open the connnection
                        myConnection.Open();
                        // Check if the quantity already exists in the database
                        if (comboBoxValue1.Text == "Profit")
                        {
                            SqlCommand command = new SqlCommand("SELECT profit From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE subcategory = @subcategory", myConnection);
                            command.Parameters.Add(new SqlParameter("@subcategory", subcategory));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        if (!ProfitProduct.ContainsKey(subcategory))
                                        {
                                            ProfitProduct.Add(subcategory, Decimal.Parse(reader["profit"].ToString()));
                                        }
                                    }
                                }
                                else // If there is no data available
                                {
                                    ProfitProduct.Add(subcategory, 0);
                                }
                            }
                        }
                        else if (comboBoxValue1.Text == "Discount")
                        {
                            SqlCommand command = new SqlCommand("SELECT discount From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE subcategory = @subcategory", myConnection);
                            command.Parameters.Add(new SqlParameter("@subcategory", subcategory));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        if(!DiscountProduct.ContainsKey(subcategory))
                                        {
                                            DiscountProduct.Add(subcategory, Decimal.Parse(reader["discount"].ToString()));
                                        }
                                    }
                                }
                                else // If there is no data available
                                {
                                    DiscountProduct.Add(subcategory, 0);
                                }
                            }
                        }
                        else if (comboBoxValue1.Text == "Value")
                        {
                            SqlCommand command = new SqlCommand("SELECT value From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE subcategory = @subcategory", myConnection);
                            command.Parameters.Add(new SqlParameter("@subcategory", subcategory));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        if (!ValueProduct.ContainsKey(subcategory))
                                        {
                                            ValueProduct.Add(subcategory, Decimal.Parse(reader["value"].ToString()));
                                        }
                                    }
                                }
                                else // If there is no data available
                                {
                                    ValueProduct.Add(subcategory, 0);
                                }
                            }
                        }
                        else if (comboBoxValue1.Text == "Quantity")
                        {
                            SqlCommand command = new SqlCommand("SELECT quantity From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE subcategory = @subcategory", myConnection);
                            command.Parameters.Add(new SqlParameter("@subcategory", subcategory));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        if (!QuantityProduct.ContainsKey(subcategory))
                                        {
                                            QuantityProduct.Add(subcategory, Int32.Parse(reader["quantity"].ToString()));
                                        }
                                    }
                                }
                                else // If there is no data available
                                {
                                    QuantityProduct.Add(subcategory, 0);
                                }
                            }
                        }
                    }
                }
            }


            if (comboBoxKey2.Text == "Product Name")
            {
                // Loop through names
                foreach (string name in Name)
                {
                    using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
                    {
                        // Open the connnection
                        myConnection.Open();
                        // Check if the quantity already exists in the database
                        if (comboBoxValue2.Text == "Profit")
                        {
                            SqlCommand command = new SqlCommand("SELECT profit From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE name = @name", myConnection);
                            command.Parameters.Add(new SqlParameter("@name", name));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        if (!ProfitProduct.ContainsKey(name))
                                        {
                                            ProfitProduct.Add(name, Decimal.Parse(reader["profit"].ToString()));
                                        }
                                    }
                                }
                                else // If there is no data available
                                {
                                    ProfitProduct.Add(name, 0);
                                }
                            }
                        }
                        else if (comboBoxValue2.Text == "Discount")
                        {
                            SqlCommand command = new SqlCommand("SELECT discount From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE name = @name", myConnection);
                            command.Parameters.Add(new SqlParameter("@name", name));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        if (!DiscountProduct.ContainsKey(name))
                                        {
                                            DiscountProduct.Add(name, Decimal.Parse(reader["discount"].ToString()));
                                        }
                                    }
                                }
                                else // If there is no data available
                                {
                                    DiscountProduct.Add(name, 0);
                                }
                            }
                        }
                        else if (comboBoxValue2.Text == "Value")
                        {
                            SqlCommand command = new SqlCommand("SELECT value From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE name = @name", myConnection);
                            command.Parameters.Add(new SqlParameter("@name", name));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        if (!ValueProduct.ContainsKey(name))
                                        {
                                            ValueProduct.Add(name, Decimal.Parse(reader["value"].ToString()));
                                        }
                                    }
                                }
                                else // If there is no data available
                                {
                                    ValueProduct.Add(name, 0);
                                }
                            }
                        }
                        else if (comboBoxValue2.Text == "Quantity")
                        {
                            SqlCommand command = new SqlCommand("SELECT quantity From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE name = @name", myConnection);
                            command.Parameters.Add(new SqlParameter("@name", name));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        if (!QuantityProduct.ContainsKey(name))
                                        {
                                            QuantityProduct.Add(name, Int32.Parse(reader["quantity"].ToString()));
                                        }
                                    }
                                }
                                else // If there is no data available
                                {
                                    QuantityProduct.Add(name, 0);
                                }
                            }
                        }
                    }
                }
            }

            if (comboBoxKey2.Text == "Category")
            {
                // Loop through names
                foreach (string category in Category)
                {
                    using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
                    {
                        // Open the connnection
                        myConnection.Open();
                        // Check if the quantity already exists in the database
                        if (comboBoxValue2.Text == "Profit")
                        {
                            SqlCommand command = new SqlCommand("SELECT profit From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE category = @category", myConnection);
                            command.Parameters.Add(new SqlParameter("@category", category));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        if (!ProfitProduct.ContainsKey(category))
                                        {
                                            ProfitProduct.Add(category, Decimal.Parse(reader["profit"].ToString()));
                                        }
                                    }
                                }
                                else // If there is no data available
                                {
                                    ProfitProduct.Add(category, 0);
                                }
                            }
                        }
                        else if (comboBoxValue2.Text == "Discount")
                        {
                            SqlCommand command = new SqlCommand("SELECT discount From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE category = @category", myConnection);
                            command.Parameters.Add(new SqlParameter("@category", category));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        if (!DiscountProduct.ContainsKey(category))
                                        {
                                            DiscountProduct.Add(category, Decimal.Parse(reader["discount"].ToString()));
                                        }
                                    }
                                }
                                else // If there is no data available
                                {
                                    DiscountProduct.Add(category, 0);
                                }
                            }
                        }
                        else if (comboBoxValue2.Text == "Value")
                        {
                            SqlCommand command = new SqlCommand("SELECT value From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE category = @category", myConnection);
                            command.Parameters.Add(new SqlParameter("@category", category));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        if (!ValueProduct.ContainsKey(category))
                                        {
                                            ValueProduct.Add(category, Decimal.Parse(reader["value"].ToString()));
                                        }
                                    }
                                }
                                else // If there is no data available
                                {
                                    ValueProduct.Add(category, 0);
                                }
                            }
                        }
                        else if (comboBoxValue2.Text == "Quantity")
                        {
                            SqlCommand command = new SqlCommand("SELECT quantity From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE category = @category", myConnection);
                            command.Parameters.Add(new SqlParameter("@category", category));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        if (!QuantityProduct.ContainsKey(category))
                                        {
                                            QuantityProduct.Add(category, Int32.Parse(reader["quantity"].ToString()));
                                        }
                                    }
                                }
                                else // If there is no data available
                                {
                                    QuantityProduct.Add(category, 0);
                                }
                            }
                        }
                    }
                }
            }
            if (comboBoxKey2.Text == "Sub-Category")
            {
                // Loop through names
                foreach (string subcategory in SubCategory)
                {
                    using (SqlConnection myConnection = new SqlConnection(connectionStringDestination))
                    {
                        // Open the connnection
                        myConnection.Open();
                        // Check if the quantity already exists in the database
                        if (comboBoxValue2.Text == "Profit")
                        {
                            SqlCommand command = new SqlCommand("SELECT profit From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE subcategory = @subcategory", myConnection);
                            command.Parameters.Add(new SqlParameter("@subcategory", subcategory));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        if (!ProfitProduct.ContainsKey(subcategory))
                                        {
                                            ProfitProduct.Add(subcategory, Decimal.Parse(reader["profit"].ToString()));
                                        }
                                    }
                                }
                                else // If there is no data available
                                {
                                    ProfitProduct.Add(subcategory, 0);
                                }
                            }
                        }
                        else if (comboBoxValue2.Text == "Discount")
                        {
                            SqlCommand command = new SqlCommand("SELECT discount From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE subcategory = @subcategory", myConnection);
                            command.Parameters.Add(new SqlParameter("@subcategory", subcategory));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        if (!DiscountProduct.ContainsKey(subcategory))
                                        {
                                            DiscountProduct.Add(subcategory, Decimal.Parse(reader["discount"].ToString()));
                                        }
                                    }
                                }
                                else // If there is no data available
                                {
                                    DiscountProduct.Add(subcategory, 0);
                                }
                            }
                        }
                        else if (comboBoxValue2.Text == "Value")
                        {
                            SqlCommand command = new SqlCommand("SELECT value From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE subcategory = @subcategory", myConnection);
                            command.Parameters.Add(new SqlParameter("@subcategory", subcategory));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        if (!ValueProduct.ContainsKey(subcategory))
                                        {
                                            ValueProduct.Add(subcategory, Decimal.Parse(reader["value"].ToString()));
                                        }
                                    }
                                }
                                else // If there is no data available
                                {
                                    ValueProduct.Add(subcategory, 0);
                                }
                            }
                        }
                        else if (comboBoxValue2.Text == "Quantity")
                        {
                            SqlCommand command = new SqlCommand("SELECT quantity From FactTable1 JOIN Product ON FactTable1.productId = Product.id WHERE subcategory = @subcategory", myConnection);
                            command.Parameters.Add(new SqlParameter("@subcategory", subcategory));

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                // Check if there is data in the table
                                if (reader.HasRows)
                                {
                                    // Retrieve the data
                                    while (reader.Read())
                                    {
                                        if (!QuantityProduct.ContainsKey(subcategory))
                                        {
                                            QuantityProduct.Add(subcategory, Int32.Parse(reader["quantity"].ToString()));
                                        }
                                    }
                                }
                                else // If there is no data available
                                {
                                    QuantityProduct.Add(subcategory, 0);
                                }
                            }
                        }
                    }
                }
            }



            // End of for loop, array should be filled
            // Build Bar Chart
            if (comboBoxValue1.Text == "Quantity")
            {
                chart9.DataSource = QuantityProduct;
                chart9.Series[0].XValueMember = "Key";
                chart9.Series[0].YValueMembers = "Value";
                chart9.DataBind();
            }
            else if (comboBoxValue1.Text == "Profit")
            {
                chart9.DataSource = ProfitProduct;
                chart9.Series[0].XValueMember = "Key";
                chart9.Series[0].YValueMembers = "Value";
                chart9.DataBind();
            }
            else if (comboBoxValue1.Text == "Discount")
            {
                chart9.DataSource = DiscountProduct;
                chart9.Series[0].XValueMember = "Key";
                chart9.Series[0].YValueMembers = "Value";
                chart9.DataBind();
            }
            else if (comboBoxValue2.Text == "Value")
            {
                chart9.DataSource = ValueProduct;
                chart9.Series[0].XValueMember = "Key";
                chart9.Series[0].YValueMembers = "Value";
                chart9.DataBind();
            }

            if (comboBoxValue2.Text == "Quantity")
            {
                chart10.DataSource = QuantityProduct;
                chart10.Series[0].XValueMember = "Key";
                chart10.Series[0].YValueMembers = "Value";
                chart10.DataBind();
            }
            else if (comboBoxValue2.Text == "Profit")
            {
                chart10.DataSource = ProfitProduct;
                chart10.Series[0].XValueMember = "Key";
                chart10.Series[0].YValueMembers = "Value";
                chart10.DataBind();
            }
            else if (comboBoxValue2.Text == "Discount")
            {
                chart10.DataSource = DiscountProduct;
                chart10.Series[0].XValueMember = "Key";
                chart10.Series[0].YValueMembers = "Value";
                chart10.DataBind();
            }
            else if (comboBoxValue2.Text == "Value")
            {
                chart10.DataSource = ValueProduct;
                chart10.Series[0].XValueMember = "Key";
                chart10.Series[0].YValueMembers = "Value";
                chart10.DataBind();
            }
        }
        }
    }

   


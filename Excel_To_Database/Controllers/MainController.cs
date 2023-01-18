using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Data.OleDb;
using System.IO;

namespace Excel_To_Database.Controllers
{
    public class MainController : Controller
    {

        // GET: Main
    
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["ExcelData"].ConnectionString) ;

        OleDbConnection Econ;

        private void ExcelConn(string filepath)

        {

            string constr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;""", filepath);

            Econ = new OleDbConnection(constr);

        }

        public ActionResult Index()
        {
         
            return View();
        }
        
        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file)

        {

            string filename = Guid.NewGuid() + Path.GetExtension(file.FileName);

            string filepath = "/excelfolder/" + filename;

            file.SaveAs(Path.Combine(Server.MapPath("/excelfolder"), filename));

            InsertExceldata(filepath, filename);
            return View();
          //  return RedirectToAction("Create");

        }

        public ActionResult Create()
        {

            return View();
        }


        private void InsertExceldata(string fileepath, string filename)

        {

            string fullpath = Server.MapPath("/excelfolder/") + filename;

            ExcelConn(fullpath);

            string query = string.Format("Select * from [{0}]", "Sheet1$");

            OleDbCommand Ecom = new OleDbCommand(query, Econ);

            Econ.Open();

            DataSet ds = new DataSet();

            OleDbDataAdapter oda = new OleDbDataAdapter(query, Econ);

            Econ.Close();

            oda.Fill(ds);

            DataTable dt = ds.Tables[0];

            SqlBulkCopy objbulk = new SqlBulkCopy(con);

            objbulk.DestinationTableName = "ExcelFileImport";

             objbulk.ColumnMappings.Add("Date Collected", "DateCollected");
             objbulk.ColumnMappings.Add("Financial Year", "FinancialYear");
             objbulk.ColumnMappings.Add("Calendar Year", "CalendarYear");
             objbulk.ColumnMappings.Add("Quarter", "Quarter");
             objbulk.ColumnMappings.Add("Calendar Week", "Calendar_Week");
             objbulk.ColumnMappings.Add("Calendar Month", "Calendar_Month");
             objbulk.ColumnMappings.Add("Financial Month", "Financial_Month");
             objbulk.ColumnMappings.Add("Country", "Country");
             objbulk.ColumnMappings.Add("Store Banner", "Store_Banner");
             objbulk.ColumnMappings.Add("Store Type", "Store_Type");
             objbulk.ColumnMappings.Add("Seller Name", "Seller_Name");
             objbulk.ColumnMappings.Add("Product Division", "Product_Division");
             objbulk.ColumnMappings.Add("RPC", "RPC");
             objbulk.ColumnMappings.Add("MPC", "MPC");
             objbulk.ColumnMappings.Add("Product Name", "Product_Name");
             objbulk.ColumnMappings.Add("Product Rating", "Product_Rating");
             objbulk.ColumnMappings.Add("Listing", "Listing");
             objbulk.ColumnMappings.Add("Availability", "Availability");
             objbulk.ColumnMappings.Add("Availability of product title", "Availability_of_product_title");
             objbulk.ColumnMappings.Add("Availability of brand name in title", "Availability_of_brand_name_in_title");
             objbulk.ColumnMappings.Add("Title should have >6 words", "Title_should_have_Morethan_6_words");
             objbulk.ColumnMappings.Add("Availability of product description", "Availability_of_product_description");
             objbulk.ColumnMappings.Add("Desciption should have >15 words", "Desciption_should_have_Morethan_15_words");
             objbulk.ColumnMappings.Add("Availability of specifications", "Availability_of_specifications");
             objbulk.ColumnMappings.Add("Number  of specifications (>5)", "No_of_specifications_Morethan_5");
             objbulk.ColumnMappings.Add("Availability of image", "Availability_of_image");
             objbulk.ColumnMappings.Add("Number Of Images", "No_Of_Images");
             objbulk.ColumnMappings.Add("Number of images (>3)", "No_of_images_Morethan_3");
             objbulk.ColumnMappings.Add("Availability of customer reviews", "Availability_of_customer_reviews");
             objbulk.ColumnMappings.Add("Number of customer reviews (>21)", "No_of_customer_reviews_Morethan_21");
             objbulk.ColumnMappings.Add("Number Of Customer Reviews", "No_Of_Customer_Reviews");
             objbulk.ColumnMappings.Add("Availability of product rating", "Availability_of_product_rating");
             objbulk.ColumnMappings.Add("Product rating >4", "Product_rating_Morethan_4");
             objbulk.ColumnMappings.Add("Availability of Seller", "Availability_of_Seller");
             objbulk.ColumnMappings.Add("Availability of breadcrumbs", "Availability_of_breadcrumbs");
             objbulk.ColumnMappings.Add("Overall Score", "Overall_Score");
             objbulk.ColumnMappings.Add("Compliance Status", "Compliance_Status");
             objbulk.ColumnMappings.Add("URL", "URL");
             objbulk.ColumnMappings.Add("Cache Page Link", "Cache_Page_Link");
             objbulk.ColumnMappings.Add("Number of words in title", "Number_of_words_in_title");
             objbulk.ColumnMappings.Add("Number of words in description", "No_of_words_in_description");
             objbulk.ColumnMappings.Add("Number of bullets", "No_of_bullets");
             objbulk.ColumnMappings.Add("Product description", "Product_description");
             objbulk.ColumnMappings.Add("Specifications or bullets", "Specifications_or_bullets");
             objbulk.ColumnMappings.Add("Trusted product description", "Trusted_product_description");
             objbulk.ColumnMappings.Add("Trusted title", "Trusted_title");
             objbulk.ColumnMappings.Add("Trusted ratings", "Trusted_ratings");
             objbulk.ColumnMappings.Add("Trusted reviews", "Trusted_reviews");
             objbulk.ColumnMappings.Add("Video availability", "Video_availability");
             objbulk.ColumnMappings.Add("Color grouping", "Color_grouping");
             objbulk.ColumnMappings.Add("Sustainability", "Sustainability");
             objbulk.ColumnMappings.Add("Product rating vs trusted source", "Product_rating_vs_trusted_source");
             objbulk.ColumnMappings.Add("Number of reviews vs trusted source", "No_of_reviews_vs_trusted_source");
           
             
            con.Open();

            objbulk.WriteToServer(dt);

            con.Close();

        }
    }
}





/*
 *             objbulk.DestinationTableName = "ExcelFile";

            objbulk.ColumnMappings.Add("Date Collected", "Date Collected");
            objbulk.ColumnMappings.Add("Financial Year", "Financial Year");
            objbulk.ColumnMappings.Add("Calendar Year", "Calendar Year");
            objbulk.ColumnMappings.Add("Quarter", "Quarter");
            objbulk.ColumnMappings.Add("Calendar Week", "Calendar Week");
            objbulk.ColumnMappings.Add("Calendar Month", "Calendar Month");
            objbulk.ColumnMappings.Add("Financial Month", "Financial Month");
            objbulk.ColumnMappings.Add("Country", "Country");
            objbulk.ColumnMappings.Add("Store Banner", "Store Banner");
            objbulk.ColumnMappings.Add("Store Type", "Store Type");
            objbulk.ColumnMappings.Add("Seller Name", "Seller Name");
            objbulk.ColumnMappings.Add("Product Division", "Product Division");
            objbulk.ColumnMappings.Add("RPC", "RPC");
            objbulk.ColumnMappings.Add("MPC", "MPC");
            objbulk.ColumnMappings.Add("Product Name", "Product Name");
            objbulk.ColumnMappings.Add("Product Rating", "Product Rating");
            objbulk.ColumnMappings.Add("Listing", "Listing");
            objbulk.ColumnMappings.Add("Availability", "Availability");
            objbulk.ColumnMappings.Add("Availability of product title", "Availability of product title");
            objbulk.ColumnMappings.Add("Availability of brand name in title", "Availability of brand name in title");
            objbulk.ColumnMappings.Add("Title should have >6 words", "Title should have >6 words");
            objbulk.ColumnMappings.Add("Availability of product description", "Availability of product description");
            objbulk.ColumnMappings.Add("Desciption should have >15 words", "Desciption should have >15 words");
            objbulk.ColumnMappings.Add("Availability of specifications", "Availability of specifications");
            objbulk.ColumnMappings.Add("Number  of specifications (>5)", "No. of specifications (>5)");
            objbulk.ColumnMappings.Add("Availability of image", "Availability of image");
            objbulk.ColumnMappings.Add("Number Of Images", "No. Of Images");
            objbulk.ColumnMappings.Add("Number of images (>3)", "No. of images (>3)");
            objbulk.ColumnMappings.Add("Availability of customer reviews", "Availability of customer reviews");
            objbulk.ColumnMappings.Add("Number of customer reviews (>21)", "No. of customer reviews (>21)");
            objbulk.ColumnMappings.Add("Number Of Customer Reviews", "No. Of Customer Reviews");
            objbulk.ColumnMappings.Add("Availability of product rating", "Availability of product rating");
            objbulk.ColumnMappings.Add("Product rating >4", "Product rating >4");
            objbulk.ColumnMappings.Add("Availability of Seller", "Availability of Seller");
            objbulk.ColumnMappings.Add("Availability of breadcrumbs", "Availability of breadcrumbs");
            objbulk.ColumnMappings.Add("Overall Score", "Overall Score");
            objbulk.ColumnMappings.Add("Compliance Status", "Compliance Status");
            objbulk.ColumnMappings.Add("URL", "URL");
            objbulk.ColumnMappings.Add("Cache Page Link", "Cache Page Link");
            objbulk.ColumnMappings.Add("Number of words in title", "Number of words in title");
            objbulk.ColumnMappings.Add("Number of words in description", "No. of words in description");
            objbulk.ColumnMappings.Add("Number of bullets", "No of bullets");
            objbulk.ColumnMappings.Add("Product description", "Product description");
            objbulk.ColumnMappings.Add("Specifications or bullets", "Specifications / bullets");
            objbulk.ColumnMappings.Add("Trusted product description", "Trusted product description");
            objbulk.ColumnMappings.Add("Trusted title", "Trusted title");
            objbulk.ColumnMappings.Add("Trusted ratings", "Trusted ratings");
            objbulk.ColumnMappings.Add("Trusted reviews", "Trusted reviews");
            objbulk.ColumnMappings.Add("Video availability", "Video availability");
            objbulk.ColumnMappings.Add("Color grouping", "Color grouping");
            objbulk.ColumnMappings.Add("Sustainability", "Sustainability");
            objbulk.ColumnMappings.Add("Product rating vs trusted source", "Product rating vs trusted source");
            objbulk.ColumnMappings.Add("Number of reviews vs trusted source", "No. of reviews vs trusted source");
*/
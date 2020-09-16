using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Web;
using System.Web.Mvc;
using WebApplication4.Models;
using System.Data.SqlClient;
using System.Configuration;


using System.Collections.Generic;


namespace WebApplication4.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        /// <summary>
        /// Post method for importing Data 
        /// </summary>
        /// <param name="postedFile"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Index(HttpPostedFileBase[] postedFile)
        {

            DataSet ds = GetDataTablesFromExcel(postedFile);
            for (int i = 0; i < ds.Tables.Count; i++)
            {
                using (SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["Constring"].ConnectionString))
                {
                    using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                    {
                        sqlBulkCopy.DestinationTableName = ds.Tables[0].TableName;
                        con.Open();
                        sqlBulkCopy.WriteToServer(ds.Tables[0]);
                        con.Close();
                    }
                }
            }

            return View();
        }

        private DataSet GetDataTablesFromExcel(HttpPostedFileBase[] postedFiles)
        {
            DataSet ds = new DataSet();
            string path = Server.MapPath("~/UploadedFiles/");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            for (int j = 0; j < postedFiles.Length; j++)
            {

                if (postedFiles[j] != null)
                {
                    try
                    {
                        HttpPostedFileBase postedFile = postedFiles[j];
                        string filePath = string.Empty;
                        filePath = path + Path.GetFileName(postedFile.FileName);
                        string fileExtension = Path.GetExtension(postedFile.FileName);

                        //Validate uploaded file and return error.
                        if (fileExtension != ".xls" && fileExtension != ".xlsx")
                        {
                            ViewBag.Message = "Please select the excel file with .xls or .xlsx extension";
                           // return Index();
                        }

                      

                        //Save file to folder
                        postedFile.SaveAs(filePath);

                        //Get file extension

                        string excelConString = "";

                        //Get connection string using extension 
                        switch (fileExtension)
                        {
                            //If uploaded file is Excel 1997-2007.
                            case ".xls":
                                excelConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES'";
                                break;
                            //If uploaded file is Excel 2007 and above
                            case ".xlsx":
                                excelConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES'";
                                break;
                        }

                        //Read data from first sheet of excel into datatable
                        DataTable dt = new DataTable();
                        excelConString = string.Format(excelConString, filePath);

                        using (OleDbConnection excelOledbConnection = new OleDbConnection(excelConString))
                        {
                            using (OleDbCommand excelDbCommand = new OleDbCommand())
                            {
                                using (OleDbDataAdapter excelDataAdapter = new OleDbDataAdapter())
                                {
                                    excelDbCommand.Connection = excelOledbConnection;

                                    excelOledbConnection.Open();
                                    //Get schema from excel sheet
                                    DataTable excelSchema = GetSchemaFromExcel(excelOledbConnection);
                                    //Get sheet name
                                    string sheetName = excelSchema.Rows[0]["TABLE_NAME"].ToString();
                                    excelOledbConnection.Close();

                                    //Read Data from First Sheet.
                                    excelOledbConnection.Open();
                                    excelDbCommand.CommandText = "SELECT * From [" + sheetName + "]";
                                    excelDataAdapter.SelectCommand = excelDbCommand;
                                    //Fill datatable from adapter
                                    excelDataAdapter.Fill(dt);
                                    excelOledbConnection.Close();
                                }
                            }
                        }

                        GetSubmittedBusRouteDataFromExcelRow(dt);
                     
                        ViewBag.Message = "Data Imported Successfully.";
                    }
                    catch (Exception ex)
                    {
                        ViewBag.Message = ex.Message;
                    }
                }
                else
                {
                    ViewBag.Message = "Please select the file first to upload.";
                }
                //return View();
            }
            Directory.Delete(path, true);

            return ds;
        }
        private static DataTable GetSchemaFromExcel(OleDbConnection excelOledbConnection)
        {
            return excelOledbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
        }

        private void GetSubmittedBusRouteDataFromExcelRow(DataTable dt)
        {
            try
            {

                int StopNumber  ;
                double StopLatitude;
                double StopLongitude;
                var StopDescription ="" ;
                int AssignedStudents;



                var CountyDistrictCode = dt.Rows[6][2].ToString();
                var DistrictName = dt.Rows[7][2].ToString();
                var DistrictBusNumber = dt.Rows[8][2].ToString();
                var DistrictRouteNumber = dt.Rows[9][2].ToString();
                var RouteTypeCode = dt.Rows[10][2].ToString();
                var StateBusNumber = dt.Rows[11][2].ToString();
                var StateRouteNumber = int.Parse(dt.Rows[12][2].ToString());
                var DestinationName = "";
                var DestinationIdentifier = "";
                var DestinationLatitude = "";
                var DestinationLongitude = "";

                //Loop through datatable and add employee data to employee table. 
                using (var context = new DemoContext())
                {
                    for (int m = 0; m <= 2; m++)
                    {

                        if (m == 0)
                        {
                            DestinationName = dt.Rows[8][6].ToString();
                            DestinationIdentifier = dt.Rows[9][6].ToString();
                            DestinationLatitude = dt.Rows[11][6].ToString();
                            DestinationLongitude = dt.Rows[12][6].ToString(); ;

                        }
                        else
                        {
                            DestinationName = dt.Rows[8][9].ToString();
                            DestinationIdentifier = dt.Rows[9][9].ToString();
                            DestinationLatitude = dt.Rows[11][9].ToString();
                            DestinationLongitude = dt.Rows[12][9].ToString(); ;
                        }
                        for (int i = 15; i <= 41; i++)
                        {
                            StopNumber=int.Parse(dt.Rows[i][0].ToString());
                            StopLatitude=double.Parse(dt.Rows[i][1].ToString());
                            StopLongitude=double.Parse(dt.Rows[i][2].ToString());
                            StopDescription=dt.Rows[i][3].ToString();
                            AssignedStudents=int.Parse(dt.Rows[i][6].ToString());

                            var saveObj = new STARS_SubmittedRouteData
                            {
                                StateBusNumber = StateBusNumber,
                                SubmittedRouteDataId = 202,
                                ImportHistoryId = 312,
                                CountyDistrictCode = CountyDistrictCode,
                                RouteTypeCode = RouteTypeCode,
                                DistrictRouteNumber = DistrictRouteNumber,
                                DistrictBusNumber = DistrictBusNumber,
                                StateRouteNumber = StateRouteNumber,
                                StopNumber = StopNumber,
                                StopLatitude = StopLatitude,
                                StopLongitude = StopLongitude,
                                StopDescription = StopDescription,
                                DestinationName = DestinationName,
                                DestinationIdentifier = DestinationIdentifier,
                                DestinationLatitude = DestinationLatitude,
                                DestinationLongitude = DestinationLongitude,
                                AssignedStudents = AssignedStudents,
                                DistrictName = DistrictName,
                            };
                            context.STARS_SubmittedRouteData.Add(saveObj);
                        }
                    }
                    context.SaveChanges();
                }
            
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
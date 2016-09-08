using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using System.Net;
using System.Net.Mail;
using System.Text;
 
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;

using BrandGoogleAnalyticsData.Models.DB;
using System.Data;
using System.Data.Objects;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;

 
//using InventoryFeed.Models.ViewModel;
using System.IO.Compression;
using System.Web.Script.Serialization;


namespace BrandGoogleAnalyticsData.Controllers
{

    public class CustomerController : Controller
    {
        IFSReportingContext db = new IFSReportingContext();
        //
        // GET: /Customer/

        public static void killProcessByName(string processName)
        {
            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName(processName);
            foreach (System.Diagnostics.Process p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch { }
                }
            }
        }

        public string GetFileName(string path) {
            string fileName="";
              Uri uri = new Uri(path);
                            if (uri.IsFile)
                                fileName = System.IO.Path.GetFileName(uri.LocalPath);
            return fileName;
                            
        }

        private static void GetPathParams(out string localPath, out string localPathwofile)
        {
            string path = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            //once you have the path you get the directory with:
            var directory = System.IO.Path.GetDirectoryName(path);
            localPath = new Uri(directory).LocalPath;
            localPathwofile = localPath + "\\..\\App_Data";
        }

        public JsonResult  GAnalytics(string path)
        {
            //string path = "C:\\Users\\Abel\\Downloads\\AugustAnalytics\\BMRACING.xlsx";

            var _app = new Excel.Application();
            var _workbooks = _app.Workbooks.Open(path);
            var _worksheet = _workbooks.ActiveSheet;
            string cell_data ="";

            Microsoft.Office.Interop.Excel.Range range = _worksheet.UsedRange;
            //Read the first cell

           
           string text_value= "";
            int end_row = range.Rows.Count;
            var end_column = range.Columns.Count;
            var group_field = "";
            var subgroup_field_text = "";
            var subgroup_field_column2 = "";
            var subgroup_field_column3 = "";
            var subgroup_field_column4 = "";
            var insert = false;
            var value = "";
            var name="";
            //[row,column]
            
            for (var index_row = 1; index_row <= end_row; index_row++) {
                insert = true;
                value = "";
                name="";
                
                for (var index_column = 1; index_column <= end_column; index_column++)
                {
                        text_value = _worksheet.Cells[index_row, index_column].Text.ToString();

                        if (text_value != "" || index_column == 2 || index_column == 3 || index_column == 4)
                        {
                            if (index_column == 1)
                            {
                                if (text_value == "Audience" ||
                                    text_value == "New vs Returning Visitor" ||
                                    text_value == "Devices" ||
                                    text_value == "Acquisition" ||
                                    text_value == "Behavior" ||
                                    text_value == "Site Speed"
                                )
                                {
                                    group_field = text_value;
                                    insert = false;
                                }
                                else if (text_value.Contains("Country"))
                                {
                                    group_field = "Country";
                                    insert = false;
                                }

                            }

                            if (!insert)
                            {
                                
                                if (index_column == 2)
                                {
                                    subgroup_field_column2 = text_value;
                                }
                                else if (index_column == 3)
                                {
                                    subgroup_field_column3 = text_value;
                                    //subgroup_field_text = subgroup_field_column3;
                                }
                                else if (index_column == 4)
                                {
                                    subgroup_field_column4 = text_value;
                                   // subgroup_field_text = subgroup_field_column4;
                                }
                                
                            }
                            else
                            { 
                                if (index_column == 1)
                                {
                                    name = text_value;
                                }
                                if (index_column == 2)
                                {
                                    subgroup_field_text = subgroup_field_column2;
                                    value = text_value;
                                }
                                else if (index_column == 3)
                                {
                                    subgroup_field_text = subgroup_field_column3;
                                    value = text_value;
                                }
                                else if (index_column == 4)
                                {
                                    subgroup_field_text = subgroup_field_column4;
                                    value = text_value;
                                }
                            }

                            string fileName = Path.GetFileNameWithoutExtension(path);

                            //cell_data += group_field + index_row + ":" + index_column + " " + text_value;
                            if (insert && value != "")
                            {
                                Library.Execute(@"insert into tblBrandGoogleAanalyticsData 
                                  (brand,group_field,subgroup_id, subgroup_field, name, value)
                                    values('" + fileName + "','" + group_field + "','" + (index_column - 1) + "','" + subgroup_field_text + "','" + name + "','" + value + @"')
                                ");
                                // cell_data += "group_field: " + group_field + " subgroup_field:(" + (index_column - 1) + ")" + subgroup_field_text + " " + "name: " + name + " " + value;
                            }
                        }
                        else
                        {
                           // subgroup_field_column2 = ""; subgroup_field_column3 = ""; subgroup_field_column4 = "";
                        }
                       
                 }
              }
                

            _workbooks.Close();

            _app.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(_app);


            return Json(new { 
                path=path, cell_data = cell_data,
                 end_column = end_column, end_row = end_row });
        }

        public ActionResult ZipRead()
        {
            string localPath, extractPath;
            GetPathParams(out localPath, out extractPath);

            string zipPath = @"C:\\Users\\Abel\\Downloads\\February.zip";

           // string extractPath = @"C:\\Users\\Abel\\Downloads\\extract";
            string files = "", extract_files_path="";
            string content=@"";
            string filenameonly = "";

            Library.Execute("delete from tblBrandGoogleAanalyticsData");


            killProcessByName("Excel");
            using (ZipArchive archive = ZipFile.OpenRead(zipPath))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.FullName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        filenameonly = Path.GetFileName(entry.FullName);

                        
                        files += filenameonly;
                        extract_files_path = extractPath + "\\" + filenameonly;

                      

                        if ((System.IO.File.Exists(extract_files_path)))
                        {
                            System.IO.File.Delete(extract_files_path);
                        }
                        entry.ExtractToFile(Path.Combine(extractPath, filenameonly));
                        //content += extract_files_path;
                      
                        content += new JavaScriptSerializer().Serialize(GAnalytics(extract_files_path).Data); 
                    }
                }
            }

            killProcessByName("Excel");

            return Content(content);
        }


        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Feed()
        {
            var feed = (from m in db.tblInventoryFeeds
                          select m).OrderByDescending(x => x.if_id);;

            return View(feed); 
        }

        public ActionResult Create()
        {

            return View();
        }

        public ActionResult Update(int id)
        {
                int if_id = id;
                var edit_query = db.tblInventoryFeeds.Where(m => m.if_id == if_id).FirstOrDefault();

                /*var edit_query = from m in db.tblInventoryFeeds
                                 where m.if_id == if_id
                                 select m; */
                return View(edit_query);
        }

    
        public ActionResult Process(string id)
        {
            IQueryable process;

            if (id == null)
                process = from m in db.tblInventoryFeedProcesses
                          select m;
            else
            {
                int if_id;
                bool check = int.TryParse(id, out if_id);
  
                process = from m in db.tblInventoryFeedProcesses
                          where m.if_id == if_id
                          select m;
            }

            return View(process);
        }
         
       
       

        public ActionResult Datatables()
        {
            return View();
        }

        // GET: /Customer/Name/{customer_no}
        
        public JsonResult Name(string customer_name) {
           
            //return (Json(customer_no));
            return Json(db.tblInvoiceLinesMasters.Select(m => new { m.CUSTOMER_NO, m.CUSTOMER_NAME }).Where(m => m.CUSTOMER_NAME == customer_name).Distinct(), JsonRequestBehavior.AllowGet);
        }

        // GET: /Customer/Number/{customer_name}

        public JsonResult Number(string customer_no) {
            return Json(db.tblInvoiceLinesMasters.Select(m => new { m.CUSTOMER_NO, m.CUSTOMER_NAME }).Where(m => m.CUSTOMER_NO== customer_no).Distinct(), JsonRequestBehavior.AllowGet);
        }

      


        public void AddEditSameFunction(int if_id,string sendtime) {
            string[] separators = { ",", ";", " " };
            string[] sendtime_array = Library.Explode(sendtime, separators);

            foreach (string individual_sendtime in sendtime_array)
            {
                tblInventoryFeedProcess subreq = new tblInventoryFeedProcess
                {
                    if_id = if_id,
                    status = "0",
                    time_split = TimeSpan.Parse(individual_sendtime, System.Globalization.CultureInfo.CurrentCulture),
                };
                db.tblInventoryFeedProcesses.Add(subreq);
                db.SaveChanges();
            } 
        }
    

        [AllowAnonymous]
        [HttpPost]
        public JsonResult SendRequest(string customer_no)
        {
            try
            {
                var protocol_addr="";

                if (Request["sendvia"] == "email") protocol_addr = Request["email"];
                else if (Request["sendvia"] == "ftp") protocol_addr = Request["ftp"];
                tblInventoryFeed req = new tblInventoryFeed {
                                customer_no = customer_no,
                                 filetype_requested = Request["type"],
                                 sendaaid_instead_brand_id = Request["sendid"],
                                 send_protocol = Request["sendvia"],
                                 protocol_address = protocol_addr.Replace(",",";"),
                                 
                                 sendbuyers_partno = Request["buyer"],
                                includeheaders = Request["header"],
                                sendtime = Request["time"].Replace(",",";"),
                                fields = Request["field"].Replace(",",";"),
                                sendday = Request["day"]
                          };

                db.tblInventoryFeeds.Add(req);
                db.SaveChanges();

                int if_id = req.if_id;  //the last id inserted

                string sendtime = Request["time"].Replace(",", ";");
                string[] separators = { ",", ";", " " };
                string[] sendtime_array = Library.Explode(sendtime, separators);

                AddEditSameFunction(if_id, sendtime);

                return Json(new { message = req});
            }
            catch (Exception ex)
            {
                return Json(new { message = ex.Message });

            }
        }

        [HttpPost]
        public JsonResult FeedDelete(int if_id)
        {
            IFSReportingContext db_local = new IFSReportingContext();
            tblInventoryFeed inv = new tblInventoryFeed(){ //selecting for update
                if_id = if_id
            };

            db_local.tblInventoryFeeds.Attach(inv);  
            db_local.tblInventoryFeeds.Remove(inv);
            db_local.SaveChanges();

            Library.Execute("DELETE tblInventoryFeedProcess WHERE if_id = " + if_id);
   
            return Json(new { message = "deleted" });
        }


        [HttpPost]
        public JsonResult FeedEdit(int if_id)
        {
            IFSReportingContext db_local = new IFSReportingContext();
            tblInventoryFeed inv = new tblInventoryFeed() //selecting for update
            {
                if_id = if_id,
            };

            db_local.tblInventoryFeeds.Attach(inv);

            var protocol_addr = "";

            if (Request["sendvia"] == "email") protocol_addr = Request["email"];
            else if (Request["sendvia"] == "ftp") protocol_addr = Request["ftp"];

            inv.customer_no = Request["customer_no"];
             inv.filetype_requested = Request["type"];
             inv.sendaaid_instead_brand_id = Request["sendid"];
             inv.send_protocol = Request["sendvia"];
             inv.protocol_address = protocol_addr.Replace(",", ";");

              inv.sendbuyers_partno= Request["buyer"];
              inv.includeheaders = Request["header"];
             inv.sendtime = Request["time"].Replace(",", ";");
             inv.fields = Request["field"].Replace(",", ";");
             inv.sendday = Request["day"];
            
            db_local.SaveChanges();

            Library.Execute("DELETE tblInventoryFeedProcess WHERE if_id = " + if_id);
            
            string sendtime = Request["time"].Replace(",", ";");
            string[] separators = { ",", ";", " " };
            string[] sendtime_array = Library.Explode(sendtime, separators);

            AddEditSameFunction(if_id,sendtime);

            return Json(new { message = "edit" });
        }

       
        public JsonResult FtpTest()
        {

            string ftp_credentials = Request["ftp"];

            string[] separators = { ",", ";", " " };
           
            string[] words = ftp_credentials.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            string ftphostname = "", ftpusername = "", ftppassword = "", ftpfolder = "";

            if (words.Length <= 4) // ftp host, username, password
            {
                ftphostname = words[0];
                ftpusername = words[1];
                ftppassword = words[2];
                if (words.Length == 4)
                    ftpfolder = words[3];
                else
                    ftpfolder = "";
            }

            FtpWebRequest ftp = (FtpWebRequest)FtpWebRequest.Create("ftp://"+ftphostname+"/"+ftpfolder+"/test.txt");
            FtpWebResponse res;
            StreamReader reader;

            ftp.Credentials = new NetworkCredential(ftpusername, ftppassword);
            ftp.KeepAlive = false;
            ftp.Method = WebRequestMethods.Ftp.ListDirectoryDetails;
            ftp.UsePassive = true;

            try
            {
                using (res = (FtpWebResponse)ftp.GetResponse())
                {
                    reader = new StreamReader(res.GetResponseStream());
                }
                 return Json(new { message = "FTP Verified", status="success" });
            }
            catch
            {
                return Json(new { message = "The system cannot connect to the specified ftp credentials", status = "failed" });
         
            }
        }



    }
}

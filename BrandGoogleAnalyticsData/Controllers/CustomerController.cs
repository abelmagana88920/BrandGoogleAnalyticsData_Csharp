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

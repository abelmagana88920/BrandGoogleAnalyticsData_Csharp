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
    public class BrandController : Controller
    {
        //
        // GET: /Brand/

        public ActionResult Index()
        {
            return View();
        }

         
        public ActionResult Upload()
        {
            return View();
        }

        [HttpPost]
        public ActionResult UploadSubmit(HttpPostedFileBase upload_file)
        {
            var path = "";
            string status = "error";
            if (upload_file != null && upload_file.ContentLength > 0)
            {
                // extract only the filename
                var fileName = Path.GetFileName(upload_file.FileName);
                var extension = Path.GetExtension(upload_file.FileName);
                if (extension != ".zip") return new HttpStatusCodeResult(404);

                // store the file inside ~/App_Data/uploads folder
                path = Path.Combine(Server.MapPath("~/App_Data/uploads"), fileName);
                upload_file.SaveAs(path);

                if ((System.IO.File.Exists(path)))
                {
                    status = ZipRead(path);
                }
                if (status == "success") return new HttpStatusCodeResult(200);
                else return new HttpStatusCodeResult(404);

                
            }

            return new HttpStatusCodeResult(404);
           

            
             
        }

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

        public string GetFileName(string path)
        {
            string fileName = "";
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
            localPathwofile = localPath + "\\..\\App_Data\\extracts";
        }

        public JsonResult GAnalytics(string path, int month, int year)
        {
            //string path = "C:\\Users\\Abel\\Downloads\\AugustAnalytics\\BMRACING.xlsx";

            var _app = new Excel.Application();
            var _workbooks = _app.Workbooks.Open(path);
            var _worksheet = _workbooks.ActiveSheet;
            string cell_data = "";

            Microsoft.Office.Interop.Excel.Range range = _worksheet.UsedRange;
            //Read the first cell


            string text_value = "";
            int end_row = range.Rows.Count;
            var end_column = range.Columns.Count;
            var group_field = "";
            var subgroup_field_text = "";
            var subgroup_field_column2 = "";
            var subgroup_field_column3 = "";
            var subgroup_field_column4 = "";
            var insert = false;
            var value = "";
            var name = "";
            //[row,column]

            for (var index_row = 1; index_row <= end_row; index_row++)
            {
                insert = true;
                value = "";
                name = "";

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
                                  (brand,group_field,subgroup_id, subgroup_field, name, value, month, year)
                                    values('" + fileName + "','" + group_field + "','" + (index_column - 1) + "','" + subgroup_field_text + "','" + name + "','" + value + "','"+month+"','"+year+@"')
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


            return Json(new
            {
                path = path,
                cell_data = cell_data,
                end_column = end_column,
                end_row = end_row
            });
        }

        public string ZipRead(string uploadfile_path)
        {
            string localPath, extractPath;
            GetPathParams(out localPath, out extractPath);

            string zipPath = @""+uploadfile_path+"";

            string month_year = Path.GetFileNameWithoutExtension(uploadfile_path);
            string[] month_year_arr = month_year.Split('-');

            int month;
            int year;

            try
            {
                month = DateTime.ParseExact(month_year_arr[0], "MMMM", System.Globalization.CultureInfo.InvariantCulture).Month;
                year = Int32.Parse(month_year_arr[1]);
            }
            catch
            {
                return "error";
            }

            // string extractPath = @"C:\\Users\\Abel\\Downloads\\extract";
            string files = "", extract_files_path = "";
            string content = @"";
            string filenameonly = "";

            Library.Execute("delete from tblBrandGoogleAanalyticsData where month='"+month+"' and year='"+year+"'");

            try
            {
                killProcessByName("Excel");
                using (ZipArchive archive = ZipFile.OpenRead(zipPath))
                {
                    bool havingexcel_files = false;
                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        if (entry.FullName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                        {
                            havingexcel_files = true;

                            filenameonly = Path.GetFileName(entry.FullName);


                            files += filenameonly;
                            extract_files_path = extractPath + "\\" + month + "-"+ year + "-" + filenameonly;

                            if ((System.IO.File.Exists(extract_files_path)))
                            {
                                System.IO.File.Delete(extract_files_path);
                            }

                            if (!extract_files_path.ToLower().Contains('~')) //if ~ does not contaimn
                            {
                                entry.ExtractToFile(Path.Combine(extractPath, month + "-" + year + "-" + filenameonly));
                                //content += extract_files_path;
                                content += new JavaScriptSerializer().Serialize(GAnalytics(extract_files_path, month, year).Data);
                            }
                        }
                    }
                    if (!havingexcel_files)
                    {
                        killProcessByName("Excel");
                        return "error";
                    } 
                }

                killProcessByName("Excel");
                return "success";
            }
           catch
            {
                return "error";
            }


           
        }






    }
}

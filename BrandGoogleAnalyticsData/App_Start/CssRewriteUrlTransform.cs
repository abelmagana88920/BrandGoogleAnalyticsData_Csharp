using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.IO;
using BrandGoogleAnalyticsData.App_Start;
using System.Web;
using System.Web.Optimization;
using System.Text.RegularExpressions;

namespace BrandGoogleAnalyticsData.App_Start
{
    public class CssUrlTransform : IBundleTransform
    {
        public void Process(BundleContext context, BundleResponse response)
        {
            Regex exp = new Regex(@"url\([^\)]+\)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            foreach (FileInfo css in response.Files)
            {
                string cssAppRelativePath = css.FullName.Replace(context.HttpContext.Request.PhysicalApplicationPath, context.HttpContext.Request.ApplicationPath).Replace(Path.DirectorySeparatorChar, '/');
                string cssDir = cssAppRelativePath.Substring(0, cssAppRelativePath.LastIndexOf('/'));
                response.Content = exp.Replace(response.Content, m => TransformUrl(m, cssDir));
            }
        }


        private string TransformUrl(Match match, string cssDir)
        {
            string url = match.Value.Substring(4, match.Length - 5).Trim('\'', '"');

            if (url.StartsWith("http://") || url.StartsWith("data:image")) return match.Value;

            if (!url.StartsWith("/"))
                url = string.Format("{0}/{1}", cssDir, url);

            return string.Format("url({0})", url);
        }

    }
    
}

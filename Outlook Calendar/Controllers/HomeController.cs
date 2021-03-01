using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Outlook_Calendar.Controllers
{
    public class HomeController : Controller
    {
        string credentialsFile = "C:\\Users\\User Name\\Desktop\\Outlook Calendar\\Outlook Calendar\\Files\\credentials.json";
        string adminCredentialsFile = "C:\\Users\\User Name\\Desktop\\Outlook Calendar\\Outlook Calendar\\Files\\adminCredentials.json";

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult OauthRedirect()
        {

            JObject credentials = JObject.Parse(System.IO.File.ReadAllText(credentialsFile));

            var redirectUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?" +
                               "&scope=" + credentials["scopes"].ToString() +
                               "&response_type=code" +
                               "&response_mode=query" +
                               "&state=themessydeveloper" +
                               "&redirect_uri=" + credentials["redirect_url"].ToString() +
                               "&client_id=" + credentials["client_id"].ToString();

            return Redirect(redirectUrl);
        }

        public ActionResult AdminOauthRedirect()
        {

            JObject credentials = JObject.Parse(System.IO.File.ReadAllText(adminCredentialsFile));

            var redirectUrl = "https://login.microsoftonline.com/common/adminconsent?" +
                               "&state=themessydeveloper" +
                               "&redirect_uri=" + credentials["redirect_url"].ToString() +
                               "&client_id=" + credentials["client_id"].ToString();

            return Redirect(redirectUrl);
        }
    }
}
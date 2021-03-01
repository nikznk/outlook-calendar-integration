using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Outlook_Calendar.Controllers
{
    public class OAuthController : Controller
    {
        string credentialsFile = "C:\\Users\\User Name\\Desktop\\Outlook Calendar\\Outlook Calendar\\Files\\credentials.json";
        string adminCredentialsFile = "C:\\Users\\User Name\\Desktop\\Outlook Calendar\\Outlook Calendar\\Files\\adminCredentials.json";
        string tokensFile = "C:\\Users\\User Name\\Desktop\\Outlook Calendar\\Outlook Calendar\\Files\\tokens.json";
        string adminTokensFile = "C:\\Users\\User Name\\Desktop\\Outlook Calendar\\Outlook Calendar\\Files\\adminTokens.json";

        public ActionResult Callback(string code, string state, string error)
        {
            JObject credentials = JObject.Parse(System.IO.File.ReadAllText(credentialsFile));

            if (!string.IsNullOrWhiteSpace(code))
            {
                RestClient restClient = new RestClient();
                RestRequest restRequest = new RestRequest();

                restRequest.AddParameter("client_id", credentials["client_id"].ToString());
                restRequest.AddParameter("scope", credentials["scopes"].ToString());
                restRequest.AddParameter("redirect_uri", credentials["redirect_url"].ToString());
                restRequest.AddParameter("code", code);
                restRequest.AddParameter("grant_type", "authorization_code");
                restRequest.AddParameter("client_secret", credentials["client_secret"].ToString());

                restClient.BaseUrl = new Uri("https://login.microsoftonline.com/common/oauth2/v2.0/token");
                var response = restClient.Post(restRequest);

                if (response.StatusCode == System.Net.HttpStatusCode.OK)
                {
                    System.IO.File.WriteAllText(tokensFile, response.Content);
                    return RedirectToAction("Index", "Home");
                }
            }

            return RedirectToAction("Error");
        }

        public ActionResult AdminCallback(string tenant, string state, string admin_consent)
        {
            JObject credentials = JObject.Parse(System.IO.File.ReadAllText(adminCredentialsFile));
            
            if (!string.IsNullOrWhiteSpace(tenant))
            {
                RestClient restClient = new RestClient();
                RestRequest restRequest = new RestRequest();

                restRequest.AddParameter("client_id", credentials["client_id"].ToString());
                restRequest.AddParameter("scope", credentials["scopes"].ToString());
                restRequest.AddParameter("grant_type", "client_credentials");
                restRequest.AddParameter("client_secret", credentials["client_secret"].ToString());

                restClient.BaseUrl = new Uri($"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token");
                var response = restClient.Post(restRequest);

                if (response.StatusCode == System.Net.HttpStatusCode.OK)
                {
                    JObject content = JObject.Parse(response.Content);
                    content["tenant"] = tenant;
                    System.IO.File.WriteAllText(adminTokensFile, content.ToString());
                    return RedirectToAction("Index", "Home");
                }
            }

            return RedirectToAction("Error");
        }

        public ActionResult RefreshToken()
        {
            JObject credentials = JObject.Parse(System.IO.File.ReadAllText(credentialsFile));
            JObject tokens = JObject.Parse(System.IO.File.ReadAllText(tokensFile));

            RestClient restClient = new RestClient();
            RestRequest restRequest = new RestRequest();

            restRequest.AddParameter("client_id", credentials["client_id"].ToString());
            restRequest.AddParameter("grant_type", "refresh_token");
            restRequest.AddParameter("scope", credentials["scopes"].ToString());
            restRequest.AddParameter("refresh_token", tokens["refresh_token"].ToString());
            restRequest.AddParameter("redirect_uri", credentials["redirect_url"].ToString());
            restRequest.AddParameter("client_secret", credentials["client_secret"].ToString());

            restClient.BaseUrl = new Uri("https://login.microsoftonline.com/common/oauth2/v2.0/token");
            var response = restClient.Post(restRequest);

            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                System.IO.File.WriteAllText(tokensFile, response.Content);
                return RedirectToAction("Index", "Home");
            }

            return RedirectToAction("Error");
        }
    }
}
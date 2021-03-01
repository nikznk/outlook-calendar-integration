using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Outlook_Calendar.Models;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Web.Mvc;

namespace Outlook_Calendar.Controllers
{
    public class CalendarController : Controller
    {
        string tokensFile = "C:\\Users\\User Name\\Desktop\\Outlook Calendar\\Outlook Calendar\\Files\\tokens.json";

        public ActionResult CreateEvent(CalendarEvent calendarEvent)
        {
            JObject tokens = JObject.Parse(System.IO.File.ReadAllText(tokensFile));

            RestClient restClient = new RestClient();
            RestRequest restRequest = new RestRequest();

            restRequest.AddHeader("Authorization", "Bearer " + tokens["access_token"].ToString());
            restRequest.AddHeader("Content-Type", "application/json");
            restRequest.AddParameter("application/json", JsonConvert.SerializeObject(calendarEvent), ParameterType.RequestBody);

            restClient.BaseUrl = new Uri("https://graph.microsoft.com/v1.0/me/calendar/events");
            var response = restClient.Post(restRequest);

            if (response.StatusCode == System.Net.HttpStatusCode.Created)
            {
                return RedirectToAction("Index", "Home");
            }

            return RedirectToAction("Error");
        }

        public ActionResult Event(string eventId)
        {
            JObject tokens = JObject.Parse(System.IO.File.ReadAllText(tokensFile));

            RestClient restClient = new RestClient();
            RestRequest restRequest = new RestRequest();

            restRequest.AddHeader("Authorization", "Bearer " + tokens["access_token"].ToString());
            restRequest.AddHeader("Prefer", "outlook.timezone=\"India Standard Time\"");
            restRequest.AddHeader("Prefer", "outlook.body-content-type=\"text\"");

            restClient.BaseUrl = new Uri($"https://graph.microsoft.com/v1.0/me/calendar/events/{eventId}");
            var response = restClient.Get(restRequest);

            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                CalendarEvent calendarEvent = JObject.Parse(response.Content).ToObject<CalendarEvent>();
                return View(calendarEvent);
            }

            return RedirectToAction("Error");
        }

        public ActionResult AllEvents()
        {
            JObject tokens = JObject.Parse(System.IO.File.ReadAllText(tokensFile));

            RestClient restClient = new RestClient();
            RestRequest restRequest = new RestRequest();

            restRequest.AddHeader("Authorization", "Bearer " + tokens["access_token"].ToString());
            restRequest.AddHeader("Prefer", "outlook.timezone=\"India Standard Time\"");
            restRequest.AddHeader("Prefer", "outlook.body-content-type=\"text\"");

            restClient.BaseUrl = new Uri("https://graph.microsoft.com/v1.0/me/calendar/events");
            var response = restClient.Get(restRequest);

            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                JObject eventsList = JObject.Parse(response.Content);
                var calendarEvents = eventsList["value"].ToObject<IEnumerable<CalendarEvent>>();
                return View(calendarEvents);
            }

            return RedirectToAction("Error");
        }

        public ActionResult UpdateEvent(string eventId)
        {
            JObject tokens = JObject.Parse(System.IO.File.ReadAllText(tokensFile));

            RestClient restClient = new RestClient();
            RestRequest restRequest = new RestRequest();

            restRequest.AddHeader("Authorization", "Bearer " + tokens["access_token"].ToString());
            restRequest.AddHeader("Prefer", "outlook.timezone=\"India Standard Time\"");
            restRequest.AddHeader("Prefer", "outlook.body-content-type=\"text\"");

            restClient.BaseUrl = new Uri($"https://graph.microsoft.com/v1.0/me/calendar/events/{eventId}");
            var response = restClient.Get(restRequest);

            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                CalendarEvent calendarEvent = JObject.Parse(response.Content).ToObject<CalendarEvent>();
                return View(calendarEvent);
            }

            return RedirectToAction("Error");
        }

        [HttpPost]
        public ActionResult UpdateEvent(string eventId, CalendarEvent calendarEvent)
        {
            JObject tokens = JObject.Parse(System.IO.File.ReadAllText(tokensFile));

            RestClient restClient = new RestClient();
            RestRequest restRequest = new RestRequest();

            restRequest.AddHeader("Authorization", "Bearer " + tokens["access_token"].ToString());
            restRequest.AddHeader("Content-Type", "application/json");
            restRequest.AddParameter("application/json", JsonConvert.SerializeObject(calendarEvent), ParameterType.RequestBody);

            restClient.BaseUrl = new Uri($"https://graph.microsoft.com/v1.0/me/calendar/events/{eventId}");
            var response = restClient.Patch(restRequest);

            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                return RedirectToAction("AllEvents", "Calendar");
            }

            return RedirectToAction("Error");
        }

        public ActionResult DeleteEvent(string eventId)
        {
            JObject tokens = JObject.Parse(System.IO.File.ReadAllText(tokensFile));

            RestClient restClient = new RestClient();
            RestRequest restRequest = new RestRequest();

            restRequest.AddHeader("Authorization", "Bearer " + tokens["access_token"].ToString());

            restClient.BaseUrl = new Uri($"https://graph.microsoft.com/v1.0/me/calendar/events/{eventId}");
            var response = restClient.Delete(restRequest);

            if (response.StatusCode == System.Net.HttpStatusCode.NoContent)
            {
                return RedirectToAction("AllEvents", "Calendar");
            }

            return RedirectToAction("Error");
        }
    }
}
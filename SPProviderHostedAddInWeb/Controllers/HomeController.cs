using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using Microsoft.IdentityModel.S2S.Tokens;
using System.Net;
using System.IO;
using System.Xml;

using SPProviderHostedAddInWeb.Models;

namespace SPProviderHostedAddInWeb.Controllers
{
    public class HomeController : Controller
    {

        [SharePointContextFilter]
        public ActionResult Index()
        {
            User spUser = null;
            SPModel spModel = new SPModel();

            // factor out into property and put in session
            string contextTokenString = TokenHelper.GetContextTokenFromRequest(Request);

            if (contextTokenString != null)
            {
                spModel.ContextToken = TokenHelper.ReadAndValidateContextToken(contextTokenString, Request.Url.Authority);
                spModel.SharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
                spModel.AccessToken = TokenHelper.GetAccessToken(spModel.ContextToken, spModel.SharepointUrl.Authority).AccessToken;

               // For simplicity, this sample assigns the access token to the button's CommandArgument property. 
               // In a production add-in, this would not be secure. The access token should be cached on the server-side.
               // CSOM.CommandArgument = accessToken; - change this to an action link Post
            }


            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            using (spModel.SPClientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (spModel.SPClientContext != null)
                {
                    //spModel.SPClientContext.

                    spUser = spModel.SPClientContext.Web.CurrentUser;
                    spModel.SPClientContext.Load(spUser, user => user.Title);
                    spModel.SPClientContext.ExecuteQuery();
                    ViewBag.UserName = spUser.Title;
                }
            }

            return View(spModel);
        }

        // This method retrieves information about the host web by using the CSOM. Could be in a Data Layer sep component - but not today
        private SPModel RetrieveWithCSOM(SPModel spModel)
        {

            //if (IsPostBack)
            //{
            //    sharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
            //}

            //spModel.SPClientContext = TokenHelper.GetClientContextWithAccessToken(sharepointUrl.ToString(), accessToken);

            // Load the properties for the web object.
            Web web = spModel.SPClientContext.Web;
            spModel.SPClientContext.Load(web);
            spModel.SPClientContext.ExecuteQuery();

            // Get the site name.
            spModel.SiteName = web.Title;

            // Get the current user.
            spModel.SPClientContext.Load(web.CurrentUser);
            spModel.SPClientContext.ExecuteQuery();
            spModel.CurrentUser = spModel.SPClientContext.Web.CurrentUser.LoginName;

            // Load the lists from the Web object.
            ListCollection lists = web.Lists;
            spModel.SPClientContext.Load<ListCollection>(lists);
            spModel.SPClientContext.ExecuteQuery();

            // Load the current users from the Web object.
            UserCollection users = web.SiteUsers;
            spModel.SPClientContext.Load<UserCollection>(users);
            spModel.SPClientContext.ExecuteQuery();

            foreach (User siteUser in users)
            {
                spModel.ListOfUsers.Add(siteUser.LoginName);
            }


            foreach (List list in lists)
            {
                spModel.ListOfLists.Add(list.Title);
            }

            return spModel;
        }

        public ActionResult CSOM_Click()
        {
            ViewBag.Message = "CSOM Clicked View";

            User spUser = null;
            SPModel spModel = new SPModel();


            string contextTokenString = TokenHelper.GetContextTokenFromRequest(Request);

            if (contextTokenString != null)
            {
                spModel.ContextToken = TokenHelper.ReadAndValidateContextToken(contextTokenString, Request.Url.Authority);

                spModel.SharepointUrl = new Uri(Request.QueryString["SPHostUrl"]);
                spModel.AccessToken = TokenHelper.GetAccessToken(spModel.ContextToken, spModel.SharepointUrl.Authority).AccessToken;

                // For simplicity, this sample assigns the access token to the button's CommandArgument property. 
                // In a production add-in, this would not be secure. The access token should be cached on the server-side.
                // CSOM.CommandArgument = accessToken; - change this to an action link Post
            }


            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            using (spModel.SPClientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (spModel.SPClientContext != null)
                {
                    //spModel.SPClientContext.

                    spUser = spModel.SPClientContext.Web.CurrentUser;
                    spModel.SPClientContext.Load(spUser, user => user.Title);
                    spModel.SPClientContext.ExecuteQuery();
                    ViewBag.UserName = spUser.Title;
                }
            }

            RetrieveWithCSOM(spModel);

            return View(spModel);           
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
    }
}

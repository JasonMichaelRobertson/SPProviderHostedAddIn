using System;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;

namespace SPProviderHostedAddInWeb.Models
{
    public class SPModel
    {
        private SharePointContextToken contextToken;
        private string accessToken;
        private Uri sharepointUrl;
        private string siteName;
        private string currentUser;
        private List<string> listOfUsers = new List<string>();
        private List<string> listOfLists = new List<string>();

        private ClientContext spClientContext;

        public SPModel()
        {

        }

        public Uri SharepointUrl
        {
            get
            {
                return sharepointUrl;
            }

            set
            {
                sharepointUrl = value;
            }
        }

        public SharePointContextToken ContextToken
        {
            get
            {
                return contextToken;
            }

            set
            {
                contextToken = value;
            }
        }

        public string AccessToken
        {
            get
            {
                return accessToken;
            }

            set
            {
                accessToken = value;
            }
        }

        public string SiteName
        {
            get
            {
                return siteName;
            }

            set
            {
                siteName = value;
            }
        }

        public string CurrentUser
        {
            get
            {
                return currentUser;
            }

            set
            {
                currentUser = value;
            }
        }

        public List<string> ListOfUsers
        {
            get
            {
                return listOfUsers;
            }

            set
            {
                listOfUsers = value;
            }
        }

        public List<string> ListOfLists
        {
            get
            {
                return listOfLists;
            }

            set
            {
                listOfLists = value;
            }
        }

        public ClientContext SPClientContext
        {
            get
            {
                return spClientContext;
            }

            set
            {
                spClientContext = value;
            }
        }
    }
}
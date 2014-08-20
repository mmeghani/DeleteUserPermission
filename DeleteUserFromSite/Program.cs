using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using Microsoft.SharePoint;

namespace DeleteUserFromSite
{
    public class Program
    {
        private static string userName = string.Empty;
        private static string spacer = string.Empty;

        static void Main(string[] args)
        {
            if (ConfigurationManager.AppSettings["UserName"] != null && !string.IsNullOrEmpty(ConfigurationManager.AppSettings["UserName"].ToString()))
            {
                userName = ConfigurationManager.AppSettings["UserName"].ToString();
                RemoveUserFromSite();
            }
            else
            {
                WriteToText("Username is missing in the config file.");
                Console.WriteLine("Username is missing in the config file.");
            }
        }

        private static void RemoveUserFromSite()
        {
            using(SPSite site = new SPSite(ConfigurationManager.AppSettings["SiteURL"].ToString()))
            {
                using (SPWeb web = site.OpenWeb())
                {
                    RemoveUserFromSite(web);
                    //SPWebCollection colWeb = web.Webs;

                    //foreach (SPWeb subsite in colWeb)
                    //{
                    //    WriteToText(" Site: " + subsite.Title);
                    //    RemoveUserFromSP(subsite.Url);
                    //}

                }
            }
        }

        private static void RemoveUserFromSite(SPWeb web)
        {
            WriteToText("Site: " + web.Title + " - URL: " + web.Url);
            spacer += "|   ";
            if (web.Webs.Count > 0)
            {
                foreach (SPWeb subweb in web.Webs)
                {
                    RemoveUserFromSite(subweb);
                }
                if (!string.IsNullOrEmpty(spacer) && spacer.Length >= 4)
                    spacer = spacer.Substring(4);
                RemoveUserFromSP(web.Url);
            }
            else
            {
                if (!string.IsNullOrEmpty(spacer) && spacer.Length >= 4)
                    spacer = spacer.Substring(4);
                RemoveUserFromSP(web.Url);
            }
        }

        public static void RemoveUserFromSP(string siteURL)
        {
            try
            {
                using (SPSite site = new SPSite(siteURL))
                {
                    using (SPWeb web = site.OpenWeb())
                    {
                        SPUserCollection userCollection = web.SiteUsers;

                        string domain = ConfigurationManager.AppSettings["DOMAIN"].ToString();
                        string claims = ConfigurationManager.AppSettings["CLAIMS"].ToString();

                        SPUser spUser = null;
                        bool userFound = false;
                        try
                        {
                            spUser = web.AllUsers[claims + "\\" + userName];
                            userFound = true;
                        }
                        catch
                        {
                            try
                            {
                                spUser = web.AllUsers[domain + "\\" + userName];
                                userFound = true;
                            }
                            catch
                            {
                                userFound = false;
                                WriteToText("|--  User NOT found ~~~~ " + userName + " Site: " + web.Title);
                            }
                        }

                        if (userFound)
                        {
                            if (web.HasUniqueRoleAssignments)
                                RemoveUserRoles(web, spUser);
                            else
                                WriteToText("|--  NO Unique Permission for site - " + web.Title);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToText("**** ERROR **** : " + ex.Message);
            }
        }

        private static void RemoveUserRoles(SPWeb web, SPUser user)
        {
            bool found = false;
            SPRoleAssignmentCollection SPRoleAssColn = web.RoleAssignments;

            for (int i = SPRoleAssColn.Count - 1; i >= 0; i--)
            {

                SPRoleAssignment roleAssignmentSingle = SPRoleAssColn[i];

                SPPrincipal wUser = (SPPrincipal)user;

                if (roleAssignmentSingle.Member.ID == wUser.ID)
                {
                    found = true;
                    SPRoleAssColn.Remove(i);
                    WriteToText("|--  Permission removed for: " + user.Name + " - from - " + web.Title); 
                }
            }

            if(!found)
                WriteToText("|--  NO permission for user: " + user.Name + " - for - " + web.Title);
        }

        private static void WriteToText(string message)
        {
            using (StreamWriter file = new StreamWriter(@"RemoveUser_" + userName + ".log", true))
            {
                        file.WriteLine(spacer + message);
            }
        }
    }
}

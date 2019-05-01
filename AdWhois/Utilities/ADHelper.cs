using System;
using System.Collections;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

//Use http://social.technet.microsoft.com/wiki/contents/articles/12037.active-directory-get-aduser-default-and-extended-properties.aspx

namespace DSHS.ESA.DCS
{
	public class ADHelper
	{
        /// <summary>
        /// The adGroup parameter value is the AD group display name.
        /// </summary>
        /// <param name="adGroup">Default Value is "G-S-DSHS DCS IT Service Desk"</param>
        /// <returns></returns>
        public static List<string> GetGroupMembers(string adGroup)
        {
            List<string> members = new List<string>();

            using (PrincipalContext ctx = new PrincipalContext(ContextType.Domain))
            {
                try
                {
                    // find the group in question
                    GroupPrincipal group = GroupPrincipal.FindByIdentity(ctx, adGroup);
                    // if found....
                    if (group != null)
                    {
                        // iterate over the group's members
                        foreach (Principal p in group.GetMembers())
                        {
                            members.Add(p.DisplayName);
                        }
                    }
                }
                catch (Exception ex) { /* Just let it fall through.*/ }
            }

            return members;
        }

		/// <summary>
		/// Returns user name or AD Login ID.
		/// </summary>
		/// <param name="path"></param>
		/// <returns></returns>
		public static string ExtractUserName(string username)
		{
			if (string.IsNullOrEmpty(username)) return username;

			string[] userPath = username.Split(new char[] { '\\' });
			return userPath[userPath.Length - 1];
		}

		/// <summary>
		/// Checks to see if user exists in AD.
		/// </summary>
		/// <param name="loginName"></param>
		/// <returns></returns>
		public static bool IsExistInAD(string loginName)
		{
			SearchResult result = ADSearchResult(loginName);
			
			return result != null;
			
		}

		/// <summary>
		/// Gets a list AD groups that a user associates with.
		/// </summary>
		/// <param name="userName"></param>
		/// <returns></returns>
		public static List<string> GetADUserGroups(string userName)
		{
			DirectorySearcher search = ADDirectorySearcher;
			search.Filter = String.Format("(cn={0})", userName);
			search.PropertiesToLoad.Add("memberOf");
			List<string> groupsList = new List<string>();

			SearchResult result = search.FindOne();
			if (result != null)
			{
				int groupCount = result.Properties["memberOf"].Count;

				for (int counter = 0; counter < groupCount; counter++)
				{
					groupsList.Add((string)result.Properties["memberOf"][counter]);
				}
			}
			return groupsList;
		}

		/// <summary>
		/// Gets a list of users from an AD group.
		/// </summary>
		/// <param name="groupName"></param>
		/// <returns></returns>
		public static List<string> GetADGroupUsers(string groupName)
		{
			SearchResult result;
			DirectorySearcher search = ADDirectorySearcher;
			search.Filter = String.Format("(cn={0})", groupName);
			search.PropertiesToLoad.Add("member");
			result = search.FindOne();

			List<string> userNames = new List<string>();
			if (result != null)
			{
				for (int counter = 0; counter <
						 result.Properties["member"].Count; counter++)
				{
					string user = (string)result.Properties["member"][counter];
					userNames.Add(user);
				}
			}
			return userNames;
		}

        /// <summary>
        /// Returns list of possible matching AD displaynames.
        /// </summary>
        /// <param name="prefixText">search string entered</param>
        /// <returns></returns>
        public static List<string> SearchADDisplayname(string prefixText)
        {
            //Without this, LDAP query does not match with AD displayNames when capital letters are present in prefixText.
            prefixText = prefixText.ToLower();

            DirectoryEntry directory = ADDirectoryEntry;//new DirectoryEntry("LDAP://DC=dshs,DC=wa,DC=lcl");

            //Note that the LDAP displayName attribute is typically capitalized. However, the below LDAP query will not work unless all letters are lowercase (reason unknown). userAccountControl:1.2.840.113556.1.4.803:=2 limits query results to active users. mail=* limits query results to records that have data in the mail (email) attribute.
            string filter = "(&(displayname=" + prefixText + "*)(objectClass=user)(objectCategory=Person)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(mail=*)(|(businessCategory=DSHS)(businessCategory=Intern)(businessCategory=Contractor)))";

            string[] strCats = { "displayname" };
            List<string> items = new List<string>();
            DirectorySearcher dirComp = new DirectorySearcher(directory, filter, strCats, SearchScope.Subtree);

            SearchResultCollection results = dirComp.FindAll();

            foreach (SearchResult result in results)
            {
                foreach (DictionaryEntry prop in result.Properties)
                {
                    if (prop.Key.Equals("displayname"))
                    {
                        System.Collections.IEnumerable propsEnum = prop.Value as System.Collections.IEnumerable;
                        foreach (object individualValue in propsEnum)
                        {
                            if (individualValue.ToString().IndexOf(prefixText) != 0)
                            {
                                items.Add(individualValue.ToString());


                            }
                        }
                    }
                }
            }

            return items;
        }

		/// <summary>
		/// Returns display name. Ex John Doe (DSHS/TSD).
		/// </summary>
		/// <param name="loginName"></param>
		/// <returns></returns>
		public static string DisplayName(string loginName)
		{
			SearchResult result = ADSearchResult(loginName);


			if (result.Properties.Contains("displayName"))
			{
				return result.Properties["displayName"][0].ToString();
			}
			else
			{
				return string.Empty;
			}
		}
        /// <summary>
        /// Returns display name. Ex John Doe (DSHS/TSD).
        /// </summary>
        /// <param name="loginName"></param>
        /// <returns></returns>
        public static string LoginName(string displayName)
        {
            SearchResult result = ADSearchResult(displayName);


            if (result.Properties.Contains("SAMAccountName"))
            {
                return result.Properties["SAMAccountName"][0].ToString();
            }
            else
            {
                return string.Empty;
            }
        }

		/// <summary>
		/// Returns email address.
		/// </summary>
		/// <param name="loginName"></param>
		/// <returns></returns>
		public static string Email(string loginName)
		{
			SearchResult result = ADSearchResult(loginName);


			if (result.Properties.Contains("mail"))
			{
				return result.Properties["mail"][0].ToString();
			}
			else
			{
				return string.Empty;
			}
		}

		/// <summary>
		/// Returns first name.
		/// </summary>
		/// <param name="loginName"></param>
		/// <returns></returns>
		public static string FirstName(string loginName)
		{
			SearchResult result = ADSearchResult(loginName);


			if (result.Properties.Contains("givenName"))
			{
				return result.Properties["givenName"][0].ToString();
			}
			else
			{
				return string.Empty;
			}
		}

		/// <summary>
		/// Returns middle initial.
		/// </summary>
		/// <param name="loginName"></param>
		/// <returns></returns>
		public static string MidleInitial(string loginName)
		{
			SearchResult result = ADSearchResult(loginName);


			if (result.Properties.Contains("Initials"))
			{
				return result.Properties["Initials"][0].ToString();
			}
			else
			{
				return string.Empty;
			}
		}

		/// <summary>
		/// Returns last name.
		/// </summary>
		/// <param name="loginName"></param>
		/// <returns></returns>
		public static string LastName(string loginName)
		{
			SearchResult result = ADSearchResult(loginName);


			if (result.Properties.Contains("sn"))
			{
				return result.Properties["sn"][0].ToString();
			}
			else
			{
				return string.Empty;
			}
		}

		/// <summary>
		/// Return phone number.
		/// </summary>
		/// <param name="loginName"></param>
		/// <returns></returns>
		public static string OfficePhone(string loginName)
		{
			SearchResult result = ADSearchResult(loginName);


			if (result.Properties.Contains("OfficePhone"))
			{
				return result.Properties["OfficePhone"][0].ToString();
			}
			else
			{
				return string.Empty;
			}
		}

		/// <summary>
		/// Returns AD GUID. It's a unique identifier for each person.
		/// </summary>
		/// <param name="loginName"></param>
		/// <returns></returns>
		public static Guid ADGuid(string loginName)
		{
			SearchResult result = ADSearchResult(loginName);


			if (result.Properties.Contains("ObjectGUID"))
			{
				return result.Properties["ObjectGUID"][0].GetType().GUID;
			}
			else
			{
				return default(Guid);
			}
		}

		/// <summary>
		/// Returns a GUID in a format of 32 digits separated by hyphens: 00000000-0000-0000-0000-000000000000
		/// </summary>
		/// <param name="loginName"></param>
		/// <returns></returns>
		public static string GuidToString(string loginName)
		{
			Guid guid = ADGuid(loginName);
			if (guid != null) return guid.ToString("D");
			else return string.Empty;
		}

		/// <summary>
		/// Removes user from a specified AD group.  Must have privilidge to add.
		/// </summary>
		/// <param name="loginName">AD user to add</param>
		/// <param name="groupDn">AD group to add</param>
		/// <remarks>https://msdn.microsoft.com/en-us/library/ms676310(v=VS.85).aspx</remarks>
		/// <returns></returns>
		public static bool RemoveUserFromGroup(string loginName, string groupDn)
		{
			try
			{

				string userName = ExtractUserName(loginName);
				DirectoryEntry entryDomain = GetDirectoryEntry(groupDn);

				//find user
				SearchResult findUser = ADSearchResult(loginName);

				if (entryDomain != null && findUser.Properties.Contains("DistinguishedName"))
				{
					entryDomain.Properties["member"].Remove(findUser.Properties["DistinguishedName"][0].ToString());
					entryDomain.CommitChanges();
					entryDomain.Close();
					return true;
				}
				else
				{
					return false;
				}

			}
			catch (System.DirectoryServices.DirectoryServicesCOMException)
			{
				throw;
			}
		}

		/// <summary>
		/// Adds user to a specified AD group.  Must have privilidge to add.
		/// </summary>
		/// <param name="loginName">AD user to add</param>
		/// <param name="groupDn">AD group to add</param>
		/// <remarks>https://msdn.microsoft.com/en-us/library/ms676310(v=VS.85).aspx</remarks>
		/// <returns></returns>
		public static bool AddUserToGroup(string loginName, string groupDn)
		{
			try
			{
				string userName = ExtractUserName(loginName);
				DirectoryEntry entryDomain = GetDirectoryEntry(groupDn);

				//find user
				SearchResult findUser = ADSearchResult(loginName);

				if (entryDomain != null && findUser.Properties.Contains("DistinguishedName"))
				{
					entryDomain.Properties["member"].Add(findUser.Properties["DistinguishedName"][0].ToString());
					entryDomain.CommitChanges();
					entryDomain.Close();
					return true;
				}
				else
				{
					return false;
				}


			}
			catch (System.DirectoryServices.DirectoryServicesCOMException)
			{
				throw;
			}
		}

		private static DirectoryEntry GetDirectoryEntry(string groupDn)
		{
			string group = ExtractUserName(groupDn);
			DirectoryEntry entryDomain = null; // new DirectoryEntry("LDAP://CN=G-S-DSHS ISSD BOA PRISM Administrators,OU=App,OU=Resource Groups,OU=Groups,OU=EXEC IT,OU=SEC,DC=dshs,DC=wa,DC=lcl");

			//CN=G-S-DSHS ISSD BOA PRISM Administrators,OU=App,OU=Resource Groups,OU=Groups,OU=EXEC IT,OU=SEC,DC=dshs,DC=wa,DC=lcl			

			//find group
			SearchResult findGroup = ADSearchResult(groupDn);

			if (findGroup.Properties.Contains("DistinguishedName"))
			{
				entryDomain = new DirectoryEntry("LDAP://" + findGroup.Properties["DistinguishedName"][0].ToString());
			}

			return entryDomain;
		}
		/// <summary>
		/// Returns Search Result of specific properties.
		/// </summary>
		/// <param name="loginName"></param>
		/// <returns></returns>
		private static SearchResult ADSearchResult(string loginName)
		{
			string userName = ExtractUserName(loginName);
			DirectorySearcher search = ADDirectorySearcher;
            search.Filter = "(|" + String.Format("(SAMAccountName={0})", userName) + String.Format("(displayName={0})", loginName) + ")";
			search.PropertiesToLoad.Add("cn");
			search.PropertiesToLoad.Add("displayName");
			search.PropertiesToLoad.Add("mail");
			search.PropertiesToLoad.Add("givenName");
			search.PropertiesToLoad.Add("sn");
			search.PropertiesToLoad.Add("Initials");
			search.PropertiesToLoad.Add("telephoneNumber");
			search.PropertiesToLoad.Add("ObjectGUID");
			search.PropertiesToLoad.Add("Name");
			search.PropertiesToLoad.Add("DistinguishedName");
            search.PropertiesToLoad.Add("SAMAccountName");
			SearchResult result = search.FindOne();

			return result;
		}

		/// <summary>
		/// Gets Directory Searcher for a domain.
		/// </summary>
		private static DirectorySearcher ADDirectorySearcher
		{
			get
			{
				DirectoryEntry entryDomain = ADDirectoryEntry;
				DirectorySearcher search = new DirectorySearcher(entryDomain);

				return search;

			}
		}

		/// <summary>
		/// Gets Directory Entry using root and domain.
		/// </summary>
		private static DirectoryEntry ADDirectoryEntry
		{
			get
			{
				DirectoryEntry entryRoot = new DirectoryEntry("LDAP://RootDSE");
				string domain = entryRoot.Properties["defaultNamingContext"][0].ToString();
				DirectoryEntry entryDomain = new DirectoryEntry("LDAP://" + domain);

				return entryDomain;

			}
		}

		/// <summary>
		/// This take a window logon ID and verified it against the provided AD Group. 
		/// It will return a true if the user logon ID is in the AD Group, otherwise return false. 
		/// </summary>
		/// <param name="userLogon"></param>
		/// <param name="adGroupName"></param>
		/// <returns>bool</returns>
		public static bool IsUserExistInADGroup(string userLogon, string adGroupName)
		{
			try
			{
				// Check to see if the current window user is in the AppDevGroup
				// This will verify that the user does have access to the current Application.
				List<string> adGroupUsr = ADHelper.GetADGroupUsers(adGroupName);
				bool isMember = adGroupUsr.Exists(x => x.Contains(ADHelper.LastName(userLogon)) &&
														 x.Contains(ADHelper.FirstName(userLogon)));

				return isMember;
			}
			catch (System.DirectoryServices.DirectoryServicesCOMException)
			{
				throw;
			}
		}
	}
}

using System;
using System.Configuration;
using System.Collections.Generic;
using System.Web.Http;
using DSHS.ESA.DCS;

namespace AdWhois.Controllers
{
    [RoutePrefix("api/whois")]
    public class AdWhoisController : ApiController
    {
        #region Version 1

        // GET api/whois/v1/GetGroupMembers
        [HttpGet]
        [Route("v1/GetGroupMembers")]
        public IEnumerable<string> GetGroupMembers()
        {
            return ADHelper.GetGroupMembers(ConfigurationManager.AppSettings["DefaultAdGroupName"]);
        }

        // Post api/whois/v1/GetAdGroupUsers
        [HttpPost]
        [Route("v1/GetAdGroupUsers")]
        public IEnumerable<string> GetAdGroupUsers([FromBody]string adGroup)
        {
            return ADHelper.GetADGroupUsers(adGroup);
        }

        #endregion

        #region Version 2

        // Post api/whois/v2/GetGroupMembers
        [HttpPost]
        [Route("v2/GetGroupMembers")]
        public IEnumerable<string> GetGroupMembers([FromBody]string adGroup)
        {
            return ADHelper.GetGroupMembers(adGroup);
        }

        #endregion
    }
}

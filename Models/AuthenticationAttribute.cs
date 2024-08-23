using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Web;
using System.Web.Http.Controllers;

namespace CORSYS_API.Models
{
    public class AuthenticationAttribute : System.Web.Http.Filters.AuthorizationFilterAttribute
    {
        private const string Realm = "My Realm";
        public override void OnAuthorization(HttpActionContext actionContext)
        {
            if (actionContext.Request.Headers.Authorization == null)
            {
                actionContext.Response = actionContext.Request.CreateResponse(HttpStatusCode.Unauthorized);
                if (actionContext.Response.StatusCode == HttpStatusCode.Unauthorized)
                {
                    actionContext.Response.Headers.Add("WWW-Authenticate",
                        string.Format("Basic realm=\"{0}\"", Realm));
                }
            }
            else
            {
                //Get the authentication token from the request header
                string authenticationToken = actionContext.Request.Headers.Authorization.Parameter;

                //Decode the string
                string decodedAuthenticationToken;
                try
                {
                    decodedAuthenticationToken = Encoding.UTF8.GetString(
                    Convert.FromBase64String(authenticationToken));

                }
                catch (Exception)
                {
                    decodedAuthenticationToken = null;
                    //throw;
                }

                //call the login method to check the username and password
                if (decodedAuthenticationToken != "hysdroexpandtestscada:cladtekbatam2020")
                {
                    actionContext.Response = actionContext.Request
                        .CreateResponse(HttpStatusCode.Unauthorized, "Unauthorized");
                }
            }
        }
    }

}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace NuruWeb.Account
{
    public partial class Register : System.Web.UI.Page
    {

        protected void Page_Load(object sender, EventArgs e)
        {
            //RegisterUser_ContinueDestinationPageUrl = Request.QueryString["ReturnUrl"];
        }

        protected void RegisterUser_CreatedUser(object sender, EventArgs e)
        {
            //FormsAuthentication.SetAuthCookie(RegisterUser.UserName, false /* createPersistentCookie */);

            //string continueUrl = RegisterUser.ContinueDestinationPageUrl;
            //if (String.IsNullOrEmpty(continueUrl))
            //{
            //    continueUrl = "~/";
            //}

            //MyShoppingCart usersShoppingCart = new MyShoppingCart();
            //String cartId = usersShoppingCart.GetShoppingCartId();
            //usersShoppingCart.MigrateCart(cartId, RegisterUser.UserName);

            //Response.Redirect(continueUrl);
        }

        protected void LoginButton_Click(object sender, ImageClickEventArgs e)
        {

        }

        protected void RegisterUser_CreatedUserFail(object sender, CreateUserErrorEventArgs e)
        {

        }

        protected void RegisterUser_Continue(object sender, EventArgs e)
        {
  //          Response.Redirect(continueUrl);
        }

        protected void RegisterUser_Cancel(object sender, EventArgs e)
        {

        }

        protected void RegisterUser_Error(object sender, CreateUserErrorEventArgs e)
        {

        }

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Nuru_Web.Styles
{
    public partial class NuruWeb : System.Web.UI.MasterPage
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (HttpContext.Current.User.Identity.IsAuthenticated)
            {
                if (Session["UserName"] == null)
                {
                    Session["UserName"] = HttpContext.Current.User.Identity.Name;
                }
            }
        }

        protected void btnShowPopup_Click(object sender, EventArgs e)
        {

        }


    }
}

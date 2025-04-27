using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI.WebControls;
using System.Web.UI;
using AjaxControlToolkit;

namespace Nuru_Web
{
    public class ClientHelper
    {
        public static void ClientMessage(Page p, string message)
        {
            if (p.Master == null)
            {
                p.ClientScript.RegisterClientScriptBlock(p.GetType(), System.Guid.NewGuid().ToString(),
                string.Format("<script>alert('{0}');</script>", message));
            }
            else
            {
                Label _lblClientMessage_ = p.Master.FindControl("_lblClientMessage_") as Label;
                _lblClientMessage_.Text = message;
                UpdatePanel upd = p.Master.FindControl("_updClientMessage_") as UpdatePanel;
                upd.Update();
                ModalPopupExtender extender = p.Master.FindControl("mdlPopup") as ModalPopupExtender;
                extender.Show();
            }
        }
   }

}

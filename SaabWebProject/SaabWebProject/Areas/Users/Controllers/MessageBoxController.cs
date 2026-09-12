using SaabWebProject.Models.Extensions.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace SaabWebProject.Areas.Users.Controllers
{
    public class MessageBoxController : Controller
    {
        public class MessageBox
        {
           
            public static JavaScriptResult Show(string message, MessageType type = MessageType.Alert, bool modal = false, MessageAlignment layout = MessageAlignment.Center, bool dismissQueue = false)
            {
                string txt = "$.noty.closeAll(); noty({ text: \"" + message + "\", type: \"" + type.ToString().ToLower() + "\", layout: \"" + layout.ToString().ToLowerFirst() + "\", dismissQueue: " + dismissQueue.ToString().ToLower() + ", modal: " + modal.ToString().ToLower() + " });";
                return new JavaScriptResult() { Script = txt };
            }


            public static JavaScriptResult CustomizeShow(string message, MessageType type = MessageType.Alert, bool modal = false, MessageAlignment layout = MessageAlignment.Center, bool dismissQueue = false)
            {
                string txt = "$.noty.closeAll(); noty({ text: \"" + message + "\", type: \"" + type.ToString().ToLower() + "\", layout: \"" + layout.ToString().ToLowerFirst() + "\", dismissQueue: " + dismissQueue.ToString().ToLower() + ", modal: " + modal.ToString().ToLower() + " ,callback:{afterClose: function () {$('#' + id).parent().remove();}} });";
                return new JavaScriptResult() { Script = txt };
            }

            public static JavaScriptResult RefreshAfterShow(string message, MessageType type = MessageType.Alert, bool modal = false, MessageAlignment layout = MessageAlignment.Center, bool dismissQueue = false)
            {
                string txt = "$.noty.closeAll(); noty({ text: \"" + message + "\", type: \"" + type.ToString().ToLower() + "\", layout: \"" + layout.ToString().ToLowerFirst() + "\", dismissQueue: " + dismissQueue.ToString().ToLower() + ", modal: " + modal.ToString().ToLower() + " ,callback:{afterClose: location.reload();} });";
                return new JavaScriptResult() { Script = txt };
            }


        }
        public enum MessageType
        {
            Success,
            Error,
            Information,
            Warning,
            Alert,
            Notification
        }
        public enum MessageAlignment
        {
            Bottom,
            BottomCenter,
            BottomLeft,
            BottomRight,
            Center,
            CenterLeft,
            CenterRight,
            Inline,
            Top,
            TopCenter,
            TopLeft,
            TopRight
        }
    }
}
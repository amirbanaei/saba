using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WSDLSample.PanelSMS
{
    using System;
    using System.Web.Services;
    using System.Diagnostics;
    using System.Web.Services.Protocols;
    using System.Xml.Serialization;
    using System.ComponentModel;


    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.3190.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Web.Services.WebServiceBindingAttribute(Name = "smsserverBinding", Namespace = "urn:smsserver")]
    [System.Xml.Serialization.SoapIncludeAttribute(typeof(input_data_type))]
    public partial class smsserver : System.Web.Services.Protocols.SoapHttpClientProtocol
    {

        private System.Threading.SendOrPostCallback sendPatternSmsOperationCompleted;

        private bool useDefaultCredentialsSetExplicitly;

        /// <remarks/>
        public smsserver()
        {
            this.Url = global::WSDLSample.Properties.Settings.Default.WSDLSample_PanelSMS_smsserver;
            if ((this.IsLocalFileSystemWebService(this.Url) == true))
            {
                this.UseDefaultCredentials = true;
                this.useDefaultCredentialsSetExplicitly = false;
            }
            else
            {
                this.useDefaultCredentialsSetExplicitly = true;
            }
        }

        public new string Url
        {
            get
            {
                return base.Url;
            }
            set
            {
                if ((((this.IsLocalFileSystemWebService(base.Url) == true)
                            && (this.useDefaultCredentialsSetExplicitly == false))
                            && (this.IsLocalFileSystemWebService(value) == false)))
                {
                    base.UseDefaultCredentials = false;
                }
                base.Url = value;
            }
        }

        public new bool UseDefaultCredentials
        {
            get
            {
                return base.UseDefaultCredentials;
            }
            set
            {
                base.UseDefaultCredentials = value;
                this.useDefaultCredentialsSetExplicitly = true;
            }
        }

        /// <remarks/>
        public event sendPatternSmsCompletedEventHandler sendPatternSmsCompleted;

        /// <remarks/>
        [System.Web.Services.Protocols.SoapRpcMethodAttribute("sendPatternSms", RequestNamespace = "188.0.240.110", ResponseNamespace = "188.0.240.110")]
        [return: System.Xml.Serialization.SoapElementAttribute("return")]
        public string sendPatternSms(string fromNum, string[] toNum, string user, string pass, string pattern_code, input_data_type[] input_data)
        {
            object[] results = this.Invoke("sendPatternSms", new object[] {
                        fromNum,
                        toNum,
                        user,
                        pass,
                        pattern_code,
                        input_data});
            return ((string)(results[0]));
        }

        /// <remarks/>
        public void sendPatternSmsAsync(string fromNum, string[] toNum, string user, string pass, string pattern_code, input_data_type[] input_data)
        {
            this.sendPatternSmsAsync(fromNum, toNum, user, pass, pattern_code, input_data, null);
        }

        /// <remarks/>
        public void sendPatternSmsAsync(string fromNum, string[] toNum, string user, string pass, string pattern_code, input_data_type[] input_data, object userState)
        {
            if ((this.sendPatternSmsOperationCompleted == null))
            {
                this.sendPatternSmsOperationCompleted = new System.Threading.SendOrPostCallback(this.OnsendPatternSmsOperationCompleted);
            }
            this.InvokeAsync("sendPatternSms", new object[] {
                        fromNum,
                        toNum,
                        user,
                        pass,
                        pattern_code,
                        input_data}, this.sendPatternSmsOperationCompleted, userState);
        }

        private void OnsendPatternSmsOperationCompleted(object arg)
        {
            if ((this.sendPatternSmsCompleted != null))
            {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.sendPatternSmsCompleted(this, new sendPatternSmsCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }

        /// <remarks/>
        public new void CancelAsync(object userState)
        {
            base.CancelAsync(userState);
        }

        private bool IsLocalFileSystemWebService(string url)
        {
            if (((url == null)
                        || (url == string.Empty)))
            {
                return false;
            }
            System.Uri wsUri = new System.Uri(url);
            if (((wsUri.Port >= 1024)
                        && (string.Compare(wsUri.Host, "localHost", System.StringComparison.OrdinalIgnoreCase) == 0)))
            {
                return true;
            }
            return false;
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.7.3190.0")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.SoapTypeAttribute(Namespace = "urn:smsserver")]
    public partial class input_data_type
    {

        private string keyField;

        private string valueField;

        /// <remarks/>
        [System.Xml.Serialization.SoapElementAttribute(IsNullable = true)]
        public string key
        {
            get
            {
                return this.keyField;
            }
            set
            {
                this.keyField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.SoapElementAttribute(IsNullable = true)]
        public string value
        {
            get
            {
                return this.valueField;
            }
            set
            {
                this.valueField = value;
            }
        }
    }

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.3190.0")]
    public delegate void sendPatternSmsCompletedEventHandler(object sender, sendPatternSmsCompletedEventArgs e);

    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.3190.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class sendPatternSmsCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs
    {

        private object[] results;

        internal sendPatternSmsCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) :
                base(exception, cancelled, userState)
        {
            this.results = results;
        }

        /// <remarks/>
        public string Result
        {
            get
            {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
    }
}

namespace WSDLSample.Properties
{


    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "16.0.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase
    {

        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));

        public static Settings Default
        {
            get
            {
                return defaultInstance;
            }
        }

        [global::System.Configuration.ApplicationScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.SpecialSettingAttribute(global::System.Configuration.SpecialSetting.WebServiceUrl)]
        [global::System.Configuration.DefaultSettingValueAttribute("http://188.0.240.110/public/class/sms/csharpservice_new/server.php")]
        public string WSDLSample_PanelSMS_smsserver
        {
            get
            {
                return ((string)(this["WSDLSample_PanelSMS_smsserver"]));
            }
        }
    }
}
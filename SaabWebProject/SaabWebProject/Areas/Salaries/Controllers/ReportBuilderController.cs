using System;
using System.Data;
using System.Web.Mvc;
using SaabWebProject.Utility;
using Stimulsoft.Report;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Mvc;

namespace SaabWebProject.Areas.Salaries.Controllers
{
    public class ReportBuilderController : Controller
    {
        // GET
        [AuthorizeAAA]
        public ActionResult Index()
        {

            return View();
        }
        [AuthorizeAAA]
        public ActionResult GetReport()
        {
            var report = new StiReport();
            report.Load(Server.MapPath("~/Assets/Report_File/Fish_Hoghghi.mrt"));
            return StiMvcDesigner.GetReportResult(report);
        }
         
        [AuthorizeAAA]
        public ActionResult DesignerEvent()
        {
            StiReport report = StiMvcDesigner.GetActionReportObject();
            return StiMvcDesigner.DesignerEventResult(report);

        }

        [AuthorizeAAA]
        public ActionResult ViewerEvent()

        {

            return StiMvcViewer.ViewerEventResult();

        }

        [AuthorizeAAA]
        public ActionResult Save()
        {
            StiReport report = StiMvcDesigner.GetActionReportObject();
            report.Save(Server.MapPath("~/Assets/Report_File/Fish_Hoghghi.mrt"));
            return StiMvcDesigner.SaveReportResult(report);
        }

        [AuthorizeAAA]
        public ActionResult CreateReport()
        {
            StiReport report = StiReport.CreateNewReport();
            return StiMvcDesigner.GetReportResult(report);
        }

        [AuthorizeAAA]
        public ActionResult PreviewReport()
        {
            StiReport report = StiMvcDesigner.GetActionReportObject();

            return StiMvcDesigner.PreviewReportResult(report);
        }

        [AuthorizeAAA]
        public ActionResult GetPreviewReport()
        {
            StiReport report = StiMvcDesigner.GetActionReportObject();
            return StiMvcDesigner.PreviewReportResult(report);
        }


    }
}
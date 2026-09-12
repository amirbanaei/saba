using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Web.Mvc;

namespace SaabWebProject.Areas.Setting.Controllers
{
    public class seegozareshController : Controller
    {
        private readonly HttpClient _httpClient = new HttpClient();
        //http://localhost:6061/ManagementAccounting/FinancialDocuments/dastehbandi_coding
        //مختصات GIS
        //http://localhost:6061/Setting/Orders/view_xandy
        public ActionResult Index()
        {
            return View();
        }

        // GET: Setting/seegozaresh/NearestPoints
        public async Task<ActionResult> NearestPoints()
        {
            var points = GetPoints();
            var results = new List<object>();

            foreach (var p in points)
            {
                var others = points.Where(x => x.Id != p.Id).ToList();
                double? minDist = null;
                double? maxDist = null;
                NearestPoint nearest = null;
                NearestPoint farthest = null;

                foreach (var q in others)
                {
                    var route = await GetRouteDistance(p, q);
                    if (route == null) continue;

                    if (minDist == null || route.Distance < minDist)
                    {
                        minDist = route.Distance;
                        nearest = new NearestPoint
                        {
                            To = q.Id,
                            Distance = route.Distance,
                            Geometry = route.Geometry,
                            Address = await GetAddress(q.Lat, q.Lng)
                        };
                    }

                    if (maxDist == null || route.Distance > maxDist)
                    {
                        maxDist = route.Distance;
                        farthest = new NearestPoint
                        {
                            To = q.Id,
                            Distance = route.Distance,
                            Geometry = route.Geometry,
                            Address = await GetAddress(q.Lat, q.Lng)
                        };
                    }
                }

                results.Add(new
                {
                    From = p.Id,
                    NearestTo = nearest?.To,
                    NearestDistance = nearest?.Distance,
                    NearestAddress = nearest?.Address,
                    FarthestTo = farthest?.To,
                    FarthestDistance = farthest?.Distance,
                    FarthestAddress = farthest?.Address
                });
            }

            return Json(results, JsonRequestBehavior.AllowGet);
        }

        private async Task<RouteResult> GetRouteDistance(Point p1, Point p2)
        {
            var url = $"https://router.project-osrm.org/route/v1/driving/{p1.Lng},{p1.Lat};{p2.Lng},{p2.Lat}?overview=full&geometries=geojson";
            var res = await _httpClient.GetAsync(url);
            if (!res.IsSuccessStatusCode) return null;

            var json = await res.Content.ReadAsStringAsync();
            var serializer = new JavaScriptSerializer();
            var data = serializer.Deserialize<OsrmResponse>(json);
            if (data?.Routes == null || data.Routes.Length == 0) return null;

            return new RouteResult
            {
                Distance = data.Routes[0].Distance,
                Geometry = data.Routes[0].Geometry,
                Lat = p2.Lat,
                Lng = p2.Lng
            };
        }

        private async Task<string> GetAddress(double lat, double lng)
        {
            try
            {
                var url = $"https://nominatim.openstreetmap.org/reverse?format=jsonv2&lat={lat}&lon={lng}";
                var res = await _httpClient.GetAsync(url);
                var json = await res.Content.ReadAsStringAsync();
                var serializer = new JavaScriptSerializer();
                dynamic data = serializer.Deserialize<dynamic>(json);
                return data.ContainsKey("display_name") ? data["display_name"] : "نامعلوم";
            }
            catch { return "نامعلوم"; }
        }

        // ---------------- کلاس‌ها ----------------
        public class Point
        {
            public int Id { get; set; }
            public double Lat { get; set; }
            public double Lng { get; set; }
        }

        public class NearestPoint
        {
            public int To { get; set; }
            public double Distance { get; set; }
            public object Geometry { get; set; }
            public string Address { get; set; }
        }

        public class RouteResult
        {
            public double Distance { get; set; }
            public object Geometry { get; set; }
            public double Lat { get; set; }
            public double Lng { get; set; }
        }

        public class OsrmResponse
        {
            public OsrmRoute[] Routes { get; set; }
        }

        public class OsrmRoute
        {
            public double Distance { get; set; }
            public object Geometry { get; set; }
        }

        // ---------------- نقاط نمونه ----------------
        private List<Point> GetPoints()
        {
            return new List<Point>
            {
                new Point{ Id=1, Lat=25.291758665247, Lng=60.644739160469 },
                new Point{ Id=2, Lat=29.471318870585, Lng=60.833265659744 },
                new Point{ Id=3, Lat=26.653929924558, Lng=61.691340853247 },
                new Point{ Id=4, Lat=30.807520001728, Lng=61.37862000121 },
                new Point{ Id=5, Lat=28.916195924233, Lng=61.320723151251 },
                new Point{ Id=6, Lat=26.577004747149, Lng=59.635116070763 },
                new Point{ Id=7, Lat=26.200244034918, Lng=60.629287118288 },
                new Point{ Id=8, Lat=27.447302619348, Lng=59.579367260692 },
                new Point{ Id=9, Lat=28.261787634321, Lng=61.487407249933 },
                new Point{ Id=10, Lat=27.254310730197, Lng=62.027496371141 },
                new Point{ Id=11, Lat=25.743364943877, Lng=59.761214560754 },
                new Point{ Id=12, Lat=27.1261939999, Lng=61.688211000383 },
                new Point{ Id=13, Lat=27.23871184813, Lng=62.048361504739 },
                new Point{ Id=14, Lat=29.4914312262468, Lng=60.7806446403265 }
            };
        }
    }
}

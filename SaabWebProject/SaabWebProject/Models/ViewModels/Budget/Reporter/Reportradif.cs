using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.ViewModels.Budget.Reporter
{
    public class Reportradif
    {
        public List<long> less { get; set; }
        public List<long> creator { get; set; }
        public List<DateTime> time { get; set; }
        public long lesssum { get; set; }

        public Reportradif()
        {
            less = new List<long>();
            creator = new List<long>();
            lesssum = new long();
            time = new List<DateTime>();
        }
    }
    public class Reportradif2
    {
        public List<long> less { get; set; }
        public List<long> creator { get; set; }
        public List<DateTime> time { get; set; }
        public long lesssum { get; set; }
        public List<string> Title { get; set; }

        public Reportradif2()
        {
            less = new List<long>();
            creator = new List<long>();
            lesssum = new long();
            time = new List<DateTime>();
            Title = new List<string>();
        }
        public void AddData(long less, DateTime time, string Title)
        {
            this.less.Add(less);
            //this.creator.Add(creator);
            this.Title.Add(Title);

            this.time.Add(time);
        }


        public class Reportradif
        {
            
        }


    }
}
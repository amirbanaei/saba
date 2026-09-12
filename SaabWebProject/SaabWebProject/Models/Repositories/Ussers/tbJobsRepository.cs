using SaabWebProject.Models.DomainModels;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Repositories.Ussers
{
    public class tbJobsRepository
    {
        SaabEntities db = new SaabEntities();

        public bool Create(tbJobs job)
        {
            try
            {
                db.tbJobs.Add(job);
                return Convert.ToBoolean(db.SaveChanges());
            }
            catch
            {
                return false;
            }
        }
        public tbJobs Find(int id)
        {
            return db.tbJobs.Find(id);
        }
        public List<tbJobs> List()
        {
            return db.tbJobs.ToList();
        }
        public bool Update(tbJobs entity)
        {

            if (entity != null)
            {
                try
                {
                    db.Entry(entity).State = EntityState.Modified;
                    return Convert.ToBoolean(db.SaveChanges());
                }
                catch
                {

                    return false;
                }
            }
            else
            {
                return true;
            }

        }



        public bool ChangeActivation(tbJobs job)
        {
            var oldjob = db.tbJobs.Find(job.jb_ID);
            if (oldjob != null && oldjob.jb_IsActive != job.jb_IsActive)
            {
                try
                {
                    oldjob.jb_IsActive = job.jb_IsActive;
                    return Convert.ToBoolean(db.SaveChanges());
                }
                catch
                {
                    return false;
                }

            }
            else
            {
                return true;
            }

        }
        public int? IfExitsJob(string jobName)
        {
            if (jobName != null)
            {
                var jb = db.tbJobs.FirstOrDefault(p => p.jb_Name.Contains(jobName));
                if (jb != null)
                {
                    return jb.jb_ID;
                }
                else
                {
                    return null;
                }
                
            }
            else
            {
                return 0;
            }
        }
    }
}
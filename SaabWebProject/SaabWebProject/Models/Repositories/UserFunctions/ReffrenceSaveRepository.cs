using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using SaabWebProject.Models.Utilitis;
using System.Data.Entity.Infrastructure;
using System.Threading.Tasks;

namespace SaabWebProject.Models.Repositories.UserFunctions
{
    public class ReffrenceSaveRepository : IInterFace<tbReffrenceSave>
    {
        SaabEntities Context;
        public ReffrenceSaveRepository()
        {
            Context = new SaabEntities();
        }
        public ReffrenceSaveRepository(SaabEntities _context)
        {
            Context = _context;
        }
        public tbReffrenceSaveLevel GetUserWarning(int userId)
        {
            var startTime = DateTime.Now.GetShamsiDayOfMonth();


            var result = Context.tbReffrenceSaveLevelUser.Where(p => p.FK_UserID == userId).ToList();

            if (result != null)
            {
                tbReffrenceSaveLevel reff = new tbReffrenceSaveLevel();

                foreach (var item in result)
                {
                    var level = item.tbReffrenceSaveLevel;
                    var endTime = level.DeadLine + startTime;
                    if (startTime >= level.StartDaysInMonth && startTime <= endTime)
                    {
                        reff = level;
                    }
                }
                return reff;
            }
            else
            {
                return new tbReffrenceSaveLevel();
            }
        }
        public List<tbReffrenceSave> GetUserPeymans(int user_id = 0, bool issabt = false, int isFor = 0)
        {
            if (issabt == true)
            {
                var today = DateTime.Now.GetShamsiDayOfMonth();
                if (user_id == 0)
                {
                    if (isFor != 0)
                    {

                        var a = Context.tbReffrenceSaveLevel.Where(p => p.tbReffrenceSave.IsFor == isFor && p.Deleted != true && p.Deleted != true && today >= p.StartDaysInMonth && today <= (p.StartDaysInMonth + p.DeadLine)).Select(p => p.tbReffrenceSave).ToList();
                        return a.Distinct().ToList();
                    }
                    else
                    {
                        var a = Context.tbReffrenceSaveLevel.Where(p => today <= p.StartDaysInMonth && p.Deleted != true && today >= (p.StartDaysInMonth + p.DeadLine)).Select(p => p.tbReffrenceSave).ToList();
                        return a.Distinct().ToList();
                    }
                }
                else
                {
                    if (isFor != 0)
                    {
                        var a = Context.tbReffrenceSaveLevelUser.Where(p => p.tbReffrenceSaveLevel.tbReffrenceSave.IsFor == isFor && p.FK_UserID == user_id && today >= p.tbReffrenceSaveLevel.StartDaysInMonth && today <= (p.tbReffrenceSaveLevel.StartDaysInMonth + p.tbReffrenceSaveLevel.DeadLine)).Select(p => p.tbReffrenceSaveLevel.tbReffrenceSave).ToList().Distinct().ToList();
                        return a;
                    }
                    else
                    {
                        return Context.tbReffrenceSaveLevelUser.Where(p => p.FK_UserID == user_id && today >= p.tbReffrenceSaveLevel.StartDaysInMonth && today <= (p.tbReffrenceSaveLevel.StartDaysInMonth + p.tbReffrenceSaveLevel.DeadLine)).Select(p => p.tbReffrenceSaveLevel.tbReffrenceSave).ToList().Distinct().ToList();

                    }
                }
            }
            else
            {
                if (user_id == 0)
                {
                    return Context.tbReffrenceSave.ToList();
                }
                else
                {
                    return Context.tbReffrenceSaveLevelUser.Where(p => p.FK_UserID == user_id).Select(p => p.tbReffrenceSaveLevel.tbReffrenceSave).ToList().Distinct().ToList();
                }
            }

        }
        public IEnumerable<tbReffrenceSave> GetPeymansHasRSave()
        {
            return Context.tbReffrenceSave.ToList();

        }
        public IEnumerable<tbReffrenceSave> GetPymanIsforone()
        {
            var ext = Context.tbReffrenceSave.Where(p => p.IsFor == 1).ToList();
            return ext;

        }

        public List<tbReffrenceSaveLevel> GetPeymanBaste(int user_id = 0, int peyman_ID = 0)
        {

            var today = DateTime.Now.GetShamsiDayOfMonth();
            if (user_id == 0)
            {
                return Context.tbReffrenceSaveLevel.Where(p => p.tbReffrenceSave.FK_PeymanID == peyman_ID && p.Deleted != true).ToList();
            }
            else
            {
                return Context.tbReffrenceSaveLevelUser.Where(p => p.tbReffrenceSaveLevel.tbReffrenceSave.FK_PeymanID == peyman_ID && p.Deleted != true && p.FK_UserID == user_id && today >= p.tbReffrenceSaveLevel.StartDaysInMonth && today <= (p.tbReffrenceSaveLevel.StartDaysInMonth + p.tbReffrenceSaveLevel.DeadLine)).Select(p => p.tbReffrenceSaveLevel).ToList();
            }
        }
        public List<tbReffrenceSaveLevel> ReffrenceSaveGetAllBaste(int peyman_ID = 0, int Isfor = 0)
        {
            if (Isfor == 0)
            {
                if (peyman_ID != 0)
                {
                    return Context.tbReffrenceSaveLevel.Where(p => p.tbReffrenceSave.FK_PeymanID == peyman_ID && p.Deleted != true).ToList();
                }
                else
                {
                    return Context.tbReffrenceSaveLevelUser.Where(p => p.Deleted != true).Select(p => p.tbReffrenceSaveLevel).ToList();
                }
            }
            return null;

        }


        public List<tbReffrenceSaveLevel> ReffrenceSaveGetAllBasteFor(int peyman_ID = 0, int user = 0)

        {
          
                if (peyman_ID != 0)
                {
                return Context.tbReffrenceSaveLevel
                    .Where(p => p.tbReffrenceSave.FK_PeymanID == peyman_ID && p.Deleted != true && p.tbReffrenceSave.IsFor == 1 )
                    .ToList();
            }
                else
                {
                    return Context.tbReffrenceSaveLevelUser.Where(p => p.Deleted != true).Select(p => p.tbReffrenceSaveLevel).ToList();
                }
          

        }

        public List<tbReffrenceSaveLevel> ReffrenceSaveGetAllBasteForPy(int peyman_ID = 0, int user = 0)

        {
            var today = DateTime.Now.GetShamsiDayOfMonth();

            if (peyman_ID != 0)
            {
                return Context.tbReffrenceSaveLevel
                    .Where(p => p.tbReffrenceSave.FK_PeymanID == peyman_ID && p.Deleted != true && p.tbReffrenceSave.IsFor == 1&&p.tbReffrenceSaveLevelUser.Any(m=>m.FK_UserID==user)&& today >= p.StartDaysInMonth && today <= (p.StartDaysInMonth + p.DeadLine))
                    .ToList();
            }
            else
            {
                return Context.tbReffrenceSaveLevelUser.Where(p => p.Deleted != true).Select(p => p.tbReffrenceSaveLevel).ToList();
            }


        }



        public List<int> GetAllMoalefeInReffrenceSaveLevel(int RRL_ID)
        {
            List<int> mylist = new List<int>();
            var moalefes2 = Context.tbReffrenceSaveLevel.Find(RRL_ID) ;
            var moalefes = Context.tbReffrenceSaveLevelMoalefe.Where(p => p.FK_LevelID == moalefes2.ID && p.Deleted2 != true);
            foreach (var item in moalefes)
            {
                mylist.Add(item.FK_MoalefeID ?? 0);
            }
            return mylist;
        }
        public async Task<List<int>> GetAllMoalefeInReffrenceSaveLevel2(int RRL_ID)
        {
            return await Task.Run(() =>
            {
                List<int> mylist = new List<int>();
                var moalefes2 = Context.tbReffrenceSaveLevel.Find(RRL_ID);
                var moalefes = Context.tbReffrenceSaveLevelMoalefe.Where(p => p.FK_LevelID == moalefes2.ID && p.Deleted2 != true);

                foreach (var item in moalefes)
                {
                    mylist.Add(item.FK_MoalefeID ?? 0);
                }
                return mylist;
            });
        }

        public int? GetTypeOfBaste(int baste_id)
        {
            return Context.tbReffrenceSaveLevel.Where(p => p.ID == baste_id ).FirstOrDefault().tbReffrenceSave.TypeOfUpload;
        }

        public string Create(tbReffrenceSave obj)
        {
            try
            {
                foreach (var level in obj.tbReffrenceSaveLevel)
                {
                    level.Data_time_create = DateTime.Now;
                }

                // Your code to add the entity to the context and save changes
                Context.tbReffrenceSave.Add(obj);

                return SaveChanges().ToString();

                // Return a success message or value as needed
            }
            catch (DbUpdateException ex)
            {
                // Log the details of the exception for debugging
                Console.WriteLine("DbUpdateException: " + ex.Message);
                Console.WriteLine("Inner Exception: " + ex.InnerException?.Message);

                // Return an error message or handle the exception appropriately
                return "false";
            }
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }
        public bool Delete(int ID)
        {
            var a = Context.tbReffrenceSave.Find(ID);
            Context.tbReffrenceSaveLevelUser.RemoveRange(a.tbReffrenceSaveLevelUser);
            Context.tbReffrenceSaveLevelMoalefe.RemoveRange(a.tbReffrenceSaveLevelMoalefe);
            Context.tbReffrenceSaveLevel.RemoveRange(a.tbReffrenceSaveLevel);
            Context.tbReffrenceSave.Remove(a);
            return SaveChanges();
        }

        public tbReffrenceSave Find(int ID)
        {
            return Context.tbReffrenceSave.Find(ID);
        }

        public List<tbReffrenceSave> Update()
        {
            return Context.tbReffrenceSave.ToList();
        }
        public tbReffrenceSave Find(int peymanId, int type, int duration)
        {
            var a = Context.tbReffrenceSave.Where(p => p.FK_PeymanID == peymanId && p.TypeOfUpload == type && p.Duration == duration).FirstOrDefault();
            if (a == null)
            {
                a = new tbReffrenceSave();
            }
            return a;
        }
        public bool SaveChanges()
        {
            using (var tran = Context.Database.BeginTransaction())
            {
                try
                {
                    Context.SaveChanges();
                    tran.Commit();
                    return true;
                }
                catch (Exception ee)
                {

                    tran.Rollback();
                    return false;
                }
            }
        }






        public string Update(tbReffrenceSave obj)
        {
            var result = Context.tbReffrenceSave.Find(obj.ID);
            result.CountLevels = obj.CountLevels;
            result.Duration = obj.Duration;
            result.FK_PeymanID = obj.FK_PeymanID;
            result.MaxDaysForSave = obj.MaxDaysForSave;
            result.StartMonth = obj.StartMonth;
            result.TypeOfUpload = obj.TypeOfUpload;
            foreach (var item in result.tbReffrenceSaveLevel.Where(p => p.Deleted != true).ToList())
            {
                item.Deleted = true;
                foreach (var item2 in item.tbReffrenceSaveLevelUser.ToList())
                {
                    item2.Deleted = true;
                }
                foreach (var item2 in item.tbReffrenceSaveLevelMoalefe.Where(p=>p.Deleted2!=true).ToList())
                {
                    item2.Deleted = true;
                }
            }
            foreach (var item in obj.tbReffrenceSaveLevel.Where(p=>p.Deleted!=true).ToList())
            {
                item.Data_time_create = DateTime.Now;

                result.tbReffrenceSaveLevel.Add(item);
            }
            return SaveChanges().ToString();
        }
    }
}
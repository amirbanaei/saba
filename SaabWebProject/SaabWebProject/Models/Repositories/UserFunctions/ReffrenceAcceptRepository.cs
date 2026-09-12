using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using SaabWebProject.Models.UnitOfWork.Interfaces;
using SaabWebProject.Models.Utilitis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Web;

namespace SaabWebProject.Models.Repositories.UserFunctions
{
    public class ReffrenceAcceptRepository : IInterFace<tbReffrenceAccept>
    {
        SaabEntities Context;
        public ReffrenceAcceptRepository()
        {
            Context = new SaabEntities();
        }
        public ReffrenceAcceptRepository(SaabEntities _context)
        {
            Context = _context;
        }

        public string Create(tbReffrenceAccept obj)
        {
            Context.tbReffrenceAccept.Add(obj);
            return Convert.ToBoolean(Context.SaveChanges()).ToString();
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        public tbReffrenceAccept Find(int ID)
        {
            return Context.tbReffrenceAccept.Find(ID);
        }


        public List<tbReffrenceAccept> Update()
        { 
            return Context.tbReffrenceAccept.ToList();
        }
        public List<tbReffrenceAccept> Listt(int user_id)
        { 
            List<tbReffrenceAccept> returnList = new List<tbReffrenceAccept>();
            List<int> monthLsit = new List<int>() { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 };
            var today = DateTime.Now.GetShamsiDayOfMonth();
            var levels = Context.tbReffrenceAcceptLevel.Where(p => p.StartDayInMonth <= today && today <= (p.StartDayInMonth + p.Duration)).ToList();
            foreach(var lvl in levels)
            {
                var baste_duration = lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.Duration;
                if (baste_duration == 1)
                {
                    var lew = lvl.ID;
                    var l=Context.tbReffrenceAcceptLevelUsers.Where(p=>p.FK_User==user_id&&p.FK_ReffrenceAcceptLevel==lew).FirstOrDefault();
                    if (l != null)
                    {
                        if (lvl.FK_ReffrenceAccept == 1076)
                        {

                        }
                        if (lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.IsFor == 1&& lvl.tbReffrenceAccept.tbReffrenceSaveLevel.Deleted!=true)
                        {
                            returnList.Add(lvl.tbReffrenceAccept);
                        }
                       
                    }
                } else if(baste_duration == 2)
                {
                    //var counter = (lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.StartMonth + 1) ;
                    //while(counter != lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.StartMonth)
                    //{

                    //}
                }else if(baste_duration == 3)
                {

                }
            }

            //var result = Context.tbReffrenceAcceptLevelUsers.Where(p=>p.FK_User == user_id && p.)
            //return Context.tbReffrenceAccept.ToList();
            return returnList;
        }

        public List<tbReffrenceAccept> ListtForproje(int user_id)
        {


            List<tbReffrenceAccept> returnList = new List<tbReffrenceAccept>();
            List<int> monthLsit = new List<int>() { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 };
            var today = DateTime.Now.GetShamsiDayOfMonth();
            var levels = Context.tbReffrenceAcceptLevel.Where(p => p.StartDayInMonth <= today && today <= (p.StartDayInMonth + p.Duration)).ToList();
            foreach (var lvl in levels)
            {
                var baste_duration = lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.Duration;
                if (baste_duration == 1)
                {
                    var lew = lvl.ID;
                    var l = Context.tbReffrenceAcceptLevelUsers.Where(p => p.FK_User == user_id && p.FK_ReffrenceAcceptLevel == lew).FirstOrDefault();
                    if (l != null)
                    {
                        if (lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.IsFor==6 && lvl.tbReffrenceAccept.tbReffrenceSaveLevel.Deleted != true)
                        {
                            returnList.Add(lvl.tbReffrenceAccept);
                        }
                    }
                }
                else if (baste_duration == 2)
                {
                    //var counter = (lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.StartMonth + 1) ;
                    //while(counter != lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.StartMonth)
                    //{

                    //}
                }
                else if (baste_duration == 3)
                {

                }
            }

            //var result = Context.tbReffrenceAcceptLevelUsers.Where(p=>p.FK_User == user_id && p.)
            //return Context.tbReffrenceAccept.ToList();
            return returnList;
        }

        public List<tbReffrenceAccept> ListtForSorattahed(int user_id)
        {


            List<tbReffrenceAccept> returnList = new List<tbReffrenceAccept>();
            List<int> monthLsit = new List<int>() { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 };
            var today = DateTime.Now.GetShamsiDayOfMonth();
            var levels = Context.tbReffrenceAcceptLevel.Where(p => p.StartDayInMonth <= today && today <= (p.StartDayInMonth + p.Duration)).ToList();
            foreach (var lvl in levels)
            {
                var baste_duration = lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.Duration;
                if (baste_duration == 1)
                {
                    var lew = lvl.ID;
                    var l = Context.tbReffrenceAcceptLevelUsers.Where(p => p.FK_User == user_id && p.FK_ReffrenceAcceptLevel == lew).FirstOrDefault();
                    if (l != null)
                    {
                        if (lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.IsFor == 7 && lvl.tbReffrenceAccept.tbReffrenceSaveLevel.Deleted != true)
                        {
                            returnList.Add(lvl.tbReffrenceAccept);
                        }
                    }
                }
                else if (baste_duration == 2)
                {
                    //var counter = (lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.StartMonth + 1) ;
                    //while(counter != lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.StartMonth)
                    //{

                    //}
                }
                else if (baste_duration == 3)
                {

                }
            }

            //var result = Context.tbReffrenceAcceptLevelUsers.Where(p=>p.FK_User == user_id && p.)
            //return Context.tbReffrenceAccept.ToList();
            return returnList;
        }



        public List<tbReffrenceAccept> ListtProjSorat(int user_id)
        {


            List<tbReffrenceAccept> returnList = new List<tbReffrenceAccept>();
            List<int> monthLsit = new List<int>() { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 };
            var today = DateTime.Now.GetShamsiDayOfMonth();
            var levels = Context.tbReffrenceAcceptLevel.Where(p => p.StartDayInMonth <= today && today <= (p.StartDayInMonth + p.Duration)).ToList();
            foreach (var lvl in levels)
            {
                var baste_duration = lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.Duration;
                if (baste_duration == 1)
                {
                    var lew = lvl.ID;
                    var l = Context.tbReffrenceAcceptLevelUsers.Where(p => p.FK_User == user_id && p.FK_ReffrenceAcceptLevel == lew).FirstOrDefault();
                    if (l != null)
                    {
                        if (lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.IsFor == 5 && lvl.tbReffrenceAccept.tbReffrenceSaveLevel.Deleted != true)
                        {
                            returnList.Add(lvl.tbReffrenceAccept);
                        }
                    }
                }
                else if (baste_duration == 2)
                {
                    //var counter = (lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.StartMonth + 1) ;
                    //while(counter != lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.StartMonth)
                    //{

                    //}
                }
                else if (baste_duration == 3)
                {

                }
            }

            //var result = Context.tbReffrenceAcceptLevelUsers.Where(p=>p.FK_User == user_id && p.)
            //return Context.tbReffrenceAccept.ToList();
            return returnList;
        }
        public List<tbReffrenceAccept> ListtProjFish(int user_id)
        {


            List<tbReffrenceAccept> returnList = new List<tbReffrenceAccept>();
            List<int> monthLsit = new List<int>() { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 };
            var today = DateTime.Now.GetShamsiDayOfMonth();
            var levels = Context.tbReffrenceAcceptLevel.Where(p => p.StartDayInMonth <= today && today <= (p.StartDayInMonth + p.Duration)).ToList();
            foreach (var lvl in levels)
            {
                var baste_duration = lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.Duration;
                if (baste_duration == 1)
                {
                    var lew = lvl.ID;
                    var l = Context.tbReffrenceAcceptLevelUsers.Where(p => p.FK_User == user_id && p.FK_ReffrenceAcceptLevel == lew).FirstOrDefault();
                    if (l != null)
                    {
                        if (lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.IsFor == 7 && lvl.tbReffrenceAccept.tbReffrenceSaveLevel.Deleted != true)
                        {
                            returnList.Add(lvl.tbReffrenceAccept);
                        }
                    }
                }
                else if (baste_duration == 2)
                {
                    //var counter = (lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.StartMonth + 1) ;
                    //while(counter != lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.StartMonth)
                    //{

                    //}
                }
                else if (baste_duration == 3)
                {

                }
            }

            //var result = Context.tbReffrenceAcceptLevelUsers.Where(p=>p.FK_User == user_id && p.)
            //return Context.tbReffrenceAccept.ToList();
            return returnList;
        }



        public List<tbReffrenceAccept> Listt2(int user_id)
        {
            List<tbReffrenceAccept> returnList = new List<tbReffrenceAccept>();
            List<int> monthLsit = new List<int>() { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 };
            var today = DateTime.Now.GetShamsiDayOfMonth();
            var levels = Context.tbReffrenceAcceptLevel.Where(p => p.StartDayInMonth <= today && today <= (p.StartDayInMonth + p.Duration)).ToList();
            foreach (var lvl in levels)
            {
                var baste_duration = lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.Duration;
                if (baste_duration == 1)
                {
                    var lew = lvl.ID;
                    var l = Context.tbReffrenceAcceptLevelUsers.Where(p => p.FK_User == user_id && p.FK_ReffrenceAcceptLevel == lew).FirstOrDefault();
                    if (l != null)
                    {
                        if (lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.IsFor !=1 && lvl.tbReffrenceAccept.tbReffrenceSaveLevel.Deleted != true)
                        {
                            returnList.Add(lvl.tbReffrenceAccept);
                                              }
                    }
                }
                else if (baste_duration == 2)
                {
                    //var counter = (lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.StartMonth + 1) ;
                    //while(counter != lvl.tbReffrenceAccept.tbReffrenceSaveLevel.tbReffrenceSave.StartMonth)
                    //{

                    //}
                }
                else if (baste_duration == 3)
                {

                }
            }

            //var result = Context.tbReffrenceAcceptLevelUsers.Where(p=>p.FK_User == user_id && p.)
            //return Context.tbReffrenceAccept.ToList();
            return returnList;
        }







        public bool SaveChanges()
        {
            throw new NotImplementedException();
        }

        public string Update(tbReffrenceAccept obj)
        {
            throw new NotImplementedException();
        }
    }
}
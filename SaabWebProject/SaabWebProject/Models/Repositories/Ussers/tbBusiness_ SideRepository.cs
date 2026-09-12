using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.SqlServer.Server;
using SaabWebProject.Models.Interfaces;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.ViewModels.Users;

namespace SaabWebProject.Models.Repositories.Ussers
{
    public class BusinessSideRepository : IInterFace<tbBusinessSide>
    {
        private SaabEntities db;

        public BusinessSideRepository(SaabEntities _db)
        {
            db = _db;
        }

        public string Create(tbBusinessSide obj)
        {
            try
            {
                var res = db.tbBusinessSide.Where(p => p.up_name.Trim() == obj.up_name.Trim()).FirstOrDefault();
                if (res != null)
                {
                    return "Exists_Position";
                }

                obj.up_Status = true;
                db.tbBusinessSide.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {
                return false.ToString();
            }
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        public tbBusinessSide Find(int ID)
        {
            try
            {
                if (ID != 0)
                {
                    return db.tbBusinessSide.Find(ID);
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                return null;
            }
        }

        public List<tbBusinessSide> Update()
        {
            return db.tbBusinessSide.ToList();
        }

        public bool SaveChanges()
        {
            return Convert.ToBoolean(db.SaveChanges());
        }

        public string Update(tbBusinessSide obj)
        {
            try
            {
                db.Entry(obj).State = System.Data.Entity.EntityState.Modified;
                return SaveChanges().ToString();
            }
            catch
            {
                return false.ToString();
            }
        }

        public List<ViewModels.BusinessSide.BusinessSideViewModel> GetAllBusinessSidesForSelect()
        {
            return db.tbBusinessSide.Select(p => new ViewModels.BusinessSide.BusinessSideViewModel { up_id = p.up_Id, up_name = p.up_name }).ToList();
        }

        public bool ActiveOrDeActive(int id)
        {
            try
            {
                var obj = db.tbBusinessSide.Find(id);
                if ((bool)obj.up_Status)
                {
                    obj.up_Status = false;
                }
                else
                {
                    obj.up_Status = true;
                }

                return SaveChanges();
            }
            catch
            {
                return false;
            }
        }

        public int? IfExitsPosition(string PositionName)
        {
            try
            {
                if (PositionName != null)
                {
                    var pos = db.tbBusinessSide.FirstOrDefault(p => p.up_name == PositionName.Trim());
                    if (pos != null)
                    {
                        return pos.up_Id;
                    }
                    else
                    {
                        return null;
                    }
                }

                return 0;
            }
            catch
            {
                return 0;
            }
        }

        public tbBusinessSide GetBusinessSide(string PositionName)
        {
            return db.tbBusinessSide.Where(p => p.up_name.Contains(PositionName)).FirstOrDefault();
        }
        
        public bool SetAccessLevel(List<string> Actions, int PosId)
        {
            var pos = db.tbBusinessSide.Find(PosId);
            foreach (var item in pos.tbLink_UserAndAction.Where(p => p.FK_User_ID == null).ToList())
            {
                db.tbLink_UserAndAction.Remove(item);
            }

            if (Actions != null)
            {
                foreach (var item in Actions)
                {
                    var it = item.Split('_');
                    if (it[0] == "child")
                    {
                        if (it[0] == "0")
                        {
                            foreach (var item1 in db.tbActions.ToList())
                            {
                                var link = new tbLink_UserAndAction();
                                link.FK_Action_ID = item1.Action_ID;
                                link.FK_Position_ID = pos.up_Id;
                                link.HasAccess = true;
                                pos.tbLink_UserAndAction.Add(link);
                            }

                            break;
                        }
                        else
                        {
                            var link = new tbLink_UserAndAction();
                            link.FK_Action_ID = Convert.ToInt32(it[1]);
                            link.FK_Position_ID = pos.up_Id;
                            link.HasAccess = true;
                            pos.tbLink_UserAndAction.Add(link);
                        }
                    }
                }
            }

            return Convert.ToBoolean(db.SaveChanges());
        }

        public ActionViewModel GetAllActionsForUsers(int UserId)
        {
            ActionViewModel wm = new ActionViewModel();
            wm.ActiveActions = new List<string>();
            wm.AllActions = db.tbActions.ToList();
            var UserActions = db.tbLink_UserAndAction.Where(p => p.FK_User_ID == UserId).ToList();
            foreach (var item in UserActions)
            {
                wm.ActiveActions.Add("child_" + item.tbActions.Action_ID);
            }

            return wm;
        }
        public ActionViewModel GetAllActionsForPositions(int positionId)
        {
            ActionViewModel wm = new ActionViewModel();
            wm.ActiveActions = new List<string>();
            wm.AllActions = db.tbActions.ToList();
            var UserActions = db.tbLink_UserAndAction.Where(p => p.FK_Position_ID == positionId).ToList();
            foreach (var item in UserActions)
            {
                wm.ActiveActions.Add("child_" + item.tbActions.Action_ID);
            }

            return wm;
        }
    }
}
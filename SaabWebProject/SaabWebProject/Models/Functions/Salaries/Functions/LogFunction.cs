using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Repositories.Salaries;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SaabWebProject.Models.Functions.Salaries.Functions
{
    public class LogFunction
    {
        SaabEntities db;
        tbLogFunctionsRepository logfuncRepo;
        public LogFunction(SaabEntities context)
        {
            db = context;
            logfuncRepo = new tbLogFunctionsRepository(db);
        }
        public bool InsertToLogFunctionTable(tbLogFunctions entity)
        {
            try
            {
                if(logfuncRepo.Create(entity) =="True")
                {
                    return true;
                }
                else
                {
                    return false;
                }
                
            }
            catch (Exception)
            {
                return false;
                throw;
            }
        }
    }
}
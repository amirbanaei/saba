using EntityFramework.BulkInsert.Extensions;
using SaabWebProject.Models.DomainModels;
using SaabWebProject.Models.Interfaces;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using Telerik.Windows.Documents.Spreadsheet.Expressions.Functions;
//using Z.EntityFramework.Extensions;
using EntityFramework.BulkExtensions.Extensions;
using System.Data;
using System.Data.SqlClient;
using System.Data.Entity.Core.EntityClient;
using System.Configuration;

namespace SaabWebProject.Models.Repositories.Salaries
{
    public class tbMoalefeDastmozdiValueFromExcelRepository : IInterFace<tbMoalefeDastmozdiValueFromExcel>
    {
        SaabEntities db;
        public tbMoalefeDastmozdiValueFromExcelRepository(SaabEntities context)
        {
            db = context;
        }
        public tbMoalefeDastmozdiValueFromExcelRepository()
        {
            db = new SaabEntities();
        }
        public string Create(tbMoalefeDastmozdiValueFromExcel obj)
        {
            try
            {
                db.tbMoalefeDastmozdiValueFromExcel.Add(obj);
                return SaveChanges().ToString();
            }
            catch (Exception)
            {

                return "false";
            }
        }
        //public async Task<int> Create2Async(List<tbMoalefeDastmozdiValueFromExcel> list_obj)
        //{
        //    try
        //    {
        //        // استفاده از BulkInsert
        //        await db.BulkInsert(list_obj, options => {
        //            options.BatchSize = 5000; // تنظیم سایز دسته‌ها
        //            options.InsertKeepIdentity = false; // آیا می‌خواهید IDها حفظ شوند؟
        //        });

        //        return list_obj.Count;
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"Error saving data: {ex.Message}");
        //        return 0;
        //    }
        //}
        //public int BulkInsertData(List<tbMoalefeDastmozdiValueFromExcel> dataList)
        //{
        //    try
        //    {
        //        // استفاده از BulkInsert برای EF6
        //        db.BulkInsert(dataList);
        //        return dataList.Count;
        //    }
        //    catch (Exception ex)
        //    {
        //        System.Diagnostics.Debug.WriteLine($"خطا: {ex.Message}");
        //        return 0;
        //    }
        //}
        public int BulkInsertData2(List<tbMoalefeDastmozdiValueFromExcel> dataList)
        {
            try
            {
                // پاک‌سازی navigation property ها
                dataList.ForEach(item =>
                {
                    item.tbContractMoalefeDastmozdi = null;
                    item.tbSavedFunctions = null;
                    item.tbUsers = null;
                });

                // غیرفعال کردن AutoDetectChanges برای افزایش کارایی
                db.Configuration.AutoDetectChangesEnabled = false;

                // اجرای BulkInsert با EntityFramework.BulkExtensions
                db.BulkInsert(dataList);

                return dataList.Count;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"خطا: {ex.Message}");
                return 0;
            }
            finally
            {
                // بازگرداندن تنظیمات EF
                db.Configuration.AutoDetectChangesEnabled = true;
            }
        }
        //public DataService(YourDbContext context)
        //{
        //    db = context;
        //}

        public int BulkInsertData(List<tbMoalefeDastmozdiValueFromExcel> dataList)
        {
            if (dataList == null || dataList.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("لیست داده برای درج خالی است.");
                return 0;
            }

            // --- گام 1: پاک‌سازی navigation property ها (همانند کد شما) ---
            // این کار برای جلوگیری از تلاش EF برای درج یا به‌روزرسانی موجودیت‌های مرتبط هنگام درج اصلی ضروری است.
            dataList.ForEach(item =>
            {
                item.tbContractMoalefeDastmozdi = null;
                item.tbSavedFunctions = null;
                item.tbUsers = null;

                // --- مهم: اگر MoalfeVal_ID ستون Identity (خودکار) است، مطمئن شوید 0 باشد یا مقداردهی نشده باشد ---
                // اگر از قبل مقداری برای آن تنظیم شده باشد و Identity باشد، می‌تواند خطا بدهد.
                // اگر MoalfeVal_ID شما Identity است و برای رکوردهای جدید مقداردهی شده، خط زیر را غیرفعال کنید یا 0 کنید:
                // item.MoalfeVal_ID = 0; // فقط در صورتی که ID را از جایی می‌گیرید و می‌خواهید پایگاه داده آن را ایجاد کند
            });

            // --- گام 2: غیرفعال کردن AutoDetectChanges برای افزایش کارایی ---
            // این یک گام استاندارد برای Bulk Insert است.
            db.Configuration.AutoDetectChangesEnabled = false;
            var transaction = db.Database.CurrentTransaction; // بررسی تراکنش جاری
            if (transaction == null)
            {
                transaction = db.Database.BeginTransaction();
            }
            try
            {
                // --- گام 3: اجرای BulkInsert با EntityFramework.BulkExtensions ---
                db.BulkInsert(dataList);
                transaction.Commit(); // کامیت کردن تراکنش در صورت موفقیت

                System.Diagnostics.Debug.WriteLine($"تعداد {dataList.Count} رکورد با موفقیت درج شد.");
                return dataList.Count;
            }
            catch (Exception ex)
            {
                transaction.Rollback(); // بازگرداندن تراکنش در صورت خطا
                System.Diagnostics.Debug.WriteLine($"خطا در BulkInsertData: {ex.Message}");

                // --- مهم: بررسی Inner Exception برای جزئیات بیشتر ---
                // این قسمت کلیدی برای عیب‌یابی خطای "KeyNotFoundException" است.
                // Inner Exception دقیقاً به شما می‌گوید که کدام کلید پیدا نشده است.
                if (ex.InnerException != null)
                {
                    System.Diagnostics.Debug.WriteLine($"Inner Exception: {ex.InnerException.Message}");
                    // اگر باز هم Inner Exception داشت، آن را هم نمایش دهید
                    if (ex.InnerException.InnerException != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"Deep Inner Exception: {ex.InnerException.InnerException.Message}");
                    }
                }
                return 0;
            }
            finally
            {
                // --- گام 4: بازگرداندن تنظیمات EF ---
                db.Configuration.AutoDetectChangesEnabled = true;
                // تراکنش در اینجا مدیریت شده، نیازی به بستن صریح ندارد (Rollback/Commit آن را می‌بندد).
            }
        }
        public string Create(List<tbMoalefeDastmozdiValueFromExcel> list_obj)
        {
            try
            {
                if (list_obj == null || !list_obj.Any())
                {
                    return "False - List is null or empty.";
                }

                // Assuming 'db' is a valid and initialized database context
                // Check if 'db' is not null and BulkInsert method is accessible
                if (db != null)
                {
                    db.BulkInsert(list_obj);
                    return "True";
                }
                else
                {
                    return "False - Database context is not initialized.";
                }
            }
            catch (Exception ex)
            {
                // Log or handle the exception appropriately
                Console.WriteLine($"Error in Create method: {ex.Message}");
                return "False - An error occurred.";
            }
        }
        public async Task<int> BatchInsertMoalefeAsync(List<tbMoalefeDastmozdiValueFromExcel> dataList, int batchSize = 50000)
        {
            try
            {
                if (dataList.Count <500)
                {
                    batchSize = 1;
                }

                int totalInserted = 0;

                for (int i = 0; i < dataList.Count; i += batchSize)
                {
                    var batch = dataList.Skip(i).Take(batchSize).ToList();

                    db.tbMoalefeDastmozdiValueFromExcel.AddRange(batch);
                    await db.SaveChangesAsync();

                    // Detach entities to prevent memory leak
                    foreach (var item in batch)
                    {
                        db.Entry(item).State = EntityState.Detached;
                    }

                    totalInserted += batch.Count;
                }

                return totalInserted;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Batch Insert Error: {ex.Message}");
                return 0;
            }
        }
        //public async Task<int> Create2Async2(List<tbMoalefeDastmozdiValueFromExcel> list_obj)
        //{
        //    try
        //    {
        //        // روش صحیح استفاده از BulkInsertAsync
        //        await using var transaction = await db.Database.BeginTransactionAsync();
        //        await db.BulkInsertAsync(list_obj);
        //        await transaction.CommitAsync();

        //        return list_obj.Count;
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"خطا در ذخیره داده‌ها: {ex.Message}");
        //        return 0;
        //    }
        //}

        //public async Task<int> Create2Async22(List<tbMoalefeDastmozdiValueFromExcel> list_obj)
        //    {
        //        try
        //        {
        //            // استفاده از BulkInsert
        //            await db.BulkInsertAsync(list_obj, config => {
        //                config.BatchSize = 5000;
        //                // سایر تنظیمات اختیاری:
        //                // config.InsertIfNotExists = true;
        //                // config.PropertiesToInclude = new List<string> { "Field1", "Field2" };
        //            });

        //            return list_obj.Count;
        //        }
        //        catch (Exception ex)
        //        {
        //            Console.WriteLine($"Error saving data: {ex.Message}");
        //            return 0;
        //        }
        //    }
        //public async Task<int> BulkInsertDataAsync(List<tbMoalefeDastmozdiValueFromExcel> dataToInsert)
        //{
        //    if (dataToInsert == null || dataToInsert.Count == 0)
        //    {
        //        return 0;
        //    }

        //    // Convert the list of objects to a DataTable
        //    DataTable dataTable = new DataTable();
        //    dataTable.Columns.Add("MoalfeVal_FKMoalafeDastmozdi", typeof(int));
        //    dataTable.Columns.Add("MoalfeVal_FKUser", typeof(int));
        //    dataTable.Columns.Add("MoalfeVal_Month", typeof(int));
        //    dataTable.Columns.Add("MoalfeVal_Year", typeof(int));
        //    dataTable.Columns.Add("MoalfeVal_Value", typeof(double));
        //    dataTable.Columns.Add("FK_SavedFunctionsID", typeof(int));
        //    dataTable.Columns.Add("Final_accept", typeof(bool));
        //    // Add other columns as needed, matching the database table schema.
        //    // Make sure to match column names exactly as they are in the database.

        //    foreach (var item in dataToInsert)
        //    {
        //        dataTable.Rows.Add(
        //            item.MoalfeVal_FKMoalafeDastmozdi,
        //            item.MoalfeVal_FKUser,
        //            item.MoalfeVal_Month,
        //            item.MoalfeVal_Year,
        //            item.MoalfeVal_Value,
        //            item.FK_SavedFunctionsID,
        //            item.Final_accept
        //        );
        //    }

        //    using (var sqlConnection = new SqlConnection(db.Database.Connection.ConnectionString))
        //    {
        //        await sqlConnection.OpenAsync();
        //        using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlConnection))
        //        {
        //            bulkCopy.DestinationTableName = "[Salary].[tbMoalefeDastmozdiValueFromExcel]"; // Replace with your actual schema and table name

        //            // Map your DataTable columns to the database table columns
        //            bulkCopy.ColumnMappings.Add("MoalfeVal_FKMoalafeDastmozdi", "MoalfeVal_FKMoalafeDastmozdi");
        //            bulkCopy.ColumnMappings.Add("MoalfeVal_FKUser", "MoalfeVal_FKUser");
        //            bulkCopy.ColumnMappings.Add("MoalfeVal_Month", "MoalfeVal_Month");
        //            bulkCopy.ColumnMappings.Add("MoalfeVal_Year", "MoalfeVal_Year");
        //            bulkCopy.ColumnMappings.Add("MoalfeVal_Value", "MoalfeVal_Value");
        //            bulkCopy.ColumnMappings.Add("FK_SavedFunctionsID", "FK_SavedFunctionsID");
        //            bulkCopy.ColumnMappings.Add("Final_accept", "Final_accept");
        //            // Add other column mappings

        //            try
        //            {
        //                await bulkCopy.WriteToServerAsync(dataTable);
        //                return dataToInsert.Count; // Assuming all rows are inserted successfully
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine($"Error during bulk insert: {ex.Message}");
        //                // Log the exception details
        //                return 0;
        //            }
        //        }
        //    }
        //     }





        public async Task<int> BulkInsertDataAsyncfish(List<tbMoalefeValueFish> dataToInsert)
        {
            if (dataToInsert == null || dataToInsert.Count == 0)
            {
                return 0;
            }

            try
            {
                DataTable dataTable = CreateDataTablefish(dataToInsert);

                string sqlConnectionString = GetSqlConnectionString();

                return await ExecuteBulkInsertfish(dataTable, sqlConnectionString);
            }
            catch (SqlException sqlEx)
            {
                LogError($"SQL Error: {sqlEx.Number} - {sqlEx.Message}", sqlEx);
                return 0;
            }
            catch (Exception ex)
            {
                LogError($"General Error: {ex.Message}", ex);
                return 0;
            }
        }


        public async Task<int> BulkInsertDataAsync(List<tbMoalefeDastmozdiValueFromExcel> dataToInsert)
        {
            if (dataToInsert == null || dataToInsert.Count == 0)
            {
                return 0;
            }

            try
            {
                DataTable dataTable = CreateDataTable(dataToInsert);

                string sqlConnectionString = GetSqlConnectionString();

                return await ExecuteBulkInsert(dataTable, sqlConnectionString);
            }
            catch (SqlException sqlEx)
            {
                LogError($"SQL Error: {sqlEx.Number} - {sqlEx.Message}", sqlEx);
                return 0;
            }
            catch (Exception ex)
            {
                LogError($"General Error: {ex.Message}", ex);
                return 0;
            }
        }
        private DataTable CreateDataTablefish(List<tbMoalefeValueFish> data)
        {
            DataTable dataTable = new DataTable();

            dataTable.Columns.Add("FK_Moalefe", typeof(int));
            dataTable.Columns.Add("FK_User", typeof(int));
            dataTable.Columns.Add("mlfvlfsh_Month", typeof(int));
            dataTable.Columns.Add("mlfvlfsh_Year", typeof(int));
            dataTable.Columns.Add("mlfvlfsh_Value", typeof(double));
            dataTable.Columns.Add("FK_EXCel", typeof(int));
            dataTable.Columns.Add("Finalaccept", typeof(bool));
            dataTable.Columns.Add("Datatmie", typeof(DateTime));

            foreach (var item in data)
            {
                dataTable.Rows.Add(
                    item.FK_Moalefe,
                    item.FK_User,
                    item.mlfvlfsh_Month,
                    item.mlfvlfsh_Year,
                    item.mlfvlfsh_Value,
                    item.FK_EXCel,

                    item.Finalaccept,
                                        item.Datatmie


                );
            }

            return dataTable;
        }

        private DataTable CreateDataTable(List<tbMoalefeDastmozdiValueFromExcel> data)
        {
            DataTable dataTable = new DataTable();

            dataTable.Columns.Add("MoalfeVal_FKMoalafeDastmozdi", typeof(int));
            dataTable.Columns.Add("MoalfeVal_FKUser", typeof(int));
            dataTable.Columns.Add("MoalfeVal_Month", typeof(int));
            dataTable.Columns.Add("MoalfeVal_Year", typeof(int));
            dataTable.Columns.Add("MoalfeVal_Value", typeof(double));
            dataTable.Columns.Add("FK_SavedFunctionsID", typeof(int));
            dataTable.Columns.Add("Final_accept", typeof(bool));

            foreach (var item in data)
            {
                dataTable.Rows.Add(
                    item.MoalfeVal_FKMoalafeDastmozdi,
                    item.MoalfeVal_FKUser,
                    item.MoalfeVal_Month,
                    item.MoalfeVal_Year,
                    item.MoalfeVal_Value,
                    item.FK_SavedFunctionsID,
                    item.Final_accept
                );
            }

            return dataTable;
        }

        private string GetSqlConnectionString()
        {
            try
            {
                // روش 1: استخراج از Entity Connection موجود
                var entityConnection = new EntityConnection(ConfigurationManager.ConnectionStrings["SaabEntities"].ConnectionString);
                return entityConnection.StoreConnection.ConnectionString;
            }
            catch
            {
                // روش 2: استفاده از connection string مستقیم (اگر روش اول کار نکرد)
                return "REPLACE_WITH_CONNECTION_STRING";
            }
        }

        private async Task<int> ExecuteBulkInsert(DataTable dataTable, string connectionString)
        {
            using (var sqlConnection = new SqlConnection(connectionString))
            {
                await sqlConnection.OpenAsync();

                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlConnection))
                {
                    bulkCopy.DestinationTableName = "[Salary].[tbMoalefeDastmozdiValueFromExcel]";
                    bulkCopy.BatchSize = 1000;
                    bulkCopy.BulkCopyTimeout = 600;

                    // تنظیم مپینگ ستون‌ها
                    bulkCopy.ColumnMappings.Add("MoalfeVal_FKMoalafeDastmozdi", "MoalfeVal_FKMoalafeDastmozdi");
                    bulkCopy.ColumnMappings.Add("MoalfeVal_FKUser", "MoalfeVal_FKUser");
                    bulkCopy.ColumnMappings.Add("MoalfeVal_Month", "MoalfeVal_Month");
                    bulkCopy.ColumnMappings.Add("MoalfeVal_Year", "MoalfeVal_Year");
                    bulkCopy.ColumnMappings.Add("MoalfeVal_Value", "MoalfeVal_Value");
                    bulkCopy.ColumnMappings.Add("FK_SavedFunctionsID", "FK_SavedFunctionsID");
                    bulkCopy.ColumnMappings.Add("Final_accept", "Final_accept");

                    await bulkCopy.WriteToServerAsync(dataTable);
                    return dataTable.Rows.Count;
                }
            }
        }
        private async Task<int> ExecuteBulkInsertfish(DataTable dataTable, string connectionString)
        {
            using (var sqlConnection = new SqlConnection(connectionString))
            {
                await sqlConnection.OpenAsync();

                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlConnection))
                {
                    bulkCopy.DestinationTableName = "[Salary].[tbMoalefeValueFish]";
                    bulkCopy.BatchSize = 1000;
                    bulkCopy.BulkCopyTimeout = 600;

                    // تنظیم مپینگ ستون‌ها
                    bulkCopy.ColumnMappings.Add("FK_Moalefe", "FK_Moalefe");
                    bulkCopy.ColumnMappings.Add("FK_User", "FK_User");
                    bulkCopy.ColumnMappings.Add("mlfvlfsh_Month", "mlfvlfsh_Month");
                    bulkCopy.ColumnMappings.Add("mlfvlfsh_Year", "mlfvlfsh_Year");
                    bulkCopy.ColumnMappings.Add("mlfvlfsh_Value", "mlfvlfsh_Value");
                    bulkCopy.ColumnMappings.Add("FK_EXCel", "FK_EXCel");
                    bulkCopy.ColumnMappings.Add("Finalaccept", "Finalaccept");
                    bulkCopy.ColumnMappings.Add("Datatmie", "Datatmie");

                    await bulkCopy.WriteToServerAsync(dataTable);
                    return dataTable.Rows.Count;
                }
            }
        }

        private void LogError(string message, Exception ex)
        {
            Console.WriteLine($"{message}\nStackTrace: {ex.StackTrace}");

            if (ex.InnerException != null)
            {
                Console.WriteLine($"Inner Exception: {ex.InnerException.Message}");
            }

            // اینجا می‌توانید خطا را به سیستم لاگینگ خود نیز ارسال کنید
        }











        public async Task<int> Create2Async(List<tbMoalefeDastmozdiValueFromExcel> list_obj)
        {
            try
            {
                
                    db.tbMoalefeDastmozdiValueFromExcel.AddRange(list_obj);
                    await db.SaveChangesAsync(); // Use async/await for asynchronous saving
                     // Commit the transaction
                    return list_obj.Count; // Return the number of saved items
                
            }
            catch (Exception ex)
            {
                // Log the exception and handle appropriately (e.g., rollback transaction)
                Console.WriteLine($"Error saving data: {ex.Message}");
                return 0;
            }
        }
        //public async Task<int> Create2Async2(List<tbMoalefeDastmozdiValueFromExcel> list_obj)
        //{
        //    try
        //    {
        //        await db.BulkInsertAsync(list_obj);
        //        return list_obj.Count;
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"Error saving data: {ex.Message}");
        //        return 0;
        //    }
        //}

        public string Createrial(List<tbMoalefeValueFish> list_obj)
        {
            try
            {
                if (list_obj == null || !list_obj.Any())
                {
                    return "False - List is null or empty.";
                }

                // Assuming 'db' is a valid and initialized database context
                // Check if 'db' is not null and BulkInsert method is accessible
                if (db != null)
                {
                    db.BulkInsert(list_obj);
                    return "True";
                }
                else
                {
                    return "False - Database context is not initialized.";
                }
            }
            catch (Exception ex)
            {
                // Log or handle the exception appropriately
                Console.WriteLine($"Error in Create method: {ex.Message}");
                return "False - An error occurred.";
            }
        }
        public async Task<int> Create2Asyncial(List<tbMoalefeValueFish> list_obj)
        {
            try
            {

                db.tbMoalefeValueFish.AddRange(list_obj);
                await db.SaveChangesAsync(); // Use async/await for asynchronous saving
                                             // Commit the transaction
                return list_obj.Count; // Return the number of saved items

            }
            catch (Exception ex)
            {
                // Log the exception and handle appropriately (e.g., rollback transaction)
                Console.WriteLine($"Error saving data: {ex.Message}");
                return 0;
            }
        }

        public string Create2(List<tbMoalefeDastmozdiValueFromExcel> list_obj)
        {
            try
            {
                db.tbMoalefeDastmozdiValueFromExcel.AddRange(list_obj);

                return SaveChanges().ToString();
            }
            catch (Exception ex)
            {

                return "false";
            }
        }

        public bool Disable(int ID)
        {
            throw new NotImplementedException();
        }

        public tbMoalefeDastmozdiValueFromExcel Find(int ID)
        {
            throw new NotImplementedException();
        }

        public List<tbMoalefeDastmozdiValueFromExcel> Update()
        {
            return db.tbMoalefeDastmozdiValueFromExcel.ToList();
        }

        public bool SaveChanges()
        {
            try
            {
                return System.Convert.ToBoolean(db.SaveChanges());

            }
            catch (Exception ex){
                
                return false;
            }

        }

        public string Update(tbMoalefeDastmozdiValueFromExcel obj)
        {
            throw new NotImplementedException();
        }

  
    }
}
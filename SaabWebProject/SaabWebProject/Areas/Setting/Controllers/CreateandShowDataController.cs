using SaabWebProject.Utility;
using Stimulsoft.Data.Expressions.NCalc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using Stimulsoft.Data.Expressions.NCalc;
using System;
using System.Web.Mvc;
//using NCalc;
using System.Web.UI.WebControls;
namespace SaabWebProject.Areas.Setting.Controllers
{
    public class CreateandShowDataController : Controller
    {
        // GET: Setting/CreateandShowData

        [AuthorizeAAA]
        public ActionResult Index()
        {
            return View();
        }

        [AuthorizeAAA]
        public ActionResult _Definition_Formula()
        {
            return PartialView("~/Areas/Salaries/Views/Formula/Index.cshtml");
        }

        [AuthorizeAAA]
        public ActionResult ForumMaker()
        {
            return PartialView("~/Areas/Salaries/Views/ReportBuilder/Index.cshtml");

        }
        [AuthorizeAAA]
        public ActionResult SetReports()
        {
            return PartialView();
        }



//        // Custom SUMIF function implementation
//        public static object CustomSUMIF2(object[] args)
//        {
//            if (args.Length < 2 || args.Length > 3)
//                throw new ArgumentException("Invalid number of arguments. Expected: SUMIF(range, condition [, sum_range])");

//            // Extract arguments from the function call
//            IEnumerable<object> range = args[0] as IEnumerable<object>;
//            string condition = args[1].ToString();
//            IEnumerable<object> sumRange = args.Length == 3 ? args[2] as IEnumerable<object> : range;

//            // Evaluate the condition against each element in the range and sum the corresponding values
//            decimal sum = 0;
//            var enumerator = range.GetEnumerator();
//            var sumEnumerator = sumRange.GetEnumerator();

//            while (enumerator.MoveNext() && sumEnumerator.MoveNext())
//            {
//                // Create a new NCalc.Expression instance to evaluate the condition for the current element
//                var element = enumerator.Current;
//                var expr = new NCalc.Expression(condition);
//                expr.Parameters["item"] = element; // Assign the current element to a parameter 'item'

//                if (Convert.ToBoolean(expr.Evaluate()))
//                {
//                    sum += Convert.ToDecimal(sumEnumerator.Current);
//                }
//            }

//            return sum;
//        }
//        public static object CustomSUMIF(object[] args)
//        {
//            if (args.Length != 5) // برای ۵ آرگومان، باید دقیقاً ۵ آرگومان موجود باشد
//                throw new ArgumentException("Invalid number of arguments. Expected: SUMIF(range, condition, sum_range)");

//            // استخراج آرگومان‌ها از فراخوانی تابع
//            IEnumerable<object> range = args[0] as IEnumerable<object>;
//            string condition = args[1].ToString();
//            IEnumerable<object> sumRange = args[2] as IEnumerable<object>;

//            // ارزیابی شرایط و جمع مقادیر متناظر با شرایط
//            decimal sum = 0;
//            var enumerator = range.GetEnumerator();
//            var sumEnumerator = sumRange.GetEnumerator();

//            while (enumerator.MoveNext() && sumEnumerator.MoveNext())
//            {
//                // ایجاد نمونه‌ای جدید از NCalc.Expression برای ارزیابی شرط برای عنصر فعلی
//                var element = enumerator.Current;
//                var expr = new NCalc.Expression(condition);
//                expr.Parameters["item"] = element; // اختصاص عنصر فعلی به پارامتر 'item'

//                if (Convert.ToBoolean(expr.Evaluate()))
//                {
//                    sum += Convert.ToDecimal(sumEnumerator.Current);
//                }
//            }

//            return sum;
//        }


//public ActionResult SetReport444s()
//    {
//        var t = 1;
//        var data = new List<int> { 3, 5, 7, 9, 11 }; // Sample data for SUMIF

//        // Define your formula strings with the variable 't' replaced by its value
//        string formulaString = "(2*1+(7*8)+(t > 5 ? (7+8) : 0))";
//        string formulaString2 = "(2*1+(7*8)+(t > 0 ? (7+8) : 0))";

//        // Replace special characters with appropriate operators or functions
//        formulaString = formulaString.Replace("_", "-"); // Replace '_' with '-'
//        formulaString = formulaString.Replace("if", "if"); // Replace 'if' with 'if'
//        formulaString = formulaString.Replace("count", "count()"); // Replace 'count' with 'count()'

//        Console.WriteLine($"Modified Formula String: '{formulaString}'");

//        // Create an expression using NCalc.Expression
//        NCalc.Expression expression = new NCalc.Expression(formulaString);
//        NCalc.Expression expression2 = new NCalc.Expression(formulaString2);

//        // Define the parameter 't' with its value for the expressions
//        expression.Parameters["t"] = t;
//        expression2.Parameters["t"] = t;

//        try
//        {
//            // Evaluate the expressions
//            object result = expression.Evaluate();
//            object result2 = expression2.Evaluate();

//            // Calculate SUMIF
//            int sumifResult = data.Where(x => x > 5).Sum();

//            var ty = t + (int)result;
//            // Print the result to the console
//            string message = $"Result of '{formulaString}' = {result}, SUMIF Result = {sumifResult}";
//            Console.WriteLine(message);

//            // Optionally, you can store the result and pass it to the View
//            ViewBag.ExpressionResult = message; // Store the result in a ViewBag

//            // Return a View with the result
//            return View();
//        }
//        catch (NCalc.EvaluationException ex)
//        {
//            // Handle specific NCalc evaluation exceptions
//            Console.WriteLine($"Error evaluating expression: {ex.Message}");

//            // Optionally, you can pass an error message to the View
//            ViewBag.ErrorMessage = $"Error evaluating expression: {ex.Message}";

//            // Return a View indicating an error occurred
//            return View("ErrorView"); // Create an ErrorView to display error messages
//        }
//        catch (Exception ex)
//        {
//            // Handle other generic exceptions
//            Console.WriteLine($"Unexpected error: {ex.Message}");

//            // Optionally, you can pass an error message to the View
//            ViewBag.ErrorMessage = $"Unexpected error: {ex.Message}";

//            // Return a View indicating an unexpected error occurred
//            return View("ErrorView"); // Create an ErrorView to display error messages
//        }
//    }


//public ActionResult SetReport44s()
//        {
//            var t = 1;

//            try
//            {
//                // Define your formula strings
//                string formulaString = "(2*1+(7*8)+(t > 5 ? (7+8) : 0))";
//                string formulaString2 = "SUMIF({7, 8}, \"t > 0\", {7, 8})";

//                // Replace special characters with appropriate operators or functions
//                formulaString = formulaString.Replace("_", "-"); // Replace '_' with '-'
//                formulaString = formulaString.Replace("if", "if"); // Replace 'if' with 'if'
//                formulaString = formulaString.Replace("count", "count()"); // Replace 'count' with 'count()'

//                Console.WriteLine($"Modified Formula String: '{formulaString}'");

//                // Create an expression using NCalc.Expression
//                NCalc.Expression expression = new NCalc.Expression(formulaString);
//                NCalc.Expression expression2 = new NCalc.Expression(formulaString2);

//                // Define the parameter 't' with its value for the expressions
//                expression.Parameters["t"] = t;
//                expression2.Parameters["t"] = t;

//                // Evaluate the expressions
//                object result = expression.Evaluate();
//                object result2 = expression2.Evaluate();

//                // Optionally, you can store the result and pass it to the View
//                ViewBag.ExpressionResult = $"Result of '{formulaString}' = {result}, Result of '{formulaString2}' = {result2}";

//                // Return a View with the result
//                return View();
//            }
//            catch (NCalc.EvaluationException ex)
//            {
//                // Handle evaluation exceptions
//                Console.WriteLine($"Error evaluating expression: {ex.Message}");

//                // Optionally, you can pass an error message to the View
//                ViewBag.ErrorMessage = $"Error evaluating expression: {ex.Message}";

//                // Return a View indicating an error occurred
//                return View("ErrorView");
//            }
//            catch (Exception ex)
//            {
//                // Handle other generic exceptions
//                Console.WriteLine($"Unexpected error: {ex.Message}");

//                // Optionally, you can pass an error message to the View
//                ViewBag.ErrorMessage = $"Unexpected error: {ex.Message}";

//                // Return a View indicating an unexpected error occurred
//                return View("ErrorView");
//            }
//        }

//        public static int SumIf(IEnumerable<int> data, Func<int, bool> condition)
//        {
//            int sum = 0;
//            foreach (int item in data)
//            {
//                if (condition(item))
//                {
//                    sum += item;
//                }
//            }
//            return sum;
//        }
    


//    public ActionResult SetReport448s()
//        {
//            var t = 1;

//            // Define the formula strings
//            string formulaString = "(2*1+(7*8)+(t > 5 ? (7+8) : 0))";
//            string formulaString4 = "(2*1+(7*8)+SUMIF({7, 8}, \"t > 0\", {7, 8}))";
//            string formulaString3 = "(2*1+(7*8)+SUMIF(new [] {7, 8}, t > 5, new [] {7, 8}))";
//            string formulaString2 = "SUMIF({7, 8}, \"t > 0\", {7, 8})";


//            // Replace special characters with appropriate operators or functions
//            formulaString = formulaString.Replace("_", "-").Replace("if", "if").Replace("count", "count()");
//            formulaString2 = formulaString2.Replace("_", "-").Replace("if", "if").Replace("count", "count()");
//            formulaString3 = formulaString2.Replace("_", "-").Replace("if", "if").Replace("count", "count()");
//            formulaString4= formulaString4.Replace("_", "-").Replace("if", "if").Replace("count", "count()");

//            Console.WriteLine($"Modified Formula String 1: '{formulaString}'");
//            Console.WriteLine($"Modified Formula String 2: '{formulaString2}'");

//            // Create expressions using NCalc.Expression
//            NCalc.Expression expression = new NCalc.Expression(formulaString);
//            NCalc.Expression expression2 = new NCalc.Expression(formulaString2);
//            NCalc.Expression expression3 = new NCalc.Expression(formulaString3);
//            NCalc.Expression expression4 = new NCalc.Expression(formulaString4);

//            // Define the parameter 't' with its value for the expressions
//            expression.Parameters["t"] = t;
//            expression2.Parameters["t"] = t;
//            expression3.Parameters["t"] = t;
//            expression4.Parameters["t"] = t;

//            // Register the custom SUMIF function
//            expression.EvaluateFunction += (name, args) =>
//            {
//                if (name == "SUMIF")
//                {
//                    args.Result = CustomSUMIF(args.Parameters);
//                }
//            };
//            expression2.EvaluateFunction += (name, args) =>
//            {
//                if (name == "SUMIF")
//                {
//                    args.Result = CustomSUMIF(args.Parameters);
//                }
//            };
//            expression3.EvaluateFunction += (name, args) =>
//            {
//                if (name == "SUMIF")
//                {
//                    args.Result = CustomSUMIF(args.Parameters);
//                }
//            };
//            expression4.EvaluateFunction += (name, args) =>
//            {
//                if (name == "SUMIF")
//                {
//                    args.Result = CustomSUMIF(args.Parameters);
//                }
//            };

//            try
//            {
//                // Evaluate the expressions
//                object result = expression.Evaluate();
//                object result2 = expression2.Evaluate();
//                object result3 = expression3.Evaluate();
//                object result4 = expression4.Evaluate();

//                var ty = t + (int)result;
//                var ty2 = t + (decimal)result2;

//                // Print the results to the console
//                string message1 = $"Result of '{formulaString}' = {result}";
//                string message2 = $"Result of '{formulaString2}' = {result2}";
//                string message3 = $"Result of '{formulaString}' = {result3}";

//                string message4 = $"Result of '{formulaString}' = {result4}";

//                Console.WriteLine(message1);
//                Console.WriteLine(message2);

//                // Optionally, you can store the results and pass them to the View
//                ViewBag.ExpressionResult1 = message1; // Store the result of formulaString in a ViewBag
//                ViewBag.ExpressionResult2 = message2; // Store the result of formulaString2 in a ViewBag

//                // Return a View with the results
//                return View();
//            }
//            catch (NCalc.EvaluationException ex)
//            {
//                // Handle specific NCalc evaluation exceptions
//                Console.WriteLine($"Error evaluating expression: {ex.Message}");

//                // Optionally, you can pass an error message to the View
//                ViewBag.ErrorMessage = $"Error evaluating expression: {ex.Message}";

//                // Return a View indicating an error occurred
//                return View("ErrorView");
//            }
//            catch (Exception ex)
//            {
//                // Handle other generic exceptions
//                Console.WriteLine($"Unexpected error: {ex.Message}");

//                // Optionally, you can pass an error message to the View
//                ViewBag.ErrorMessage = $"Unexpected error: {ex.Message}";

//                // Return a View indicating an unexpected error occurred
//                return View("ErrorView");
//            }
//        }

//        // Assuming ASP.NET MVC

//        public ActionResult SetReport44s4()
//        {
//            var t = 1;
//            string formulaString = "(2*1+(7*8)+(t > 5 ? (7+8) : 0))";
//            string formulaString2 = "(2*1+(7*8)+(t > 0 ? (7+8) : 0))";


//            // Replace special characters with appropriate operators or functions
//            formulaString = formulaString.Replace("_", "-"); // Replace '_' with '-'
//            formulaString = formulaString.Replace("if", "if"); // Replace 'if' with 'if'
//            formulaString = formulaString.Replace("count", "count()"); // Replace 'count' with 'count()'

//            Console.WriteLine($"Modified Formula String: '{formulaString}'");

//            // Create an expression using NCalc.Expression
//            NCalc.Expression expression = new NCalc.Expression(formulaString);
//            NCalc.Expression expression2 = new NCalc.Expression(formulaString2);
//            expression.Parameters["t"] = t;
//            expression2.Parameters["t"] = t;
//            try
//            {
//                // Evaluate the expression
//                object result = expression.Evaluate();
//                object result2 = expression2.Evaluate();

//                var ty = t+(int)result;
//                // Print the result to the console
//                string message = $"Result of '{formulaString}' = {result}";
//                Console.WriteLine(message);

//                // Optionally, you can store the result and pass it to the View
//                ViewBag.ExpressionResult = message; // Store the result in a ViewBag

//                // Return a View with the result
//                return View();
//            }
//            catch (NCalc.EvaluationException ex)
//            {
//                // Handle specific NCalc evaluation exceptions
//                Console.WriteLine($"Error evaluating expression: {ex.Message}");

//                // Optionally, you can pass an error message to the View
//                ViewBag.ErrorMessage = $"Error evaluating expression: {ex.Message}";

//                // Return a View indicating an error occurred
//                return View("ErrorView"); // Create an ErrorView to display error messages
//            }
//            catch (Exception ex)
//            {
//                // Handle other generic exceptions
//                Console.WriteLine($"Unexpected error: {ex.Message}");

//                // Optionally, you can pass an error message to the View
//                ViewBag.ErrorMessage = $"Unexpected error: {ex.Message}";

//                // Return a View indicating an unexpected error occurred
//                return View("ErrorView"); // Create an ErrorView to display error messages
//            }
//        }



    }
}


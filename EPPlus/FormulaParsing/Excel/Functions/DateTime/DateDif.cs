using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime {
    public class DateDif : ExcelFunction {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context) {
            ValidateArguments(arguments, 3);

            var startDateArg = arguments.ElementAt(0);
            var endDateArg = arguments.ElementAt(1);
            var unitArg = arguments.ElementAt(2);

            var startDate = System.DateTime.FromOADate(ArgToDecimal(arguments, 0));
            var endDate = System.DateTime.FromOADate(ArgToDecimal(arguments, 1));
            var unit = ArgToString(arguments, 2).ToUpper();

            if (startDate > endDate) {
                return CreateResult(0, DataType.Integer); // Error
            }

            double result;

            switch (unit) {
                case "Y": // Years
                    result = CalculateYears(startDate, endDate);
                    break;
                case "M": // Months
                    result = CalculateMonths(startDate, endDate);
                    break;
                case "D": // Days
                    result = (endDate - startDate).TotalDays;
                    break;
                case "MD": // Days ignoring months and years
                    result = CalculateDaysIgnoringMonthsYears(startDate, endDate);
                    break;
                case "YM": // Months ignoring years
                    result = CalculateMonthsIgnoringYears(startDate, endDate);
                    break;
                case "YD": // Days ignoring years
                    result = CalculateDaysIgnoringYears(startDate, endDate);
                    break;
                default:
                    return CreateResult(0, DataType.Integer); // Error
            }

            return CreateResult(result, DataType.Integer);
        }

        private int CalculateYears(System.DateTime startDate, System.DateTime endDate) {
            int years = endDate.Year - startDate.Year;

            // Adjust if the end date hasn't reached the anniversary yet
            if (endDate.Month < startDate.Month ||
                (endDate.Month == startDate.Month && endDate.Day < startDate.Day)) {
                years--;
            }

            return years;
        }

        private int CalculateMonths(System.DateTime startDate, System.DateTime endDate) {
            int months = (endDate.Year - startDate.Year) * 12 + endDate.Month - startDate.Month;

            // Adjust if the end day hasn't reached the start day yet
            if (endDate.Day < startDate.Day) {
                months--;
            }

            return months;
        }

        private int CalculateDaysIgnoringMonthsYears(System.DateTime startDate, System.DateTime endDate) {
            // Calculate the difference in days within the same month context
            var tempStart = new System.DateTime(endDate.Year, endDate.Month, startDate.Day);

            if (tempStart > endDate) {
                // If the start day doesn't exist in the end month, use the last day of previous month
                tempStart = new System.DateTime(endDate.Year, endDate.Month, 1).AddDays(-1);
                tempStart = new System.DateTime(tempStart.Year, tempStart.Month, System.Math.Min(startDate.Day, System.DateTime.DaysInMonth(tempStart.Year, tempStart.Month)));
            }

            return (endDate - tempStart).Days;
        }

        private int CalculateMonthsIgnoringYears(System.DateTime startDate, System.DateTime endDate) {
            int months = endDate.Month - startDate.Month;

            if (months < 0) {
                months += 12;
            }

            // Adjust if the end day hasn't reached the start day yet
            if (endDate.Day < startDate.Day) {
                months--;
                if (months < 0) {
                    months += 12;
                }
            }

            return months;
        }

        private int CalculateDaysIgnoringYears(System.DateTime startDate, System.DateTime endDate) {
            // Create dates with the same year to calculate day difference
            var adjustedStart = new System.DateTime(endDate.Year, startDate.Month, startDate.Day);

            if (adjustedStart > endDate) {
                // If we're in the previous year cycle, subtract a year
                adjustedStart = adjustedStart.AddYears(-1);
            }

            return (endDate - adjustedStart).Days;
        }
    }
}

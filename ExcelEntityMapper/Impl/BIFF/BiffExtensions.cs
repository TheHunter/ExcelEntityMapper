using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using NPOI.SS.UserModel;

namespace ExcelEntityMapper.Impl.BIFF
{
    internal static class BiffExtensions
    {
        private static readonly Delegate numericDelegate;
        private static readonly Delegate stringDelegate;


        static BiffExtensions()
        {
            Type type = typeof (ICell);

            ParameterExpression instanceInvoker = Expression.Parameter(typeof(ICell), "cell");
            ParameterExpression numericParameter = Expression.Parameter(typeof(double), "value");
            MethodInfo stringMethod = type.GetMethod("SetCellValue", new[] { typeof(string) });

            ParameterExpression stringParameter = Expression.Parameter(typeof(string), "value");
            MethodInfo numericMethod = type.GetMethod("SetCellValue", new[] { typeof(double) });
            

            numericDelegate = Expression.Lambda<Action<ICell, double>>
                                            (
                                                Expression.Call(instanceInvoker, numericMethod, numericParameter),
                                                instanceInvoker,
                                                numericParameter
                                            )
                                            .Compile();

            stringDelegate = Expression.Lambda<Action<ICell, string>>
                                            (
                                                Expression.Call(instanceInvoker, stringMethod, stringParameter),
                                                instanceInvoker,
                                                stringParameter
                                            )
                                            .Compile();
        }

        #region IRow extension
        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        internal static int RowNumber(this IRow row)
        {
            return row.RowNum;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        internal static IRow RowBelow(this IRow row)
        {
            return row.RowBelow(1);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row"></param>
        /// <param name="step"></param>
        /// <returns></returns>
        internal static IRow RowBelow(this IRow row, int step)
        {
            return row.Sheet.GetRow(row.RowNum + step);
        }
        #endregion

        #region ISheet extension
        internal static IRow Row(this ISheet sheet, int row)
        {
            return sheet.GetRow(row) ?? sheet.CreateRow(row);
        }

        internal static ICell Cell(this IRow row, int columnIndex)
        {
            return row.GetCell(columnIndex) ?? row.CreateCell(columnIndex);
        }

        

        #endregion


        #region ICell extension
        internal static string GetString(this ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.NUMERIC:
                    return cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
                default:
                    return cell.StringCellValue.Trim();
            }
        }

        internal static bool IsEmpty(this ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.BLANK:
                case CellType.FORMULA:
                case CellType.ERROR:
                case CellType.Unknown:
                    return true;
                case CellType.STRING:
                    return string.IsNullOrEmpty(cell.StringCellValue.Trim());
                default:
                    return false;
            }
        }

        internal static void SetCellValue(this ICell cell, object value)
        {
            if (XLEntityHelper.IsNumericValue(value))
                numericDelegate.DynamicInvoke(cell, value);
            else
                stringDelegate.DynamicInvoke(cell, value);
        }

        #endregion
    }
}

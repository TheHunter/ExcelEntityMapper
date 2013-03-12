using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Linq.Expressions;
using SmartXLS;
using ExcelEntityMapper.Impl;

namespace ExcelEntityMapper
{
    /// <summary>
    /// 
    /// </summary>
    internal static class SourceHelper
    {

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="TSource"></typeparam>
        /// <returns></returns>
        internal static TSource CreateInstance<TSource>()
        {
            Type type = typeof(TSource);
            try
            {
                return (TSource)type.InvokeMember(null, BindingFlags.CreateInstance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, null, null);
            }
            catch (Exception ex)
            {
                throw new NotSupportedException(string.Format("Error into building a new object, type: {0}", type.Name), ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="expr"></param>
        /// <returns></returns>
        internal static KeyValuePair<string, Type> GetMemberInfo(System.Linq.Expressions.Expression expr)
        {
            string name = null;
            Type type = null;

            if (expr == null)
            {
                throw new NullReferenceException("Expression cannot be null.");
            }
            else
            {
                if (expr is MemberExpression)
                {
                    MemberExpression memberExpr = (MemberExpression)expr;
                    switch (memberExpr.Member.MemberType)
                    {
                        case System.Reflection.MemberTypes.Property:
                            {
                                PropertyInfo info = (PropertyInfo)memberExpr.Member;
                                type = info.PropertyType;
                                break;
                            }
                        case System.Reflection.MemberTypes.Field:
                            {
                                FieldInfo info = (FieldInfo)memberExpr.Member;
                                type = info.FieldType;
                                break;
                            }
                        default:
                            {
                                throw new ArgumentException("The member type is not a right expression.");
                            }
                    }
                    string temp = expr.ToString();
                    name = temp.Substring(temp.IndexOf(".") + 1);
                }
                else if (expr is LambdaExpression)
                {
                    return GetMemberInfo(((LambdaExpression)expr).Body);
                }
                //else if (expr is InvocationExpression)
                //{
                //    return GetMemberInfo(((InvocationExpression)expr).Expression);
                //}
                else if (expr is ConditionalExpression)
                {
                    return new KeyValuePair<string, Type>("[Conditional]", type);
                }
                else if (expr is MethodCallExpression)
                {
                    return GetMemberInfo(((MethodCallExpression)expr).Object);
                }
                else if (expr is UnaryExpression && expr.NodeType == ExpressionType.Convert)
                {
                    return GetMemberInfo(((UnaryExpression)expr).Operand);
                }
                else
                {
                    throw new ArgumentException("Expression cannot be used in this context.");
                }
                return new KeyValuePair<string, Type>(name, type);
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        internal static WorkBook VerifyWorkBook(IXLWorkBook workbook, string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
                throw new ArgumentException("It's impossible to read/write a sheet if the sheetname is not set.");

            if (workbook == null)
                throw new NotImplementedException("no workbook was referenced.");

            XWorkBook wb = workbook as XWorkBook;

            if (wb == null)
                throw new InvalidCastException("workbook type is not valid.");

            int index = wb.Workbook.findSheetByName(sheetName);

            if (index == -1)
            {
                throw new ArgumentException("There's no sheet with the name specified in argument.");
            }

            wb.Workbook.Sheet = index;
            return wb.Workbook;
        }

    }
}

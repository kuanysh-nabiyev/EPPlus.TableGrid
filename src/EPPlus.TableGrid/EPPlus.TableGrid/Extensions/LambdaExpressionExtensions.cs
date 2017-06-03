using System;
using System.Linq.Expressions;

namespace EPPlus.TableGrid.Extensions
{
    internal static class LambdaExpressionExtensions
    {
        public static string GetPropertyName<T, TKey>(this Expression<Func<T, TKey>> propertyExpression)
        {
            MemberExpression mbody = propertyExpression.Body as MemberExpression;

            if (mbody == null)
            {
                UnaryExpression ubody = propertyExpression.Body as UnaryExpression;
                if (ubody != null)
                {
                    mbody = ubody.Operand as MemberExpression;
                }

                if (mbody == null)
                {
                    throw new ArgumentException("Expression is not a MemberExpression", nameof(propertyExpression));
                }
            }

            return mbody.Member.Name;
        }
    }
}
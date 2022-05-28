using MgSoftDev.OXExcel.Commons;
using MgSoftDev.OXExcel.Entities.Table;
using MgSoftDev.OXExcel.Helpers.Extensions;

namespace MgSoftDev.OXExcel.OpenXmlProvider.Helpers.Extensions
{
    internal static class OxTableExtension
    {

        internal static bool HiddenForFilter(this object value, List<OxTableColumnsEntity> columns )
        {
            var fcols = columns.Where(w => w.CustomColumnFilter != null || w.ColumnFilter != null).ToList();
            if (fcols.Count == 0) return false;
            foreach (var c in fcols)
            {
                var val = value.GetPropertyVal(c.PropertyPath,"");
                if (c.ColumnFilter != null)
                {
                    if (!c.ColumnFilter.Exists(e => e.Val.Equals(val)))
                        return true;
                }
                else if (c.CustomColumnFilter != null)
                {
                    var res = true;
                    if (c.CustomColumnFilter.Condition == OxCustomFilterCondition.None)
                    {
                        if (c.CustomColumnFilter.Operator.EvaluateCondition(val, c.CustomColumnFilter.Val))
                            res = false;
                    }
                    if (c.CustomColumnFilter.Condition == OxCustomFilterCondition.And)
                    {
                        if (c.CustomColumnFilter.Operator.EvaluateCondition(val, c.CustomColumnFilter.Val) &&
                            c.CustomColumnFilter.Operator2.EvaluateCondition(val, c.CustomColumnFilter.Val2))
                            res = false;
                    }
                    if (c.CustomColumnFilter.Condition == OxCustomFilterCondition.Or)
                    {
                        if (c.CustomColumnFilter.Operator.EvaluateCondition(val, c.CustomColumnFilter.Val) ||
                            c.CustomColumnFilter.Operator2.EvaluateCondition(val, c.CustomColumnFilter.Val2))
                            res = false;
                    }
                    if(res)
                        return true;
                }
            }
            return false;
        }

        internal static bool EvaluateCondition(this OxFilterOperators value, string obj1, string obj2)
        {
            double num;
            double num2;
            switch (value)
            {
                case OxFilterOperators.Equal:
                    return obj1.Equals(obj2);
                case OxFilterOperators.LessThan:
                    if (double.TryParse(obj1, out num) && double.TryParse(obj2, out num2))
                        return num < num2;
                    break;
                case OxFilterOperators.LessThanOrEqual:
                    if (double.TryParse(obj1, out num) && double.TryParse(obj2, out num2))
                        return num <= num2;
                    break;
                case OxFilterOperators.NotEqual:
                    return !obj1.Equals(obj2);
                case OxFilterOperators.GreaterThanOrEqual:
                    if (double.TryParse(obj1, out num) && double.TryParse(obj2, out num2))
                        return num >= num2;
                    break;
                case OxFilterOperators.GreaterThan:
                    if (double.TryParse(obj1, out num) && double.TryParse(obj2, out num2))
                        return num > num2;
                    break;
                case OxFilterOperators.StartWith:
                    return obj1.StartsWith(obj2);
                case OxFilterOperators.EndWith:
                    return obj1.EndsWith(obj2);
                case OxFilterOperators.Contrains:
                    return obj1.Contains(obj2);
                case OxFilterOperators.NotContrains:
                    return !obj1.Contains(obj2);
                
            }
            return true;
        }


        internal static string ApplyOperator(this string value, OxFilterOperators filterOperators)
        {
            switch (filterOperators)
            {                
                case OxFilterOperators.StartWith:
                    return  value + "*";
                case OxFilterOperators.EndWith:
                    return "*" + value;
                case OxFilterOperators.Contrains:
                    return "*" + value +"*";
                case OxFilterOperators.NotContrains:
                    return "*" + value + "*";
            }
            return value;
        }

        internal static string GetSubTotalFormula(this OxTableColumnsEntity value)
        {
            var hidden = value.TotalRow.IncludeHidden;
            switch (value.TotalRow.RowFormula)
            {
                case TotalsRowFormulas.None:
                    return "";
                case TotalsRowFormulas.Sum:
                    return $"=SUBTOTAL({(hidden ? 9 : 109)},[{value.Header}])";
                case TotalsRowFormulas.Minimum:
                    return $"=SUBTOTAL({(hidden ? 5 : 105)},[{value.Header}])";
                case TotalsRowFormulas.Maximum:
                    return $"=SUBTOTAL({(hidden ? 4 : 104)},[{value.Header}])";
                case TotalsRowFormulas.Average:
                    return $"=SUBTOTAL({(hidden ? 1 : 101)},[{value.Header}])";
                case TotalsRowFormulas.Count:
                    return $"=SUBTOTAL({(hidden ? 2 : 102)},[{value.Header}])";
                case TotalsRowFormulas.CountNumbers:
                    return $"=SUBTOTAL({(hidden ? 3 : 103)},[{value.Header}])";
                case TotalsRowFormulas.StandardDeviation:
                    return $"=SUBTOTAL({(hidden ? 7 : 107)},[{value.Header}])";
                case TotalsRowFormulas.Variance:
                    return $"=SUBTOTAL({(hidden ? 10 : 110)},[{value.Header}])";
                case TotalsRowFormulas.Custom:
                    return value.TotalRow.CustomFormula;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }
    }
}

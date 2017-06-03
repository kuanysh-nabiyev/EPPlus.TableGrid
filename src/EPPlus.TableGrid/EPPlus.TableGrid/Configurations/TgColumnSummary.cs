using System;
using System.ComponentModel.DataAnnotations;
using EPPlus.TableGrid.Exceptions;
using EPPlus.TableGrid.Extensions;

namespace EPPlus.TableGrid.Configurations
{
    public class TgColumnSummary
    {
        /// <summary>aggreagate function (sum, count, average and etc.)</summary>
        public AggregateFunction AggregateFunction { get; set; }

        /// <summary>summary style</summary>
        public TgExcelStyle Style { get; set; }

        internal void Validate()
        {
            if (AggregateFunction == null)
                throw new RequiredPropertyException(nameof(AggregateFunction), GetType());
        }
    }

    public class AggregateFunction
    {
        public AggregateFunction(AggregateFunctionType type)
        {
            Type = type;
        }

        public AggregateFunction(AggregateFunctionType type, string condition)
        {
            if (!type.GetDisplayName().EndsWith("IF", StringComparison.OrdinalIgnoreCase))
            {
                throw new IncorrectTgOptionsException(
                    $"Cannot set condition property to {type.GetDisplayName()} function. " +
                     "Function must end with 'If' string, such as SumIf, CcountIf");
            }
            Type = type;
            Condition = condition;
        }

        /// <summary>Aggregate function type (sum, count, average and etc.)</summary>
        public AggregateFunctionType Type { get; private set; }

        /// <summary>Condition (criteria)</summary>
        public string Condition { get; private set; }

        internal bool HasCondition => Condition != null;
    }

    public enum AggregateFunctionType
    {
        /// <summary>
        /// counts the number of cells in a range that contain members
        /// </summary>
        [Display(Name = "COUNT")]
        Count,
        /// <summary>
        /// counts the number of cells in a range that are not empty
        /// </summary>
        [Display(Name = "COUNTA")]
        CountA,
        /// <summary>
        /// counts the number of empty cells in a range
        /// </summary>
        [Display(Name = "COUNTBLANK")]
        CountBlank,
        /// <summary>
        /// counts the number of cells in a range that meet the given condition
        /// </summary>
        [Display(Name = "COUNTIF")]
        CountIf,
        /// <summary>
        /// returns the average (arithmetic mean) of its arguments, 
        /// which can be number or names, arrays, or references that contain numbers
        /// </summary>
        [Display(Name = "AVERAGE")]
        Average,
        /// <summary>
        /// returns the average (arithmetic mean) of its arguments, 
        /// evaluating text and FALSE in arguments as 0; TRUE evaluates as 1.
        /// Arguments can be number or names, arrays, or references
        /// </summary>
        [Display(Name = "AVERAGEA")]
        AverageA,
        /// <summary>
        /// returns the largest value in a set of values. Ignores logical values and text
        /// </summary>
        [Display(Name = "MAX")]
        Max,
        /// <summary>
        /// returns the smallest value in a set of values. Ignores logical values and text
        /// </summary>
        [Display(Name = "MIN")]
        Min,
        /// <summary>
        /// calculates standard deviation based on the entire population given as arguments (ignores logical values and text)
        /// </summary>
        [Display(Name = "STDEV.P")]
        StDevP,
        /// <summary>
        /// estimates standard deviation based on sample (ignores logical values and text)
        /// </summary>
        [Display(Name = "STDEV.S")]
        StDevS,
        /// <summary>
        /// Adds all the numbers in a range of cells 
        /// </summary>
        [Display(Name = "SUM")]
        Sum,
        /// <summary>
        /// adds the cells by a given condition
        /// </summary>
        [Display(Name = "SUMIF")]
        SumIf,
        /// <summary>
        /// calculates variance based on the entire population given as arguments (ignores logical values and text)
        /// </summary>
        [Display(Name = "VAR.P")]
        VarP,
        /// <summary>
        /// estimates variance based on sample (ignores logical values and text)
        /// </summary>
        [Display(Name = "VAR.S")]
        VarS,
    }
}
using System;
using System.Linq.Expressions;
using EPPlus.TableGrid.Core.Exceptions;
using EPPlus.TableGrid.Core.Extensions;

namespace EPPlus.TableGrid.Core.Configurations
{
    public class TgGroupOptions<T>
    {
        /// <summary>group by column</summary>
        public Expression<Func<T, object>> GroupingColumn { get; set; }

        /// <summary>grouping type</summary>
        public GroupingType GroupingType { get; set; }

        /// <summary>is group collapsable</summary>
        public bool IsGroupCollapsable { get; set; }

        internal string GetGroupingColumnName()
        {
            return GroupingColumn.GetPropertyName();
        }

        internal void Validate()
        {
            if (GroupingColumn == null)
                throw new RequiredPropertyException(nameof(GroupingColumn), this.GetType());
        }
    }

    public enum GroupingType
    {
        GroupHeaderOnColumn,
        GroupHeaderOnRow
    }
}
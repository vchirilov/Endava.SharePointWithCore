using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SharePointWithCore
{
    public static class Extensions
    {
        public static IEnumerable<IEnumerable<T>> Batch<T>(this IEnumerable<T> items,  int size)
        {
            return items
                .Select((item, inx) => new { item, inx })
                .GroupBy(x => x.inx / size)
                .Select(g => g.Select(x => x.item));
        }

        public static IEnumerable<int> ToArray(this int value)
        {
            int[] array = new int[value];

            for(var i=1; i<=value; i++)
                array[i - 1] = i;

            return array;
        }
    }
}

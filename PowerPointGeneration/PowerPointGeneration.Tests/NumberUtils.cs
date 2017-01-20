using System;
using System.Collections;
using System.Collections.Generic;

namespace PowerPointGeneration.Tests
{
    public static class NumberUtils
    {
        public static bool NextBool(this Random r, int truePercentage = 50)
        {
            return r.NextDouble() < truePercentage / 100.0;
        }

        public static T RemoveFirst<T>(this IList<T> list)
        {
            if (list.Count <= 0) { throw new Exception("Can RemoveFirst from an Empty list");}
            var e = list[0];
            list.RemoveAt(0);
            return e;
        }
    }
}
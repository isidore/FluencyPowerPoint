using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace PowerPointGeneration.Tests
{
    public class Timings : IEnumerable
    {
        private int last = 0;
        private List<Tuple<int, int, float>> times = new List<Tuple<int,int, float>>(); 
        public void Add(int index, float timing)
        {
            times.Add(new Tuple<int, int, float>(last, index, timing));
            last = index;
        }

        public float Get(int counter)
        {
            return times.First(t => t.Item1 <= counter && counter < t.Item2).Item3;
        }

        public IEnumerator GetEnumerator()
        {
            throw new NotImplementedException();
        }
    }
}
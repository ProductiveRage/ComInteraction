using System;
using System.Collections.Generic;

namespace COMInteraction.Misc
{
    public class NonNullImmutableList<T> : IEnumerable<T> where T : class
    {
        private List<T> _data;
        public NonNullImmutableList()
        {
            _data = new List<T>();
        }
        public NonNullImmutableList(IEnumerable<T> values)
        {
            if (values == null)
                throw new ArgumentNullException("values");
            var data = new List<T>();
            foreach (var value in values)
            {
                if (value == null)
                    throw new ArgumentException("Null entry encountered in values");
                data.Add(value);
            }
            _data = data;
        }

        public NonNullImmutableList<T> Add(T value)
        {
            if (value == null)
                throw new ArgumentNullException("value");
            var dataNew = new List<T>(_data);
            dataNew.Add(value);
            return new NonNullImmutableList<T>()
            {
                _data = dataNew
            };
        }

        public NonNullImmutableList<T> RemoveAt(int index)
        {
            if ((index < 0) || (index >= _data.Count))
                throw new ArgumentOutOfRangeException("index");
            var dataNew = new List<T>(_data);
            dataNew.RemoveAt(index);
            return new NonNullImmutableList<T>()
            {
                _data = dataNew
            };
        }

        public bool Contains(T value)
        {
            if (value == null)
                throw new ArgumentNullException("value");
            return _data.Contains(value);
        }

        public IEnumerator<T> GetEnumerator()
        {
            return _data.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}

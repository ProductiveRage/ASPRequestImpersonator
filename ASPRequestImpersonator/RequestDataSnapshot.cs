using System;
using System.Collections.Generic;

namespace ASPRequestImpersonator
{
    public class RequestDataSnapshot
    {
        private Dictionary<string, RequestStringList> _data;
        public RequestDataSnapshot(IEnumerable<KeyValuePair<string, RequestStringList>> data)
        {
            if (data == null)
                throw new ArgumentNullException("data");

            var cleanedData = new Dictionary<string, RequestStringList>(StringComparer.CurrentCultureIgnoreCase);
            foreach (var entry in data)
            {
                if (entry.Key == null)
                    throw new ArgumentException("Null key encountered in data");
                if (cleanedData.ContainsKey(entry.Key))
                    throw new ArgumentException("Duplicate key encountered in data");
                cleanedData.Add(entry.Key, entry.Value);
            }
            _data = cleanedData;
        }

        /// <summary>
        /// Key matching is case insensitive
        /// </summary>
        public ICollection<string> Keys
        {
            get { return _data.Keys; }
        }

        /// <summary>
        /// An exception will be raised for a null key, but not for an empty one - the ASP Request object supports keys that are constructed solely of whitespace
        /// (though won't contain any that are a blank string).
        /// </summary>
        public RequestStringList this[string key]
        {
            get
            {
                if (key == null)
                    throw new ArgumentNullException("key");
                if (!_data.ContainsKey(key))
                    throw new ArgumentOutOfRangeException("key");
                return _data[key];
            }
        }
    }
}

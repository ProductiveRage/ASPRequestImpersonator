using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ASPRequestImpersonator
{
    /// <summary>
    /// This provides an object that can be returned from Request.Querystring(key) calls (for example), that VBScript can either use immediately as a string (via the default
    /// Value property) or which it can loop round for cases where it must consider multiple values for a key.
    /// </summary>
    [ComVisible(true)]
    public class RequestStringList
    {
        private List<string> _values;
        public RequestStringList(IEnumerable<string> values)
        {
            if (values == null)
                throw new ArgumentNullException("values");

            var valueList = new List<string>();
            foreach (var value in values)
            {
                if (value == null)
                    throw new ArgumentException("Null entry encountered in values");
                valueList.Add(value);
            }
            _values = valueList;
        }

        /// <summary>
        /// Default value is a string or comma-separated values (values will not be UrlEncoded)
        /// </summary>
        [DispId(0)]
        public string Value
        {
            get { return string.Join(", ", _values.ToArray()); }
        }

        // ICollection methods we need for VBScript enumeration (values will not be UrlEncoded)
        public IEnumerator GetEnumerator()
        {
            return _values.GetEnumerator();
        }
        public int Count
        {
            get { return _values.Count; }
        }
    }
}

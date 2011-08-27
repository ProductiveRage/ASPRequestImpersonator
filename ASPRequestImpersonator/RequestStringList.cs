using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ASPRequestImpersonator
{
    /// <summary>
    /// This provides an object that can be returned from Request.QueryString(key) calls (for example), that VBScript can either use immediately as a string (via the default
    /// Value property) or which it can loop round for cases where it must consider multiple values for a key.
    /// </summary>
    [ComVisible(true)]
	public class RequestStringList : IManagedRequestStringList
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
        /// Default value is a string or comma-separated values (values will not be UrlEncoded).
		/// Note: VBScript seems to handle a "Value" method or property specially and use it as default, so DispId(0) doesn't need to be set for that case, but it
		/// makes sense to explicitly mark it as default using the appropriate Dispatch Id.
		/// </summary>
        [DispId(0)]
        public string Value
        {
            get { return string.Join(", ", _values.ToArray()); }
        }

		/// <summary>
		/// VBScript will enumerate over this data even without the DispId(-4) attribute, but it's technically correct to mark the enumerator method with a -4 Dispatch
		/// Id and it has no negative effect
		/// </summary>
		[DispId(-4)]
		public IEnumerator GetEnumerator()
        {
            return _values.GetEnumerator();
        }
        public int Count
        {
            get { return _values.Count; }
        }

		IEnumerator<string> IEnumerable<string>.GetEnumerator()
		{
			return _values.GetEnumerator();
		}
	}
}

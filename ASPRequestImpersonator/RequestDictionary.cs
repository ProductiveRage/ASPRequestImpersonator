using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Web;

namespace ASPRequestImpersonator
{
    /// <summary>
    /// Return a dictionary-like object that can be enumerated over (which runs through the Keys) and which will then respond to calls to the "Item" method for Value data,
    /// returning an object that can either be displayed directly as a string (via RequestStringList's default Value property). Note that the default method in this class
    /// is GetSummary which returns a string combining all of the data, the RequestImpersonator class has to implement a specific QueryString(string) method to return
    /// items data for a Key when the RequestDictionary.Item method is not called.
    /// 
    /// To summarise:
    ///  A call to Request.QueryString.Key(key) will be handled by the RequestDictionary.Item(key) method
    ///  A call to Request.QueryString(key) will be handled by the RequestImpersonator.QueryString(key) method
    /// </summary>
    [ComVisible(true)]
	public class RequestDictionary : IManagedRequestDictionary
    {
        private RequestDataSnapshot _data;
        public RequestDictionary(RequestDataSnapshot data)
        {
            if (data == null)
                throw new ArgumentNullException("data");
            _data = data;
        }

		/// <summary>
		/// VBScript will enumerate over this data even without the DispId(-4) attribute, but it's technically correct to mark the enumerator method with a -4 Dispatch
		/// Id and it has no negative effect
		/// </summary>
		[DispId(-4)]
		public IEnumerator GetEnumerator()
        {
            return _data.Keys.GetEnumerator();
        }
        public int Count
        {
            get { return _data.Keys.Count; }
        }

        // Additional data we want to emulate a Request dictionary
        public RequestDictionary Keys
        {
            get { return this; }
        }
		public RequestStringList Item(string key)
        {
			if (!_data.Keys.Contains(key))
                return new RequestStringList(new string[0]);
			return _data[key];
        }

		// Implement managed interface
		IEnumerator<string> IEnumerable<string>.GetEnumerator()
		{
			return _data.Keys.GetEnumerator();
		}
		IManagedRequestStringList IManagedRequestDictionary.this[string key]
		{
			get { return Item(key); }
		}

		/// <summary>
        /// Combine the values into a single summary - the values will be UrlEncoded (this is used - for example - to when Request.QueryString is called and the values
		/// are returned in one url-safe string).
		/// Note: VBScript seems to handle a "Value" method or property specially and use it as default, so DispId(0) doesn't need to be set for that case, but it
		/// makes sense to explicitly mark it as default using the appropriate Dispatch Id.
        /// </summary>
        [DispId(0)]
		public string Value
        {
			get
			{
				var content = new StringBuilder();
				foreach (var key in _data.Keys)
				{
					foreach (string value in _data[key])
					{
						if (content.Length > 0)
							content.Append("&");
						content.AppendFormat("{0}={1}", HttpUtility.UrlEncode(key), HttpUtility.UrlEncode(value));
					}
				}
				return content.ToString();
			}
        }
	}
}

using System;
using System.Collections;
using System.Runtime.InteropServices;
using System.Text;
using System.Web;

namespace ASPRequestImpersonator
{
    /// <summary>
    /// Return a dictionary-like object that can be enumerated over (which runs through the Keys) and which will then respond to calls to the "Item" method for Value data,
    /// returning an object that can either be displayed directly as a string (via RequestStringList's default Value property). Note that the default method in this class
    /// is GetSummary which returns a string combining all of the data, the RequestImpersonator class has to implement a specific Querystring(string) method to return
    /// items data for a Key when the RequestDictionary.Item method is not called.
    /// 
    /// To summarise:
    ///  A call to Request.Querystring.Key(key) will be handled by the RequestDictionary.Item(key) method
    ///  A call to Request.Querystring(key) will be handled by the RequestImpersonator.Querystring(key) method
    /// </summary>
    [ComVisible(true)]
    public class RequestDictionary
    {
        private RequestDataSnapshot _data;
        public RequestDictionary(RequestDataSnapshot data)
        {
            if (data == null)
                throw new ArgumentNullException("data");
            _data = data;
        }

        // ICollection methods we need for VBScript enumeration
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
        public RequestStringList Item(string name)
        {
            if (!_data.Keys.Contains(name))
                return new RequestStringList(new string[0]);
            return _data[name];
        }

        /// <summary>
        /// Combine the values into a single summary - note: the values will be UrlEncoded (this is used - for example - to when Request.Querystring is called
        /// and the values are returned in one url-safe string)
        /// </summary>
        [DispId(0)]
        public string GetSummary()
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

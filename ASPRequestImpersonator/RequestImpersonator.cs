using System;
using System.Runtime.InteropServices;

namespace ASPRequestImpersonator
{
	/// <summary>
    /// We need to wrap this using a LateBindingComWrapper so that COM can interact with the overloaded methods. We need the separate methods signatures to enable all of the
    /// following forms to be possible:
    /// 
    ///  Request.Querystring
    ///   - Returns a url-encoded string containing all of the querystring data, handled by RequestImpersonator.Querystring()
    ///   
    ///  Request.Querystring(key) 
    ///   - Returns a string that combines the values for that key, handled by RequestImpersonator.Querystring(key)
    ///   
    ///  For .. in Request.Querystring(key)
    ///   - Enumerates through values for that key, handled by RequestImpersonator.Querystring(key)
    ///   
    ///  Request.Querystring.Item(key) 
    ///   - Returns a string that combines the values for that key, handled by RequestImpersonator.Querystring()
    ///   
    ///  For .. in Request.Querystring.Item(key)
    ///   - Enumerates through values for that key, handled by RequestImpersonator.Querystring()
    ///   
    /// An indexed property is implemented - for requests of the form Request(key) or Request.Item(key) - which considers data from Querystring, Form, and ServerVariables
    /// (in that order). It wil never combine values for the same key from the different sets but if there is no data for a specified key in Querystring but there IS data
    /// in Form, then it will return data from Form.
    /// </summary>
	[ComVisible(true)]
    public class RequestImpersonator
	{
        private RequestDataSnapshot _formData, _querystringData, _serverVariablesData;
        public RequestImpersonator(RequestDataSnapshot formData, RequestDataSnapshot querystringData, RequestDataSnapshot serverVariablesData)
        {
            if (formData == null)
                throw new ArgumentNullException("formData");
            if (querystringData == null)
                throw new ArgumentNullException("querystringData");
            if (serverVariablesData == null)
                throw new ArgumentNullException("serverVariablesData");

            _formData = formData;
            _querystringData = querystringData;
            _serverVariablesData = serverVariablesData;
        }

        public RequestDictionary Querystring()
        {
            return new RequestDictionary(_querystringData);
        }
        public RequestStringList Querystring(string name)
        {
            if (name == null)
                throw new ArgumentNullException("name");
            if (!_querystringData.Keys.Contains(name))
                return new RequestStringList(new string[0]);
            return _querystringData[name];
        }

        public RequestDictionary Form()
        {
            return new RequestDictionary(_formData);
        }
        public RequestStringList Form(string name)
        {
            if (name == null)
                throw new ArgumentNullException("name");
            if (!_formData.Keys.Contains(name))
                return new RequestStringList(new string[0]);
            return _formData[name];
        }

        public RequestDictionary ServerVariables()
        {
            return new RequestDictionary(_serverVariablesData);
        }
        public RequestStringList ServerVariables(string name)
        {
            if (name == null)
                throw new ArgumentNullException("name");
            if (!_serverVariablesData.Keys.Contains(name))
                return new RequestStringList(new string[0]);
            return _serverVariablesData[name];
        }

        /// <summary>
        /// Try to retrieve data from the internal lists - Querystring takes precedence over Form which takes precedence over ServerVariables. An exception will
        /// be thrown for null name argument. Requests with name values that are not present in any list will receive an empty RequestStringList.
        /// </summary>
        public RequestStringList this[string name]
        {
            get
            {
                if (name == null)
                    throw new ArgumentNullException("name");
                if (_querystringData.Keys.Contains(name))
                    return _querystringData[name];
                if (_formData.Keys.Contains(name))
                    return _formData[name];
                if (_serverVariablesData.Keys.Contains(name))
                    return _serverVariablesData[name];
                return new RequestStringList(new string[0]);
            }
        }
	}
}

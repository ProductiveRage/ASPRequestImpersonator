using System;
using System.Runtime.InteropServices;

namespace ASPRequestImpersonator
{
	/// <summary>
    /// We need to wrap this using a LateBindingComWrapper so that COM can interact with the overloaded methods. We need the separate methods signatures to enable all of the
    /// following forms to be possible:
    /// 
    ///  Request.QueryString
    ///   - Returns a url-encoded string containing all of the querystring data, handled by RequestImpersonator.QueryString()
    ///   
    ///  Request.QueryString(key) 
    ///   - Returns a string that combines the values for that key, handled by RequestImpersonator.QueryString(key)
    ///   
    ///  For .. in Request.QueryString(key)
    ///   - Enumerates through values for that key, handled by RequestImpersonator.QueryString(key)
    ///   
    ///  Request.QueryString.Item(key) 
    ///   - Returns a string that combines the values for that key, handled by RequestImpersonator.QueryString()
    ///   
    ///  For .. in Request.QueryString.Item(key)
    ///   - Enumerates through values for that key, handled by RequestImpersonator.QueryString()
    ///   
    /// An indexed property is implemented - for requests of the form Request(key) or Request.Item(key) - which considers data from QueryString, Form, and ServerVariables
    /// (in that order). It wil never combine values for the same key from the different sets but if there is no data for a specified key in QueryString but there IS data
    /// in Form, then it will return data from Form.
    /// 
    /// Note: This doesn't need to be ComVisible since we're never returning an instance of it through COM, only one wrapped in a LateBindingComWrapper.
    /// </summary>
    public class RequestImpersonator : IManagedRequestImpersonator
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

        public RequestDictionary Form()
        {
            return new RequestDictionary(_formData);
        }

		/// <summary>
		/// This will raise an exception for null key value (to be consistent with the ASP Request)
		/// </summary>
		public RequestStringList Form(string key)
        {
            if (key == null)
				throw new ArgumentNullException("key");
			if (!_formData.Keys.Contains(key))
                return new RequestStringList(new string[0]);
			return _formData[key];
        }

		public RequestDictionary QueryString()
		{
			return new RequestDictionary(_querystringData);
		}

		/// <summary>
		/// This will raise an exception for null key value (to be consistent with the ASP Request)
		/// </summary>
		public RequestStringList QueryString(string key)
		{
			if (key == null)
				throw new ArgumentNullException("key");
			if (!_querystringData.Keys.Contains(key))
				return new RequestStringList(new string[0]);
			return _querystringData[key];
		}

		public RequestDictionary ServerVariables()
        {
            return new RequestDictionary(_serverVariablesData);
        }
		
		/// <summary>
		/// This will raise an exception for null key value (to be consistent with the ASP Request)
		/// </summary>
		public RequestStringList ServerVariables(string key)
        {
            if (key == null)
				throw new ArgumentNullException("key");
			if (!_serverVariablesData.Keys.Contains(key))
                return new RequestStringList(new string[0]);
			return _serverVariablesData[key];
        }

        /// <summary>
        /// Try to retrieve data from the internal lists - QueryString takes precedence over Form which takes precedence over ServerVariables. Requests with name values
		/// that are not present in any list will receive an empty RequestStringList (this includes the case of a null key, to be consistent with the ASP Request)
        /// </summary>
		public RequestStringList this[string key]
        {
            get
            {
				if (key == null)
					return new RequestStringList(new string[0]);
				if (_querystringData.Keys.Contains(key))
					return _querystringData[key];
				if (_formData.Keys.Contains(key))
					return _formData[key];
				if (_serverVariablesData.Keys.Contains(key))
					return _serverVariablesData[key];
                return new RequestStringList(new string[0]);
            }
        }

		// Implement managed interface
		IManagedRequestStringList IManagedRequestImpersonator.this[string key]
		{
			get { throw new NotImplementedException(); }
		}
		IManagedRequestDictionary IManagedRequestImpersonator.Form
		{
			get { return Form(); }
		}
		IManagedRequestDictionary IManagedRequestImpersonator.QueryString
		{
			get { return QueryString(); }
		}
		IManagedRequestDictionary IManagedRequestImpersonator.ServerVariables
		{
			get { return ServerVariables(); }
		}
	}
}

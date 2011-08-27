using System;
using System.Runtime.InteropServices;

namespace ASPRequestImpersonator
{
    /// <summary>
	/// This is the class that is returned through COM that impersonates the ASP Request object. Having it also implement IManagedRequestImpersonator means that managed
	/// code can access the data easily as well.
    /// </summary>
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ComVisible(true)]
	public class RequestImpersonatorCom : LateBindingComWrapper, IManagedRequestImpersonator
    {
		private RequestImpersonator _requestImpersonator;
        public RequestImpersonatorCom(RequestDataSnapshot formData, RequestDataSnapshot querystringData, RequestDataSnapshot serverVariablesData)
            : base(new RequestImpersonator(formData, querystringData, serverVariablesData))
		{
			_requestImpersonator = new RequestImpersonator(formData, querystringData, serverVariablesData);
		}

		/// <summary>
		/// This will never return null
		/// </summary>
		public IManagedRequestDictionary Form
		{
			get { return _requestImpersonator.Form(); }
		}

		/// <summary>
		/// This will never return null
		/// </summary>
		public IManagedRequestDictionary QueryString
		{
			get { return _requestImpersonator.QueryString(); }
		}

		/// <summary>
		/// This will never return null
		/// </summary>
		public IManagedRequestDictionary ServerVariables
		{
			get { return _requestImpersonator.ServerVariables(); }
		}

		/// <summary>
		/// Try to retrieve data from the internal lists - QueryString takes precedence over Form which takes precedence over ServerVariables. An exception will
		/// be thrown for null name argument. Requests with name values that are not present in any list will receive an empty RequestStringList.
		/// </summary>
		public RequestStringList this[string key]
		{
			get { return _requestImpersonator[key]; }
		}

		IManagedRequestStringList IManagedRequestImpersonator.this[string key]
		{
			get { return this[key]; }
		}
	}
}

using System.Runtime.InteropServices;

namespace ASPRequestImpersonator
{
    /// <summary>
    /// This is the class that is returned through COM that impersonates the ASP Request object
    /// </summary>
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ComVisible(true)]
    public class RequestImpersonatorCom : LateBindingComWrapper
    {
        public RequestImpersonatorCom(RequestDataSnapshot formData, RequestDataSnapshot querystringData, RequestDataSnapshot serverVariablesData)
            : base(new RequestImpersonator(formData, querystringData, serverVariablesData)) { }
    }
}

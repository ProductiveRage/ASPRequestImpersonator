using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ASPRequestImpersonator
{
    /// <summary>
    /// Wrap an object in an IReflect implementation such that method and properties calls are passed through to the object in such a manner that methods with different
    /// signatures can be supported - ordinarily we would not be able to support this when accessing from VBScript through COM.
    /// </summary>
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ComVisible(true)]
    public class LateBindingComWrapper : IReflect
	{
		private object _target;
		public LateBindingComWrapper(object target)
		{
			if (target == null)
				throw new ArgumentNullException("target");

			_target = target;
		}

        public object InvokeMember(string name, BindingFlags invokeAttr, Binder binder, object target, object[] args, ParameterModifier[] modifiers, CultureInfo culture, string[] namedParameters)
		{
            return _target.GetType().InvokeMember(name, invokeAttr, binder, _target, args, modifiers, culture, namedParameters);
		}

        public FieldInfo GetField(string name, BindingFlags bindingAttr)
        {
            return _target.GetType().GetField(name, bindingAttr);
        }

        public FieldInfo[] GetFields(BindingFlags bindingAttr)
        {
            return _target.GetType().GetFields(bindingAttr);
        }

        public MemberInfo[] GetMember(string name, BindingFlags bindingAttr)
        {
            return _target.GetType().GetMember(name, bindingAttr);
        }

        public MemberInfo[] GetMembers(BindingFlags bindingAttr)
        {
            return _target.GetType().GetMembers(bindingAttr);
        }

        public MethodInfo GetMethod(string name, BindingFlags bindingAttr)
		{
			return _target.GetType().GetMethod(name, bindingAttr);
		}

        public MethodInfo GetMethod(string name, BindingFlags bindingAttr, Binder binder, Type[] types, ParameterModifier[] modifiers)
		{
			return _target.GetType().GetMethod(name, bindingAttr, binder, types, modifiers);
		}
		
        public MethodInfo[] GetMethods(BindingFlags bindingAttr)
		{
			return _target.GetType().GetMethods();
		}
		
        public PropertyInfo GetProperty(string name, BindingFlags bindingAttr)
		{
			return _target.GetType().GetProperty(name, bindingAttr);
		}
		
        public PropertyInfo GetProperty(string name, BindingFlags bindingAttr, Binder binder, Type returnType, Type[] types, ParameterModifier[] modifiers)
		{
			return _target.GetType().GetProperty(name, bindingAttr, binder, returnType, types, modifiers);
		}
		
        public PropertyInfo[] GetProperties(BindingFlags bindingAttr)
		{
			return _target.GetType().GetProperties();
		}

        public Type UnderlyingSystemType
        {
            get { return _target.GetType().UnderlyingSystemType; }
        }
    }
}

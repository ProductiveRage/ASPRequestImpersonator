# ASP Request Impersonator

## Overview

A class intended to be used by COM components (specifically VBScript WSC controls in one particular case) that require a reference to the ASP Request object - this class implements the QueryString, Form and ServerVariables collections so that a real Request reference isn't required to test the controls.

This is awkward because there are many ways in which these collections can be accessed.

eg. given the querystring "?key1=value+1&key2=value+2&key3=value+3", it can be accessed in the following ways -

    Request.QueryString

Returns a url-encoded string containing all of the querystring data.
     
    Request.QueryString.Count
    Request.QueryString.Keys.Count

Return the number of keys in the collection.
     
    Request.QueryString(key) 
    Request.QueryString.Item(key) 

Returns a string that combines the values for that key (eg. Request.QueryString("key2") returns "value 2, value 3".

    For .. in Request.QueryString
    For .. in Request.QueryString.Keys

Enumerates through the keys.
     
    For .. in Request.QueryString(key)
    For .. in Request.QueryString.Item(key)

Enumerates through values for the specified key - eg. "For .. in Request.QueryString("key2") loops over "value 2" and "value 3".

The value returned from a Request.QueryString(key) call against an ASP Request object is an IStringList, though this has a default method that returns a string (displaying as a comma separated list, if there are multiple values for the key).

## The problem

The difficulty with writing a COM component to imitate the Request object is that effectively a property or method with an optional argument is required to enable the various calls, particularly:

    Request.QueryString(key) 
    Request.QueryString.Item(key) 

But this isn't easy and I'm not sure it's even possible without the optional parameter support in .Net 4.0 - in fact, I failed to make it work even with that. Others have tried to solve this Request-mimicking problem on StackOverflow ([here](http://stackoverflow.com/questions/317759/why-is-the-indexer-on-my-net-component-not-always-accessible-from-vbscript)) but I didn't find that the accepted answer worked.

## The solution

This also comes from the StackOverflow thread - using the IReflect and ClassInterfaceType.AutoDispatch to wrap access to a class that has method signatures

    public RequestDictionary Querystring()
    public RequestStringList Querystring(string key)

and have them both accessible through the COM component.
# ASP Request Impersonator

A class intended to be used by VBScript WSC controls (or other COM components) that requires a reference to the ASP Request object - this class implements the QueryString, Form and ServerVariables collections so that a real Request reference isn't required to test the controls.
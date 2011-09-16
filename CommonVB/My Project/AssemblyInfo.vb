Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security.Permissions

' General Information about an assembly is controlled through the following 
' set of attributes. Change these attribute values to modify the information
' associated with an assembly.

' Review the values of the assembly attributes

<Assembly: AssemblyTitle("Resource Refactoring Tool - Common Libraries")> 
<Assembly: AssemblyDescription("")> 
<Assembly: AssemblyCompany("Microsoft Corporation")> 
<Assembly: AssemblyProduct("Power Toys for Visual Studio - Resource Refactoring Tool")> 
<Assembly: AssemblyCopyright("Copyright © Microsoft Corporation 2006")> 
<Assembly: AssemblyTrademark("")> 

<Assembly: ComVisible(False)>

'The following GUID is for the ID of the typelib if this project is exposed to COM
<Assembly: Guid("9fdfc109-0f40-49ce-ae9e-79c8c6c06102")> 

' Version information for an assembly consists of the following four values:
'
'      Major Version
'      Minor Version 
'      Build Number
'      Revision
'
' You can specify all the values or you can default the Build and Revision Numbers 
' by using the '*' as shown below:
' <Assembly: AssemblyVersion("1.0.*")> 

<Assembly: AssemblyVersion("2.1.0.0")> 
<Module: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA2210:AssembliesShouldHaveValidStrongNames")> 
'<Assembly: SecurityPermission(SecurityAction.RequestMinimum, Execution:=True)> 

' Module is not CLS Complient because we are using VS Extensibility assembilies.
<Assembly: CLSCompliant(False)> 
<Module: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1020:AvoidNamespacesWithFewTypes", Scope:="namespace", Target:="Microsoft.VSPowerToys.ResourceRefactor.Common")> 
Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security.Permissions


<Assembly: ComVisible(False)>

<Assembly: AssemblyTitle("Visual Studio Resource Refactoring Tool - VB part")> 

<Assembly: Guid("9fdfc109-0f40-49ce-ae9e-79c8c6c06102")> 

<Module: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA2210:AssembliesShouldHaveValidStrongNames")> 

' Module is not CLS Complient because we are using VS Extensibility assembilies.
<Assembly: CLSCompliant(False)> 
<Module: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1020:AvoidNamespacesWithFewTypes", Scope:="namespace", Target:="Microsoft.VSPowerToys.ResourceRefactor.Common")> 
Imports System.Reflection
Imports System.Runtime.InteropServices

<Assembly: AssemblyTitle("Visual Studio Resource Refactoring Tool - User Interface")> 
<Assembly: AssemblyDescription("User interface library")> 

<Assembly: Guid("14ebbcaa-04db-4177-9569-dc8f61a58eaa")> 
<Module: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA2210:AssembliesShouldHaveValidStrongNames")> 
' Module is not CLS Complient because we are using VS Extensibility assembilies.
<Assembly: CLSCompliant(False)> 
<Assembly: ComVisible(False)> 
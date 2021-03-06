using System.Reflection;
using System.Runtime.CompilerServices;
using System;
using System.Security.Permissions;
using System.Diagnostics.CodeAnalysis;

[assembly: AssemblyTitle("Power Toys for Visual Studio - Resource Refactoring Tool")]

[module: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA2210:AssembliesShouldHaveValidStrongNames")]

// Module is not CLS Complient because we are using VS Extensibility assembilies.
[assembly: CLSCompliant(false)]

[assembly: System.Runtime.InteropServices.ComVisible(false)]
[module: SuppressMessage("Microsoft.Naming", "CA1703:ResourceStringsShouldBeSpelledCorrectly", Scope = "resource", Target = "Microsoft.VSPowerToys.ResourceRefactor.CommandBar.resources")]
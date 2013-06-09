using System.Reflection;
using System.Runtime.InteropServices;

[assembly: AssemblyTitle("Resource Refactoring Tool - Common Libraries (C# part)")]
[assembly: ComVisible(false)]

[assembly: Guid("219f64e3-8a9a-42b6-8a29-0d703e13e242")]

[module: System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA2210:AssembliesShouldHaveValidStrongNames")]

// Module is not CLS Complient because we are using VS Extensibility assembilies.
[assembly: System.CLSCompliant(false)]
Imports System.Diagnostics.CodeAnalysis
Imports System.Runtime.InteropServices

''' <summary>
''' Contains information about paragraph formatting attributes in a rich edit control.
''' PARAFORMAT2 is a Microsoft Rich Edit 2.0 extension of the PARAFORMAT structure.
''' Microsoft Rich Edit 2.0 allows you to use either structure with the EM_GETPARAFORMAT and EM_SETPARAFORMAT messages.
''' </summary>
<StructLayout(LayoutKind.Sequential)> _
<SuppressMessage("Microsoft.Naming", "CA1705:LongAcronymsShouldBePascalCased")> _
<SuppressMessage("Microsoft.Performance", "CA1815:OverrideEqualsAndOperatorEqualsOnValueTypes")> _
Public Structure PARAFORMAT2
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Public cbSize As Int32
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Public dwMask As Int32
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Public wNumbering As Int16
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Public wReserved As Int16
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2111:PointersShouldNotBeVisible")> _
    Public dxStartIndent As IntPtr
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2111:PointersShouldNotBeVisible")> _
    Public dxRightIndent As IntPtr
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2111:PointersShouldNotBeVisible")> _
    Public dxOffset As IntPtr
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Public wAlignment As Int16
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Public cTabCount As Short
    <MarshalAs(UnmanagedType.ByValArray, SizeConst:=32)> _
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2111:PointersShouldNotBeVisible")> _
    Public rgxTabs() As IntPtr
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2111:PointersShouldNotBeVisible")> _
    Public dySpaceBefore As IntPtr
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2111:PointersShouldNotBeVisible")> _
    Public dySpaceAfter As IntPtr
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Security", "CA2111:PointersShouldNotBeVisible")> _
    Public dyLineSpacing As IntPtr
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Public sStyle As Short
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Public bLineSpacingRule As Byte
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Public bOutlineLevel As Byte
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Public wShadingWeight As Int16
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Public wShadingStyle As Int16
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Public wNumberingStart As Int16
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Public wNumberingStyle As Int16
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Public wNumberingTab As Int16
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Public wBorderSpace As Int16
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Public wBorderWidth As Int16
    <System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1051:DoNotDeclareVisibleInstanceFields")> _
    Public wBorders As Int16
End Structure

''' <summary>Contains native method wrappers</summary>
NotInheritable Class NativeMethods

    ''' <summary>Hiding the default constructor</summary>
    Partial Private Sub New()
    End Sub

    ''' <summary><see cref="SendMessage"/> wrapper for using with <see cref="PARAFORMAT2"/> structure</summary>
    ''' <param name="hWnd"> handle to the window whose window procedure will receive the message</param>
    ''' <param name="msg">The message to be sent.</param>
    ''' <param name="wParam">Additional message-specific information.</param>
    ''' <param name="lParam">Additional message-specific information - the <see cref="PARAFORMAT2"/> structure in thdis case.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("user32", CharSet:=CharSet.Auto)> _
   Public Shared Function SendMessage( _
           ByVal hWnd As HandleRef, _
           ByVal msg As Int32, _
           ByVal wParam As IntPtr, _
           ByRef lParam As PARAFORMAT2) As IntPtr
    End Function

#Region "Win32 API Constants"
    Public Const GetParaFormatMessage As Int32 = 1085
    Public Const SetParaFormatMessage As Int32 = 1095
    Public Const NumberingValid As Int32 = &H20
    Public Const NumberingStyleValid As Int32 = &H2000&
    Public Const NumberingBullet As Int32 = 2
    Public Const NumberingStart As Int32 = &H8000&
#End Region

End Class
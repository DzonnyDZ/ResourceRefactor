﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.1
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System

Namespace My.Resources
    
    'This class was auto-generated by the StronglyTypedResourceBuilder
    'class via a tool like ResGen or Visual Studio.
    'To add or remove a member, edit your .ResX file then rerun ResGen
    'with the /str option, or rebuild your VS project.
    '''<summary>
    '''  A strongly-typed resource class, for looking up localized strings, etc.
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0"),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute()>  _
    Friend Class Strings
        
        Private Shared resourceMan As Global.System.Resources.ResourceManager
        
        Private Shared resourceCulture As Global.System.Globalization.CultureInfo
        
        <Global.System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")>  _
        Friend Sub New()
            MyBase.New
        End Sub
        
        '''<summary>
        '''  Returns the cached ResourceManager instance used by this class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend Shared ReadOnly Property ResourceManager() As Global.System.Resources.ResourceManager
            Get
                If Object.ReferenceEquals(resourceMan, Nothing) Then
                    Dim temp As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager("Microsoft.VSPowerToys.ResourceRefactor.Strings", GetType(Strings).Assembly)
                    resourceMan = temp
                End If
                Return resourceMan
            End Get
        End Property
        
        '''<summary>
        '''  Overrides the current thread's CurrentUICulture property for all
        '''  resource lookups using this strongly typed resource class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend Shared Property Culture() As Global.System.Globalization.CultureInfo
            Get
                Return resourceCulture
            End Get
            Set
                resourceCulture = value
            End Set
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to &lt;Add new resource file...&gt;.
        '''</summary>
        Friend Shared ReadOnly Property CreateNewResourceText() As String
            Get
                Return ResourceManager.GetString("CreateNewResourceText", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Could not extract string to resource.
        '''</summary>
        Friend Shared ReadOnly Property ErrorCaption() As String
            Get
                Return ResourceManager.GetString("ErrorCaption", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Refactoring aborted because file {0} could not be checked out for editing..
        '''</summary>
        Friend Shared ReadOnly Property FileCheckoutError() As String
            Get
                Return ResourceManager.GetString("FileCheckoutError", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Refactoring aborted because file {0} is read only and can not be edited..
        '''</summary>
        Friend Shared ReadOnly Property FileReadOnlyError() As String
            Get
                Return ResourceManager.GetString("FileReadOnlyError", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to This is not a valid resource name..
        '''</summary>
        Friend Shared ReadOnly Property InvalidResourceNameError() As String
            Get
                Return ResourceManager.GetString("InvalidResourceNameError", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to You have not selected any resource entry to replace the string..
        '''</summary>
        Friend Shared ReadOnly Property NoResourceSelected() As String
            Get
                Return ResourceManager.GetString("NoResourceSelected", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to  .
        '''</summary>
        Friend Shared ReadOnly Property PreviewChangeSeperator() As String
            Get
                Return ResourceManager.GetString("PreviewChangeSeperator", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to Resource name already exists in the selected resource file, please select another name..
        '''</summary>
        Friend Shared ReadOnly Property ResourceAlreadyExists() As String
            Get
                Return ResourceManager.GetString("ResourceAlreadyExists", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to ^[A-Za-z_][A-Za-z_0-9]*$.
        '''</summary>
        Friend Shared ReadOnly Property ResourceNameValidationRegExp() As String
            Get
                Return ResourceManager.GetString("ResourceNameValidationRegExp", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to http://go.microsoft.com/fwlink/?LinkId=77546.
        '''</summary>
        Friend Shared ReadOnly Property SendFeedbackURL() As String
            Get
                Return ResourceManager.GetString("SendFeedbackURL", resourceCulture)
            End Get
        End Property
    End Class
End Namespace

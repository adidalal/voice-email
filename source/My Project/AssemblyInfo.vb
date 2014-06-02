Imports System.Resources

Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security

' General Information about an assembly is controlled through the following 
' set of attributes. Change these attribute values to modify the information
' associated with an assembly.

' Review the values of the assembly attributes

<Assembly: AssemblyTitle("Voice Email")> 
<Assembly: AssemblyDescription("An Outlook extension to reply to emails with recorded messages. Please make sure lame.exe and BasicAudio.dll are in C:/OutlookAddin if you experience any crashes.")> 
<Assembly: AssemblyCompany("Rutgers University, NJ")> 
<Assembly: AssemblyProduct("Voice Email")> 
<Assembly: AssemblyCopyright("Aditya Dalal - http://adityadalal.com")> 
<Assembly: AssemblyTrademark("")> 

' Setting ComVisible to false makes the types in this assembly not visible 
' to COM components.  If you need to access a type in this assembly from 
' COM, set the ComVisible attribute to true on that type.
<Assembly: ComVisible(False)> 

'The following GUID is for the ID of the typelib if this project is exposed to COM
<Assembly: Guid("40e35727-b0b6-414b-9042-1f30911e24a1")> 

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

<Assembly: AssemblyVersion("3.0.0.0")> 
<Assembly: AssemblyFileVersion("3.0.0.0")> 

Friend Module DesignTimeConstants
    Public Const RibbonTypeSerializer As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Serialization.RibbonTypeCodeDomSerializer, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Public Const RibbonBaseTypeSerializer As String = "System.ComponentModel.Design.Serialization.TypeCodeDomSerializer, System.Design"
    Public Const RibbonDesigner As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Design.RibbonDesigner, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
End Module
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAnalysiser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Dim lastStage As Integer







Public Function Analysis(pos As Integer)
    Dim data As WeldData
    data = buf(pos)
    
    
    
    
End Function








'
'



'
'

'UPSET
' Enable analysis
' Checking this option will include the Upset in the analysis; Otherwise it is excluded from the analysis.
'
'Minimum (mm)
' The minimum/maximum settings in millimetres (mm), wherein the result of the Upset must fall . It must be greater or equal to the minimum setting, and less or equal to the maximum setting for the criterion to pass. Otherwise, the criterion is deemed as a fail.
'
'Maximum (mm)
'
'FORGE
' Enable analysis
' Checking this option will include the Forge in the analysis; Otherwise it is excluded from the analysis.
'
'Minimum (t)
' The minimum/maximum settings in metric tonnes (tonnes), wherein the result of the Forge force must fall . It must be greater or equal to the minimum setting, and less or equal to the maximum setting for the criterion to pass. Otherwise, the criterion is deemed as fail.
'
'Maximum (t)
'
'Slippage
' Enable analysis
' Checking this option will include the Upset in the analysis; Otherwise it is excluded from the analysis.
'
'UPSET Time(s)
' The Upset setting in millimetres (mm) and the Upset Time in seconds (s), together define the slope of the Upset. If this slope is similar in comparison to the result of the weld’s Upset slope, it is deemed as a fail; Otherwise it is deemed as a pass.
'
'UPSET (mm)
'
'Current Interrupt in Boost
' Enable analysis
' Checking this option will include the Current Interrupt in the analysis; Otherwise it is excluded from the analysis.
'
'Current (A)
' The Current setting in Amperes (A) and the Time setting in seconds (s), are used together to analyse the Current interrupt in the Boost stage. The Current setting defines the interrupt level. Any Current reading during the Boost stage that is less than or equal to the setting defined, is deemed as a Current interrupt. If the Current interrupt goes for more than the Time setting defined, it is deemed as a fail; Otherwise it is deemed as a pass.
'
'Time (s)
'
'Short-Circuit in Boost
' Enable analysis
' Checking this option will include the Short-Circuit in the analysis; Otherwise it is excluded from the analysis.
'
'Current (A)
' The Current setting in Amperes (A) and the Time setting in seconds (s), are used together to analyse the Short-Circuit in the Boost stage. The Current setting defines the short circuit level. Any Current reading during the Boost stage that is greater than or equal to the setting defined is deemed as a Short-Circuit. If the Short-Circuit goes for more than the Time setting defined, it is deemed as a fail; Otherwise it is deemed as a pass.
'
'Time (s)
'
'Total Rail Usage
' Enable analysis
' Checking this option will include the Total Rail Usage in the analysis; Otherwise it is excluded from the analysis.
'
'Total Rail(mm)
' The Total Rail setting in millimetres (mm) where the result of the total rail usage must exceed. If the result is greater than the Total Rail setting, it is deemed as a pass; Otherwise it is deemed as a fail.
'



Attribute VB_Name = "modCommon"
Option Explicit
'This is example showing how to redefine base value of error map enum.
'Demo defines H_EXTBASE=1 in project settings, so VB picks up
'the following constant as base number for ENUM_ERRMAP.
'We just shift it one up compared to default value.
Public Const ERRMAP_BASE = vbObjectError + 4096 + 1



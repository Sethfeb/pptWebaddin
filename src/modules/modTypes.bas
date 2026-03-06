Attribute VB_Name = "modTypes"
'==============================================================================
' modTypes.bas
' 프로젝트 전역 공용 타입 정의 모듈
' 반드시 다른 모든 모듈보다 먼저 로드되어야 합니다.
'==============================================================================
Option Explicit

Public Type SpecRecord
    EquipID   As String
    ShortCode As String
    SpecName  As String
    SpecValue As String
    Unit      As String
    Revision  As Long
End Type

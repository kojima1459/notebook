Attribute VB_Name = "modTypes"
Option Explicit

' ============================================================================
' modTypes - shared Public Type definitions
' ----------------------------------------------------------------------------
' Both admin (writer) and chatbot (reader) use Chunk, so its declaration
' lives here to avoid the structure drifting out of sync.
' ============================================================================

Public Type Chunk
    Id As String
    Source As String
    Page As Long
    StartCharOffset As Long
    Text As String
End Type

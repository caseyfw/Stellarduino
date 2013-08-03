Attribute VB_Name = "ErrorConstants"
'---------------------------------------------------------------------
' Copyright © 2000-2001 SPACE.com Inc., New York, NY
'
' Permission is hereby granted to use this Software for any purpose
' including combining with commercial products, creating derivative
' works, and redistribution of source or binary code, without
' limitation or consideration. Any redistributed copies of this
' Software must include the above Copyright Notice.
'
' THIS SOFTWARE IS PROVIDED "AS IS". SPACE.COM, INC. MAKES NO
' WARRANTIES REGARDING THIS SOFTWARE, EXPRESS OR IMPLIED, AS TO ITS
' SUITABILITY OR FITNESS FOR A PARTICULAR PURPOSE.
'---------------------------------------------------------------------
'   ==================
'   ERRORCONSTANTS.BAS
'   ==================
'
' Declarations of error codes and error strings used in the ASCOM
' Telescope Interface implementation.
'
' Written:  27-Jun-00   Robert B. Denny <rdenny@dc3.com>
'
' Edits:
'
' When      Who     What
' --------- ---     --------------------------------------------------
' 27-Jun-00 rbd     Initial edit
' 13-Jan-00 rbd     Add ERR_SOURCE, necessary for proper errors, and
'                   property range error, used in template.
'---------------------------------------------------------------------

Option Explicit

Public Const ERR_SOURCE As String = "ASCOM EQASCOM Driver"

Public Const SCODE_NOT_IMPLEMENTED As Long = vbObjectError + &H400
Public Const MSG_NOT_IMPLEMENTED As String = _
    " is not implemented by this telescope driver object."

Public Const SCODE_PROP_RANGE_ERROR As Long = vbObjectError + &H401
Public Const MSG_PROP_RANGE_ERROR As String = _
    " is out of range in this telescope driver object."

Public Const SCODE_NOT_VALID As Long = vbObjectError + &H402
Public Const MSG_NOT_VALID As String = _
    " is invalid or returned an invalid response."

Public Const SCODE_PROP_NOT_SET As Long = vbObjectError + &H403
Public Const MSG_PROP_NOT_SET As String = _
    " has not been initialised."

Public Const SCODE_NOT_CONNECTED As Long = vbObjectError + &H404
Public Const MSG_NOT_CONNECTED As String = _
    "The scope is not connected."

Public Const SCODE_VAL_OUTOFRANGE As Long = vbObjectError + &H405
Public Const MSG_VAL_OUTOFRANGE As String = _
    "The property value is out of range"

Public Const SCODE_SETUP_CONNECTED = vbObjectError + &H406
Public Const MSG_SETUP_CONNECTED = _
    "You cannot change the driver's configuration while it is connected to a telescope."

Public Const SCODE_WRONG_TRACKING As Long = vbObjectError + &H407
Public Const MSG_WRONG_TRACKING As String = _
    "Wrong tracking state"

Public Const SCODE_SCOPE_PARKED As Long = vbObjectError + &H410
Public Const MSG_SCOPE_PARKED As String = _
    " function is not possible while scope is parked."

Public Const SCODE_ALTAZ_SLEW_ERROR As Long = vbObjectError + &H411
Public Const MSG_ALTAZ_SLEW_ERROR As String = _
    "AltAz slew is not allowed while scope is Tracking."

Public Const SCODE_RADEC_SLEW_ERROR As Long = vbObjectError + &H412
Public Const MSG_RADEC_SLEW_ERROR As String = _
    "RaDec slew is not allowed while scope is not Tracking."



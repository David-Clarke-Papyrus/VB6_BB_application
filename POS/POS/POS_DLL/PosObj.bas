Attribute VB_Name = "PosObj"
Option Explicit
Global sInBox As String
Global sOutBox As String
Global sServerInbox As String
Global sServerOutbox As String
Global sPOSSQLServer As String
Public oZSession As z_ZSession
Public Const POLL_INTERVAL As Long = 3000
Public oPC As z_POSCLIConnection

Global lngExchangeNumber As Long




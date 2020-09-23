VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7358
   ClientLeft      =   65
   ClientTop       =   403
   ClientWidth     =   17550
   LinkTopic       =   "Form1"
   ScaleHeight     =   7358
   ScaleWidth      =   17550
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOutput 
      Height          =   949
      Left            =   3360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   240
      Width           =   12375
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   960
      Top             =   3000
      _ExtentX        =   1006
      _ExtentY        =   1006
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1335
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   15975
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()
Dim objDatabase As clsAPI
Set objDatabase = New clsAPI

    txtOutput.Text = vbNullString

Dim strUrl As String
strUrl = "webapi.selcomm.com"

Dim JSON As New ChilkatJsonObject
success = JSON.UpdateString("periodKey", "A001")
success = JSON.UpdateNumber("vatDueSales", "105.50")
success = JSON.UpdateNumber("vatDueAcquisitions", "-100.45")
success = JSON.UpdateNumber("totalVatDue", "5.05")
success = JSON.UpdateNumber("vatReclaimedCurrPeriod", "105.15")
success = JSON.UpdateNumber("netVatDue", "100.10")
success = JSON.UpdateInt("totalValueSalesExVAT", 300)
success = JSON.UpdateInt("totalValuePurchasesExVAT", 300)
success = JSON.UpdateInt("totalValueGoodsSuppliedExVAT", 3000)
success = JSON.UpdateInt("totalAcquisitionsExVAT", 3000)
success = JSON.UpdateBool("finalised", 1)

'''{
  "ChargeCode": "FG",
  "Price": "12.50",
  "MarkUp": "150",
  "ChargeDescription": "Special price for Gordon",
  "From": "2020-01-01 00:00:00",
  "To": "9999-01-01 00:00:00",
  "PlanId": 1234,
  "PlanOptionId": 1
}'''
txtOutput.Text = objDatabase.Authenticate(strUrl, """iotBilling""")

txtOutput.Text = objDatabase.Post(strUrl, "", "", JSON.Emit())

End Sub


VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmRlogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "rlogin"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRemoteUser 
      Height          =   285
      Left            =   6360
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton btnDisconnect 
      Caption         =   "Disconnect"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txtRemoteHost 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton btnConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5520
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   9720
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   513
      LocalPort       =   1002
   End
   Begin VB.TextBox txtTerm 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   600
      Width           =   9975
   End
   Begin VB.Label lblRemoteUser 
      Caption         =   "Remote User:"
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label lblHost 
      Caption         =   "Remote Host:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   150
      Width           =   1095
   End
End
Attribute VB_Name = "frmRlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' rlogin example
'
'   This is a very simplified example of how to use the rlogin protocol to connect
'   to remote hosts.  To the end user, an rlogin client has the look and feel of a
'   telnet session... but it's not... kinda...  In actuality rlogin is a watered
'   down subset of telnet, and is much easier to deal with for the programmer.
'
'   Most of the negotiation phase of telnet is not present in rlogin, and the myriad
'   of session options is also out the window.  For this reason, rlogin is a good
'   starting point for beginning socket programmers.
'
'   Now... just because rlogin is "easy", doesn't mean it's not powerful.  With an
'   rlogin client you basically have the same functionality as if you where sitting
'   in front of the console at the server... all you need now is root access and
'   you can be really dangerous.  In fact, this example is the framework for a project
'   that will allow remote-control of our servers from a web page.  This code will be
'   turned into a COM Component DLL which will be registered on our web server and
'   it's methods and properties will be called from an ASP page.  Now that's powerful!
'
'   To get your rlogin client to work, there is some configuration on the server side
'   that must take place.  For starters, you will need an account that will be used
'   to login to.  In the home directory of that account you will need to create a file
'   named .rhosts which will need at least one line that designates who is allowed to
'   rlogin.  The format for this line is "host" "user" (without the quotes), where
'   "host" is the hostname or IP address of the machine where the client is running,
'   and "user" is the account under which the client is running.  For security purposes,
'   it is better to use the IP address rather than the hostname because it is harder to
'   spoof an IP address (most networks don't allow duplicate IP addresses, yet may allow
'   duplicate hostnames).  You can also just use "+ +" (without the quotes) to specify
'   that any user from any host can rlogin... but, you guessed it... this would send
'   your local administrator into a fit of rage!
'
'   This code was developed and tested in the following environment;
'       - Micro$oft Windows 2000 Professional, 5.00.2195, SP3
'       - Micro$oft Visual Studio 6.0 Enterprise Edition
'       - redhat Linux 8.0 (rlogin server)
'       - Sun Solaris 5.7 (rlogin server)
'       - NCR MP-RAS SVR4 (rlogin server)
'       - OpenVMS Alpha 7.2-1 (rlogin server)
'
'   This example should work with little to no modification on most Unix, VMS, and
'   Windows platforms which offer rlogin services.  The rlogin server socket should
'   be 513 on standard server installations, and the client socket must be in the
'   "well-known" or "privileged" range.  This example uses one of the sockets from
'   the unreserved well-known ports range of 1001-1009.
'
'   For more information on the rlogin protocol see RFC1282.
'
'
'   Enjoy!
'
'   R.G. Brown
'   Senior Software Engineer
'   Intel Corporation
'
'
'   HISTORY
'     v1.0.0 - 11 May 2003
'       » Original version
'
'     v1.0.1 - 12 May 2003
'       » Implemented keyboard buffering... lack of buffering was causing issues on
'         Solaris and MP-RAS
'       » Testing on OpenVMS revealed need to set terminal characteristics _after_
'         connection... specifically SET TERM/PASTHRU
'       » Added Winsock.Close in Winsock_Error subroutine
'

Option Explicit                                 'good programming practice

'Windows API function to get the user name of the currently logged on user
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                         (ByVal lpBuffer As String, nSize As Long) As Long

'Connect button click
Private Sub btnConnect_Click()
    If Not Winsock.State = sckConnected Then    'if socket NOT connected...
        Winsock.RemoteHost = txtRemoteHost.Text 'set socket RemoteHost
        Winsock.Connect                         'request socket connect
    End If
    txtTerm.SetFocus                            'set focus to terminal window
End Sub

'Disconnect button click
Private Sub btnDisconnect_Click()
    If Not Winsock.State = sckClosed Then       'if socket NOT closed...
        Winsock.Close                           'request socket close
        btnConnect.Enabled = True               'enable Connect buttton
        btnDisconnect.Enabled = False           'disable Disconnect button
    End If
    txtTerm.Text = ""                           'clear terminal
    txtTerm.SetFocus                            'set focus to terminal window
End Sub
'Application start-up
Private Sub Form_Load()
    Dim strUserName As String                   'variable to hold user name
    strUserName = Space(512)                    'prep user name variable
    GetUserName strUserName, Len(strUserName)   'Windows API function call
    strUserName = Trim(strUserName)             'trim spaces (Chr(0) terminated)
    Winsock.Tag = strUserName                   'stuff user name into Winsock tag :o)
End Sub

'Application closing
Private Sub Form_Unload(Cancel As Integer)
    If Not Winsock.State = sckClosed Then       'if socket NOT closed...
        Winsock.Close                           'request socket close
    End If
End Sub

'Key pressed in terminal window
Private Sub txtTerm_KeyPress(KeyAscii As Integer)
    If Winsock.State = sckConnected Then        'if socket IS connected...
        Select Case KeyAscii                    'determine key
            Case vbKeyBack
                '                               'ignore backspace
            Case vbKeyReturn
                Winsock.SendData txtTerm.Tag & vbCr  'send tag and CR to socket
                txtTerm.Tag = ""                     'clear tag
            Case Else
                txtTerm.Tag = txtTerm.Tag & Chr(KeyAscii) 'append key to tag
        End Select
    End If
End Sub

'Ignore arrow keys pressed in terminal window
Private Sub txtTerm_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft, vbKeyUp, vbKeyRight, vbKeyDown
            KeyCode = 0
    End Select
End Sub

'Event raised when server has closed socket
Private Sub Winsock_Close()
    btnDisconnect_Click                         'imitate Disconnect button click
End Sub

'Event raised when socket has connected
Private Sub Winsock_Connect()
    '*** Excerpt from RFC1282 ***
    'Upon connection establishment, the client sends four null-terminated
    'strings to the server.  The first is an empty string (i.e., it
    'consists solely of a single zero byte), followed by three non-null
    'strings: the client username, the server username, and the terminal
    'type and speed.  More explicitly:
    '
    '     <null>
    '     client-user-name<null>
    '     server-user-name<null>
    '     terminal-type/speed<null>
    '
    '     For example:
    '
    '     <null>
    '     bostic<null>
    '     kbostic<null>
    '     vt100/9600<null>
    '
    Winsock.SendData Chr(0) & _
                     Winsock.Tag & _
                     txtRemoteUser.Text & Chr(0) & _
                     "dumb/128000" & Chr(0)
    btnConnect.Enabled = False                  'disable Connect button
    btnDisconnect.Enabled = True                'enable Disconnect button
End Sub

'Event raised when data arrives at socket
Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)

    Dim strData                                 'variable to receive socket data
    Winsock.GetData strData, vbString           'request socket data
    
    '*** Excerpt from RFC1282 ***
    'The server returns a zero byte to indicate that it has received these
    'strings and is now in data transfer mode.  Window size negotiation
    'may follow this initial exchange (see below).
    If strData = Chr(0) Then                    'server says c'mon in
        Exit Sub                                'no need to process further
    End If
    
    If strData = Chr(1) Then                    'server says GO AWAY
        btnDisconnect_Click                     'imitate Disconnect button click
        Exit Sub                                'no need to process further
    End If
    
    'Window size negotiation (see RFC1282)
    'server shouldn't ask for negotiation for a "dumb" terminal, as in this example,
    'but most likely will ask for other terminal types (ansi, vtNNN, etc.)
    'format is FFFFssrrccxpyp (FF FF s s rows chars/row xpixels ypixels)
    If strData = Chr(128) Then
        Winsock.SendData Chr(255) & Chr(255) & "ss24800714"
        Exit Sub                                'no need to process further
    End If
    
    'Add socket data to terminal window.
    '
    'There is an inherent problem here with the TextBox control... it has a max size
    'of approximately 32K.  If it reaches this maximum, this example will seem to quit
    'responding, even though in the background everything is functioning as normal.
    'The problem here is only that the TextBox has nowhere to put the extra data.
    'To handle this properly, there should probably be a check of the terminal text
    'length, and whack some number of characters from the head when the tail reaches
    'somewhere in the 75% full range.
    txtTerm.Text = txtTerm.Text & strData
    
    'The following two lines of code move the cursor to the end of the terminal window
    'to provide the visual effect of data being appended to the end of the terminal
    'window, while previous data scrolls off the top of the screen.
    '
    'Unfortunately... this is a very bad example of how to do this.  The main reason
    'this is a bad example is performance... the more data in the TextBox, the more
    'the user sees the screen flicker.  This is because when the TextBox control gets
    'data appended to it, by default it returns the cursor to row 1, column 1 of the
    'TextBox, then we come along and tell it to go to the last row and column.
    '
    'The best way to handle this for a commercial grade application would be to use
    'the Windows API function SendMessage, but to keep this example simple, we'll just
    'stick to the "poor-man's" way.
    '
    'For those who are interested in learning more about how to use the SendMessage API
    'there is a good Micro$oft Knowledge Base Article with examples at:
    'http://support.microsoft.com/support/kb/articles/q161/2/70.asp
    txtTerm.SelStart = Len(txtTerm.Text)
    txtTerm.SelLength = 0

End Sub

'Houston... we have a problem
'This event occurs whenever an error occurs in background processing in a Winsock control,
'such as a failure to connect or a failure to send or receive in the background.
Private Sub Winsock_Error(ByVal Number As Integer, Description As String, _
                          ByVal Scode As Long, ByVal Source As String, _
                          ByVal HelpFile As String, ByVal HelpContext As Long, _
                          CancelDisplay As Boolean)
    If Not Winsock.State = sckClosed Then       'if socket NOT closed...
        Winsock.Close                           'request socket close
    End If
    MsgBox "Winsock Error: " & Number & vbCrLf & Description
End Sub

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "OAuth2 Examples"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "GitHub"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SalesForce"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "LinkedIn"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CheckBox CheckAbort 
      Caption         =   "Abort by Checking this Box"
      Height          =   255
      Left            =   8760
      TabIndex        =   4
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox textBox1 
      Height          =   5535
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1680
      Width           =   11175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "GeoOp"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Google"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Facebook"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const facebookAuthEndpoint As String = "https://www.facebook.com/dialog/oauth"
Const facebookTokenEndpoint As String = "https://graph.facebook.com/oauth/access_token"
Const googleAuthEndpoint As String = "https://accounts.google.com/o/oauth2/v2/auth"
Const googleTokenEndpoint As String = "https://www.googleapis.com/oauth2/v4/token"
Const linkedinAuthEndpoint As String = "https://www.linkedin.com/oauth/v2/authorization"
Const linkedinTokenEndpoint As String = "https://www.linkedin.com/oauth/v2/accessToken"
Const salesForceAuthEndpoint As String = "https://login.salesforce.com/services/oauth2/authorize"
Const salesForceTokenEndpoint As String = "https://login.salesforce.com/services/oauth2/token"
Const gitAuthEndpoint As String = "https://github.com/login/oauth/authorize"
Const gitTokenEndpoint As String = "https://github.com/login/oauth/access_token"
Const geoopAuthEndpoint As String = "https://login.geoop.com/oauth2/code"
Const geoopTokenEndpoint As String = "https://login.geoop.com/oauth2/token"

' Replace these with actual values.
Const facebookClientId As String = "FACEBOOK-CLIENT-ID"
Const facebookClientSecret As String = "FACEBOOK-CLIENT-SECRET"

Const googleClientId As String = "GOOGLE-CLIENT-ID"
Const googleClientSecret As String = "GOOGLE-CLIENT-SECRET"

Const linkedinClientId As String = "LINKEDIN-CLIENT-ID"
Const linkedinClientSecret As String = "LINKEDIN-CLIENT-SECRET"

Const salesForceClientId As String = "SALESFORCE-CLIENT-ID"
Const salesForceClientSecret As String = "SALESFORCE-CLIENT-SECRET"

Const gitClientId As String = "GITHUB-CLIENT-ID"
Const gitClientSecret As String = "GITHUB-CLIENT-SECRET"

Const geoopClientId As String = "GEOOP-CLIENT-ID"

Dim TokenFilename As String
Dim ListenPort As Integer
Dim AuthorizationEndpoint As String
Dim TokenEndpoint As String
Dim ClientId As String
Dim ClientSecret As String
Dim LocalHost As String
Dim CodeChallenge As Integer
Dim CodeChallengeMethod As String
Dim Scope As String

Private Sub doOAuth2()

    Dim oauth2 As New ChilkatOAuth2
    Dim success As Long
    
    '  (If GeoOp) This should match the Site URL configured for your GeoOp Application, such as "http://localhost:3017/"
    oauth2.ListenPort = ListenPort
    
    oauth2.AuthorizationEndpoint = AuthorizationEndpoint
    oauth2.TokenEndpoint = TokenEndpoint
    
    '  Replace the client ID with an actual value.
    oauth2.ClientId = ClientId
    '  (If GeoOp) The ClientSecret should remain empty for a GeoOp public application
    oauth2.ClientSecret = ClientSecret

    oauth2.Scope = Scope
    oauth2.CodeChallenge = CodeChallenge
    oauth2.CodeChallengeMethod = CodeChallengeMethod
    
    oauth2.LocalHost = LocalHost
    
    '  Begin the OAuth2 three-legged flow.  This returns a URL that should be loaded in a browser.
    Dim url As String
    url = oauth2.StartAuth()
    If (oauth2.LastMethodSuccess <> 1) Then
        textBox1.Text = oauth2.LastErrorText
        Exit Sub
    End If

'    If (Len(url) > 0) Then
'        textBox1.Text = url
'        oauth2.Cancel
'        Exit Sub
'    End If
    
    ' At this point, your application should load the URL in a browser.
    ' Warning: I've found that passing an invalid URL can crash the program, and also
    ' can crash the VB6 IDE if debugging.  (for example, if a url does not begin with "https://")
    CreateObject("Wscript.Shell").Run url

    '  Now wait for the authorization.
    '  We'll wait for a max of 90 seconds.
    Dim numMsWaited As Long
    numMsWaited = 0
    Do While (numMsWaited < 90000) And (oauth2.AuthFlowState < 3)
        oauth2.SleepMs 50
        numMsWaited = numMsWaited + 50
        ' Keep the UI responsive..
        DoEvents
        If (CheckAbort.Value = 1) Then
            textBox1.Text = "Aborted."
            success = oauth2.Cancel()
            Exit Sub
        End If
    Loop
    
    '  If there was no response from the browser within 90 seconds, then
    '  the AuthFlowState will be equal to 1 or 2.
    '  1: Waiting for Redirect. The OAuth2 background thread is waiting to receive the redirect HTTP request from the browser.
    '  2: Waiting for Final Response. The OAuth2 background thread is waiting for the final access token response.
    '  In that case, cancel the background task started in the call to StartAuth.
    If (oauth2.AuthFlowState < 3) Then
        success = oauth2.Cancel()
        textBox1.Text = "No response from the browser!"
        Exit Sub
    End If
    
    '  Check the AuthFlowState to see if authorization was granted, denied, or if some error occurred
    '  The possible AuthFlowState values are:
    '  3: Completed with Success. The OAuth2 flow has completed, the background thread exited, and the successful JSON response is available in AccessTokenResponse property.
    '  4: Completed with Access Denied. The OAuth2 flow has completed, the background thread exited, and the error JSON is available in AccessTokenResponse property.
    '  5: Failed Prior to Completion. The OAuth2 flow failed to complete, the background thread exited, and the error information is available in the FailureInfo property.
    If (oauth2.AuthFlowState = 5) Then
        textBox1.Text = "OAuth2 failed to complete." & vbCrLf & oauth2.FailureInfo
        Exit Sub
    End If
    
    If (oauth2.AuthFlowState = 4) Then
        textBox1.Text = "OAuth2 authorization was denied." & vbCrLf & oauth2.AccessTokenResponse
        Exit Sub
    End If
    
    If (oauth2.AuthFlowState <> 3) Then
        textBox1.Text = "Unexpected AuthFlowState:" & oauth2.AuthFlowState
        Exit Sub
    End If
    
    textBox1.Text = "OAuth2 authorization granted!" & vbCrLf & "Access Token = " & oauth2.AccessToken
    
    '  Save the entire JSON response, which includes the access token, for future calls.
    '  The JSON AccessTokenResponse looks something like this:
    '  {"access_token":"e6dqdG....mzjpT04w==","token_type":"Bearer","expires_in":2592000,"owner_id":984236}
    
    Dim fac As New CkFileAccess
    success = fac.WriteEntireTextFile(App.Path & "\qa_data\tokens\" & TokenFilename, oauth2.AccessTokenResponse, "utf-8", 0)
    
    textBox1.Text = textBox1.Text & vbCrLf & "Success."
End Sub

' OAuth2 for Facebook
Private Sub Command1_Click()

    TokenFilename = "facebook.json"
    ListenPort = 3017
    AuthorizationEndpoint = facebookAuthEndpoint
    TokenEndpoint = facebookTokenEndpoint
    ClientId = facebookClientId
    ClientSecret = facebookClientSecret
    LocalHost = "localhost"
    CodeChallenge = 1
    CodeChallengeMethod = "S256"
    ' Set the Scope to a comma-separated list of permissions the app wishes to request.
    ' See https://developers.facebook.com/docs/facebook-login/permissions/ for a full list of permissions.
    Scope = "public_profile,user_friends,email,user_posts,user_likes,user_photos"
    doOAuth2

End Sub

' Google Apps (using Google Drive as an example..)
Private Sub Command2_Click()

    TokenFilename = "google.json"
    ListenPort = 3017
    AuthorizationEndpoint = googleAuthEndpoint
    TokenEndpoint = googleTokenEndpoint
    ClientId = googleClientId
    ClientSecret = googleClientSecret
    LocalHost = "localhost"
    CodeChallenge = 1
    CodeChallengeMethod = "S256"
    ' This is the scope for Google Drive.
    ' See https://developers.google.com/identity/protocols/googlescopes
    Scope = "https://www.googleapis.com/auth/drive"
    doOAuth2

End Sub

' OAuth2 for GeoOp
Private Sub Command3_Click()

    TokenFilename = "geoop.json"
    ListenPort = 3017
    AuthorizationEndpoint = geoopAuthEndpoint
    TokenEndpoint = geoopTokenEndpoint
    ClientId = geoopClientId
    '  The ClientSecret should remain empty for a GeoOp public application
    ClientSecret = ""
    '  Setting LocalHost equal to "none" prevents the "redirect_uri" query param from being sent in the initial HTTP request.
    '  Note: The GeoOp Application should still have a redirect URL specified as "http://localhost:3017/", where the port
    '  number matches the ListenPort above.
    LocalHost = "none"
    CodeChallenge = 0
    CodeChallengeMethod = ""
    Scope = "default"
    doOAuth2

End Sub

' LinkedIn
Private Sub Command4_Click()
    TokenFilename = "linkedin.json"
    ListenPort = 3017
    AuthorizationEndpoint = linkedinAuthEndpoint
    TokenEndpoint = linkedinTokenEndpoint
    ClientId = linkedinClientId
    ClientSecret = linkedinClientSecret
    LocalHost = "localhost"
    CodeChallenge = 1
    CodeChallengeMethod = "S256"
    Scope = ""
    doOAuth2
End Sub

' SalesForce
Private Sub Command5_Click()
    TokenFilename = "salesforce.json"
    ListenPort = 3017
    AuthorizationEndpoint = salesForceAuthEndpoint
    TokenEndpoint = salesForceTokenEndpoint
    ClientId = salesForceClientId
    ClientSecret = salesForceClientSecret
    LocalHost = "localhost"
    CodeChallenge = 1
    CodeChallengeMethod = "S256"
    Scope = ""
    doOAuth2
End Sub

' GitHub
Private Sub Command6_Click()
    TokenFilename = "github.json"
    ListenPort = 3017
    AuthorizationEndpoint = gitAuthEndpoint
    TokenEndpoint = gitTokenEndpoint
    ClientId = gitClientId
    ClientSecret = gitClientSecret
    LocalHost = "localhost"
    CodeChallenge = 1
    CodeChallengeMethod = "S256"
    Scope = ""
    doOAuth2
End Sub

Private Sub Form_Load()
    Dim glob As New ChilkatGlobal
    If glob.UnlockBundle("Anything for 30-day trial") <> 1 Then
        MsgBox "Failed to unlock Chilkat."
        textBox1.Text = glob.LastErrorText
    End If

    Dim fac As New CkFileAccess
    fac.DirAutoCreate App.Path & "\qa_data\tokens\"
End Sub


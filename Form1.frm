VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   508
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEmbedClr3Bit 
      Caption         =   "Embed Clr 3 Bit"
      Height          =   255
      Left            =   4080
      TabIndex        =   11
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdExtractClr3Bit 
      Caption         =   "Extract Clr 3 Bit"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdEmbedClr4Bit 
      Caption         =   "Embed Clr 4 Bit"
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdExtractClr4Bit 
      Caption         =   "Extract Clr 4 Bit"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.PictureBox picProgress 
      Height          =   135
      Left            =   1920
      ScaleHeight     =   5
      ScaleMode       =   0  'User
      ScaleWidth      =   109
      TabIndex        =   7
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton cmdExtractGray3Bit 
      Caption         =   "Extract BW 3 Bit"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   1335
   End
   Begin VB.PictureBox picParent 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   2
      Top             =   960
      Width           =   975
      Begin VB.Image Image1 
         Height          =   255
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdEmbedGray3Bit 
      Caption         =   "Embed BW 3 Bit"
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.PictureBox picChild 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4080
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9240
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "<<<< Embed In Parent Extract To Child >>>>"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Child Picture"
      Height          =   255
      Left            =   4920
      TabIndex        =   5
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Parent Picture"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOpenParent 
         Caption         =   "Open Parent Picture"
      End
      Begin VB.Menu mnuFileOpenChild 
         Caption         =   "Open Child Picture"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save Parent Picture As"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'**********************************************************
'**********************************************************
'***                                                    ***
'***                         *                          ***
'***                         *                          ***
'***                    * * * * * *                     ***
'***                         *                          ***
'***                         *                          ***
'***                         *                          ***
'***                         *                          ***
'***                         *                          ***
'***                                                    ***
'**********************************************************
'************For God so loved the world that***************
'*************He gave His only begotten Son****************
'************that whosoever believeth in Him***************
'******************should not perish***********************
'****************but have eternal life*********************
'*********************(John 3:16)**************************
'**********************************************************
'Summary
'This application will embed/hide a picture within  another picture
'and extract it at any time.
'There is little or no distinction between the original and
'one that has been emdedded with another.
'There are three options for embedding. 3 Bit Greyscale, 3 Bit and 4 Bit Color.
'On the 3 Bit Greyscale the child picture is converted to greyscale.
'On the 3 Bit and 4 Bit Color the child picture retains its color.
'The trick to this program is utilizing the lower order 3 and 4 bits
'of the RGB in each pixel in the parent picture.
'These bits are cleared to zero then loaded with our child information
'as follows;

'On the 4 Bit color option, the child's higher order 4 Bits for each pixel's
'RGB is placed in the cleared lower order of the parent's.

'On the 3 Bit color option, the child's higher order 3 Bits for each pixel's
'RGB is placed in the cleared lower order of the parent's.

'On the 3 Bit greyscale option, the child's RGB for each pixel
'is converted to a greyscale byte. This byte is further divided into the
'higher order 3 bits
'middle order 3 bits
'lower  order 2 bits
'and placed in the lower order 3 bits of the parent's RGB in each pixel.

'As for extraction the process is reversed.

Option Explicit
Dim lngColorParent As Long
Dim lngColorChild As Long
Dim intParentR As Integer, intParentG As Integer, intParentB As Integer
Dim intChildR As Integer, intChildG As Integer, intChildB As Integer
Dim intMaxPicWidth As Integer
Dim intMaxPicHeight As Integer

Private Sub Form_Load()
'Set all dimensions
Form1.ScaleMode = 3
Form1.Height = 5900
Form1.Width = 7500
picParent.ScaleMode = 3
picParent.Left = 8
picParent.Picture = picParent.Image
picParent.AutoRedraw = True
picChild.AutoRedraw = True
picProgress.AutoRedraw = False
picProgress.ScaleWidth = 100
picProgress.ScaleHeight = 10
 'These two variables may changed to encrypt larger pictures
 'Of course, our form size might to to be changed accordingly too
 intMaxPicWidth = 200
 intMaxPicHeight = 250
End Sub

Private Sub mnuFileOpenParent_Click()
                           'Common Dialog Window
  '***************************************************************************
  CommonDialog1.CancelError = True           'Enable on error or cancel GoTo
    On Error GoTo cancelPressed
  CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
  CommonDialog1.DialogTitle = "Open Picture" 'Title displayed
  'CommonDialog1.InitDir = App.Path           'Start Directory
  CommonDialog1.Filter = "Pictures (*.bmp;*.gif;*.jpg)|*.bmp;*.gif;*.jpg| All (*.*)|*.*"
  CommonDialog1.FileName = ""
  CommonDialog1.ShowOpen
  '***************************************************************************

'Paint parent picture and resize accordingly
  Image1.Picture = LoadPicture(CommonDialog1.FileName)
picParent.AutoRedraw = True
picParent.Width = Image1.Width
picParent.Height = Image1.Height
picParent.Picture = Image1.Picture
If picParent.Width > intMaxPicWidth Then picParent.Width = intMaxPicWidth
If picParent.Height > intMaxPicHeight Then picParent.Height = intMaxPicHeight
'Resize Child picture box to the same size too
picChild.AutoRedraw = True
picChild.Width = picParent.Width
picChild.Height = picParent.Height
picChild.Picture = picChild.Image
cancelPressed:
End Sub

Private Sub mnuFileOpenChild_Click()
                           'Common Dialog Window
  '***************************************************************************
  CommonDialog1.CancelError = True           'Enable on error or cancel GoTo
    On Error GoTo cancelPressed
  CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
  CommonDialog1.DialogTitle = "Open Picture" 'Title displayed
  'CommonDialog1.InitDir = App.Path           'Start Directory
  CommonDialog1.Filter = "Pictures (*.bmp;*.gif;*.jpg )|*.bmp;*.gif;*.jpg| All (*.*)|*.*"
  CommonDialog1.FileName = ""
  CommonDialog1.ShowOpen
  '***************************************************************************

'Paint child picture. No resizing
picChild.AutoRedraw = True
picChild.Picture = LoadPicture(CommonDialog1.FileName)
cancelPressed:
End Sub

Private Sub mnuFileSaveAs_Click()
                         'Common Dialog Window
'***************************************************************************
CommonDialog1.CancelError = True
  On Error GoTo cancelPressed
CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt
CommonDialog1.DialogTitle = "Save Parent Picture"
'CommonDialog1.InitDir = App.Path
CommonDialog1.Filter = "Pictures (Bitmap (*.bmp)|*.bmp"
CommonDialog1.FileName = "Picture"
CommonDialog1.ShowSave
SavePicture picParent.Picture, CommonDialog1.FileName
'***************************************************************************
cancelPressed:
End Sub

Private Sub cmdEmbedGray3Bit_Click()
Dim x As Integer, y As Integer
Dim intBWrgb As Integer
Dim intBWr As Integer
Dim intBWg As Integer
Dim intBWb As Integer
'Here we will embed the Child picture to the Parent
'after we convert it to greyscale.

'Scan each pixel
For y = 0 To picParent.Height - 1
For x = 0 To picParent.Width - 1

'Get parent pixel color
lngColorParent = picParent.Point(x, y)

'Extract RGB
rtnGetParentColors

'                [p=parent c=child]
'Clear Parent's lower 3 bits for embedding picture
       '(pppppppp  and 11111000) = ppppp000
intParentR = (intParentR And 248) 'ppppp000
intParentG = (intParentG And 248) 'ppppp000
intParentB = (intParentB And 248) 'ppppp000

'Get Child's pixel color
lngColorChild = picChild.Point(x, y)

'Extract RGB
rtnGetChildColors

'Convert to greyscale
intBWrgb = (intChildR + intChildG + intChildB) / 3

'Extract the
'higher order 3 bits
intBWr = (intBWrgb And 224) 'ccc00000
'middle order 3 bits
intBWg = (intBWrgb And 28)  '000ccc00
'lower order 2 bits
intBWb = (intBWrgb And 3)   '000000cc

'Logical shift right to get bits in lower order
 intBWr = Int(intBWr / 32) '00000ccc
 intBWg = Int(intBWg / 4)  '00000ccc
'intBWb = (intBWrgb /  1)  '000000cc

'Combine and paint           ppppp000  +  00000ccc = pppppccc
picParent.PSet (x, y), RGB((intParentR Or intBWr), (intParentG Or intBWg), (intParentB Or intBWb))
Next x

'Progress bar
picProgress.Line (0, 0)-(y * 100 / picParent.Height, 10), vbBlue, BF
Next y

'Finished so clear Progress bar
picProgress.Picture = LoadPicture("")
picParent.Picture = picParent.Image
picChild.Picture = LoadPicture("")
End Sub

Private Sub cmdExtractGray3Bit_Click()
Dim x As Integer, y As Integer
Dim intNewPxl As Integer
'Here we will extract the Child greyscale picture from the Parent

'Scan each pixel
For y = 0 To picParent.Height - 1
For x = 0 To picParent.Width - 1

'Get parent pixel color
lngColorParent = picParent.Point(x, y)

'Extract RGB
rtnGetParentColors

'                                    [p=parent c=child]
'Extract the lower 3 Bits of R and  G and B and combine to our original greyscale pixel
'  =  (pppppccc and 00000111) * LSL5   (pppppccc and 00000111) LSL2   (pppppccc and 00000011)
'  =               (00000ccc) * LSL5              (00000ccc) * LSL2         (000000cc)
'  =               (ccc00000)                     (000ccc00)                (000000cc)
'  =               (cccccccc) original greyscale pixel. LSL = Logical Bit Shift Left
intNewPxl = ((intParentR And 7) * 32) + ((intParentG And 7) * 4) + (intParentB And 7)

'Paint it to Child
picChild.PSet (x, y), RGB(intNewPxl, intNewPxl, intNewPxl)
Next x

'Progress bar
picProgress.Line (0, 0)-(y * 100 / picParent.Height, 10), vbBlue, BF
Next y

'Finished so clear Progress bar
picProgress.Picture = LoadPicture("")
End Sub

Private Sub cmdEmbedClr4Bit_Click()
Dim x As Integer, y As Integer
'Here we will embed the Child picture to the Parent
' in 4 Bit Color

'Scan each pixel
For y = 0 To picParent.Height - 1
For x = 0 To picParent.Width - 1

'Get parent pixel color
lngColorParent = picParent.Point(x, y)

'Extract RGB
rtnGetParentColors

'                [p=parent c=child]
'Clear parent's lower 4 bits for embedding picture
'       (pppppppp  and 11110000) = pppp0000
intParentR = (intParentR And 240) 'pppp0000
intParentG = (intParentG And 240) 'pppp0000
intParentB = (intParentB And 240) 'pppp0000

'Get Child pixel color
lngColorChild = picChild.Point(x, y)

'Extract RGB
rtnGetChildColors

'Logical shift right to get higher order 4 bits in lower order 4 bits
intChildR = Int(intChildR / 16) '76543210 = 00007654
intChildG = Int(intChildG / 16) '76543210 = 00007654
intChildB = Int(intChildB / 16) '76543210 = 00007654

'Combine and paint           pppp0000  +  0000cccc = ppppcccc
picParent.PSet (x, y), RGB((intParentR Or intChildR), (intParentG Or intChildG), (intParentB Or intChildB))
Next x

'Progress bar
picProgress.Line (0, 0)-(y * 100 / picParent.Height, 10), vbBlue, BF
Next y

'Finished so clear Progress bar
picProgress.Picture = LoadPicture("")
picParent.Picture = picParent.Image
picChild.Picture = LoadPicture("")

End Sub

Private Sub cmdExtractClr4Bit_Click()
Dim x As Integer, y As Integer
'Here we will extract the Child 4 Bit color picture from the Parent

'Scan each pixel
For y = 0 To picParent.Height - 1
For x = 0 To picParent.Width - 1
lngColorParent = picParent.Point(x, y)

'Extract RGB
rtnGetParentColors

'                                           [p=parent c=child]
'Combine and paint      ppppcccc  +  00001111 * LSL4                     ,same,same
'                     =              0000cccc * LSL4                     ,same,same
'                     =              cccc0000                            ,same,same
picChild.PSet (x, y), RGB((intParentR And 15) * 16, (intParentG And 15) * 16, (intParentB And 15) * 16)
Next x

'Progress bar
picProgress.Line (0, 0)-(y * 100 / picParent.Height, 10), vbBlue, BF
Next y

'Finished so clear Progress bar
picProgress.Picture = LoadPicture("")

End Sub

Private Sub cmdEmbedClr3Bit_Click()
Dim x As Integer, y As Integer
'Here we will embed the Child picture to the Parent
'in 3 Bit Color

'Scan each pixel
For y = 0 To picParent.Height - 1
For x = 0 To picParent.Width - 1

'Get parent pixel color
lngColorParent = picParent.Point(x, y)

'Extract RGB
rtnGetParentColors

'                [p=parent c=child]
'Clear parent's lower 3 bits for embedding picture
'       (pppppppp  and 11111000) = ppppp000
intParentR = (intParentR And 240) 'ppppp000
intParentG = (intParentG And 240) 'ppppp000
intParentB = (intParentB And 240) 'ppppp000

'Get Child pixel color
lngColorChild = picChild.Point(x, y)

'Extract RGB
rtnGetChildColors

'Logical shift right to get higher order 3 bits in lower order 3 bits
intChildR = Int(intChildR / 32) '76543210 = 00000765
intChildG = Int(intChildG / 32) '76543210 = 00000765
intChildB = Int(intChildB / 32) '76543210 = 00000765

'Combine and paint           ppppp000  +  00000ccc = pppppccc
picParent.PSet (x, y), RGB((intParentR Or intChildR), (intParentG Or intChildG), (intParentB Or intChildB))
Next x

'Progress bar
picProgress.Line (0, 0)-(y * 100 / picParent.Height, 10), vbBlue, BF
Next y

'Finished so clear Progress bar
picProgress.Picture = LoadPicture("")
picParent.Picture = picParent.Image
picChild.Picture = LoadPicture("")

End Sub

Private Sub cmdExtractClr3Bit_Click()
Dim x As Integer, y As Integer
'Here we will extract the Child 3 Bit color picture from the Parent

'Scan each pixel
For y = 0 To picParent.Height - 1
For x = 0 To picParent.Width - 1
lngColorParent = picParent.Point(x, y)

'Extract RGB
rtnGetParentColors

'                          [p=parent c=child]
'Combine and paint      pppppccc  +  00000111 * LSL5                     ,same,same
'                     =              00000ccc * LSL5                     ,same,same
'                     =              ccc00000                            ,same,same
picChild.PSet (x, y), RGB((intParentR And 7) * 32, (intParentG And 7) * 32, (intParentB And 7) * 32)
Next x

'Progress bar
picProgress.Line (0, 0)-(y * 100 / picParent.Height, 10), vbBlue, BF
Next y

'Finished so clear Progress bar
picProgress.Picture = LoadPicture("")

End Sub

Private Sub rtnGetParentColors()
'Extract RGB from pixel
intParentR = lngColorParent Mod 256
intParentB = Int(lngColorParent / 65536)
intParentG = (lngColorParent - (intParentB * 65536) - intParentR) / 256
End Sub
Private Sub rtnGetChildColors()
'Extract RGB from pixel
intChildR = lngColorChild Mod 256
intChildB = Int(lngColorChild / 65536)
intChildG = (lngColorChild - (intChildB * 65536) - intChildR) / 256
End Sub

Private Sub mnuFileExit_Click()
Unload Form1
End
End Sub


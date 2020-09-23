Attribute VB_Name = "Module1"
Option Explicit

'---- Used for Storing Palette Information ----
Public Type typRGBColor
    Red As Long
    Blue As Long
    Green As Long
End Type

Public Pal(255) As typRGBColor

'---- Used for Storing Frame Information ----
Public Type typFrame
    Position As Long
    Size As Long
    Unknown As Byte
    TypeNumber As Integer
    FrameNumber As Integer
    fUnknown As String * 4
    CompressionType As Integer
    XLength As Integer
    YLength As Integer
    XOffset As Integer
    YOffset As Integer
    LinePosition() As Long
End Type

'---- Used for Storing Image Type Information ----
Public Type typImageType
    Position As Long
    Size As Long
    Unknown As String * 4
    NumberofFrames As Integer
    CurrentFrame As Long
    Frame() As typFrame
End Type

Public ImageType() As typImageType
Public NumberOfTypes As Long
Public CurrentType As Long

'Just some variables for file access
Private Filenum As Long
Private FileLoaded As Boolean

Public Function LoadPal(filepath As String)
    Dim i As Long, bytTemp As Byte
    'Find a free file number
    Filenum = FreeFile
    'Open the file
    Open filepath For Binary As #Filenum
        'The first 4 bytes are not needed
        Seek #Filenum, 5
        'Get the 256 colour palette
        For i = 0 To 255
            'Each byte has a value between 0 and 63, which is equal to 0 to 255
            Get #Filenum, , bytTemp: Pal(i).Red = (bytTemp / 63) * 255
            Get #Filenum, , bytTemp: Pal(i).Green = (bytTemp / 63) * 255
            Get #Filenum, , bytTemp: Pal(i).Blue = (bytTemp / 63) * 255
        Next i
    'Close it when done
    Close #Filenum
End Function

Public Function LongPal(index As Long) As Long
    'Return the long value of the RGB colour
    LongPal = RGB(Pal(index).Red, Pal(index).Green, Pal(index).Blue)
End Function

Public Function OpenFile(Path As String) As Long
Dim temp As Integer, i As Long
'Take care of errors
On Error GoTo Errhand:
'Find a free number
Filenum = FreeFile

'If there is already an open file, close it
If FileLoaded = True Then CloseFile
'Open the file
Open Path For Binary As #Filenum
FileLoaded = True
'The first byte is the number of types of image
Get #Filenum, 85, temp: NumberOfTypes = temp
'redim the array
ReDim ImageType(NumberOfTypes - 1)

'The next data we need is located at the 129th byte
Seek #Filenum, 129
'For all the types
For i = 0 To NumberOfTypes - 1
 'The first long is the position less one
 Get #Filenum, , ImageType(i).Position: ImageType(i).Position = ImageType(i).Position + 1
 'The second is the size of the image type
 Get #Filenum, , ImageType(i).Size
Next i

'Return the number of types less one
OpenFile = NumberOfTypes - 1

Exit Function
Errhand:
MsgBox "Could not open file"
End Function

Public Function ReadFrames(CurrentType As Long) As Long
Dim i As Long
On Error GoTo Errhand:

'Go to the Position of the current type
Seek #Filenum, ImageType(CurrentType).Position
'The next 4 bytes are unknown data, but store them anyway
Get #Filenum, , ImageType(CurrentType).Unknown
'The next 2 bytes are the number of frames
Get #Filenum, , ImageType(CurrentType).NumberofFrames

'redim the array for the number of frames
ReDim ImageType(CurrentType).Frame(ImageType(CurrentType).NumberofFrames)

Dim tmp1 As Byte, tmp2 As Byte, tmp3 As Byte
'For each of the frames...
For i = 0 To ImageType(CurrentType).NumberofFrames - 1
 'Get the first 3 bytes...
 Get #Filenum, , tmp1: Get #Filenum, , tmp2: Get #Filenum, , tmp3
 'And using this nasty formula calculate the position in the file of the image data
 ImageType(CurrentType).Frame(i).Position = 65536 * tmp3 + 256& * tmp2 + tmp1 + ImageType(CurrentType).Position
 'The next byte is again, unknown
 Get #Filenum, , ImageType(CurrentType).Frame(i).Unknown
 'The next 2 bytes...
 Get #Filenum, , tmp1: Get #Filenum, , tmp2
 'Can be used to calculate the size of the frame
 ImageType(CurrentType).Frame(i).Size = 256& * tmp2 + tmp1
Next i

'Return the number of frames less one
ReadFrames = ImageType(CurrentType).NumberofFrames - 1

Exit Function
Errhand:
MsgBox "Could not read Frames"
End Function

Public Function DrawImage(CurrentType As Long, CurrentFrame As Long)
Dim i As Long, tmp1 As Byte, tmp2 As Byte
On Error GoTo Errhand:
'Clear the image, otherwise we can get trails
frmMain.Picture1.Cls

'Just for reference so we know what frame we're looking at
ImageType(CurrentType).CurrentFrame = CurrentFrame
'If the size is 0 (or less), the image is either corrupt, or has no image data
If ImageType(CurrentType).Frame(CurrentFrame).Size < 1 Then Exit Function

'Seek to the position of the frame image data
Seek #Filenum, ImageType(CurrentType).Frame(CurrentFrame).Position
'First 2 bytes is the type number
Get #Filenum, , ImageType(CurrentType).Frame(CurrentFrame).TypeNumber
'Next 2 bytes is the frame number
Get #Filenum, , ImageType(CurrentType).Frame(CurrentFrame).FrameNumber
'Next 2 is unknown data
Get #Filenum, , ImageType(CurrentType).Frame(CurrentFrame).fUnknown
'Next 2 is the Compression Type
Get #Filenum, , ImageType(CurrentType).Frame(CurrentFrame).CompressionType
'Next 2 is the length of the x axis
Get #Filenum, , ImageType(CurrentType).Frame(CurrentFrame).XLength
'Next 2 is the length of the y axis
Get #Filenum, , ImageType(CurrentType).Frame(CurrentFrame).YLength
'Next 2 is the X offset
Get #Filenum, , ImageType(CurrentType).Frame(CurrentFrame).XOffset
'Next 2 is the Y offset
Get #Filenum, , ImageType(CurrentType).Frame(CurrentFrame).YOffset

Dim ylen As Long
'ylen is just so I don't have to type out the other statement all the time
ylen = ImageType(CurrentType).Frame(CurrentFrame).YLength
'Resize the array for imagedata
ReDim ImageType(CurrentType).Frame(CurrentFrame).LinePosition(ylen)

'Set the line position of each frame to the correct position
For i = 0 To ImageType(CurrentType).Frame(CurrentFrame).YLength - 1
 'Given by the current position...
 ImageType(CurrentType).Frame(CurrentFrame).LinePosition(i) = Seek(Filenum)
 'in addition to these two bytes
 Get #Filenum, , tmp1: Get #Filenum, , tmp2
 'using this formula
 ImageType(CurrentType).Frame(CurrentFrame).LinePosition(i) = ImageType(CurrentType).Frame(CurrentFrame).LinePosition(i) + (tmp2 * 256& + tmp1)
Next i

Dim StXPos As Long, StYPos As Long, xpos As Long, ypos As Long, datlen As Long
'set where to start drawing
StXPos = 150 - ImageType(CurrentType).Frame(CurrentFrame).XOffset: StYPos = ImageType(CurrentType).Frame(CurrentFrame).YOffset - 50
'set the current position inthe right place
xpos = ImageType(CurrentType).Frame(CurrentFrame).XLength: ypos = -1
Do
 
 'until xpos is less than the total length
 Do Until xpos < ImageType(CurrentType).Frame(CurrentFrame).XLength
  'increment the yposition
  ypos = ypos + 1
  'if ypos is the same as the ylength then we've reached the end of the line, and exit
  If ypos = ImageType(CurrentType).Frame(CurrentFrame).YLength Then Exit Do
  'otherwise, find where the data starts
  Seek #Filenum, ImageType(CurrentType).Frame(CurrentFrame).LinePosition(ypos)
  'get the value
  Get #Filenum, , tmp1
  'and assign it to xpos
  xpos = tmp1
 Loop
 'if we are at the end of the line, exit
 If ypos = ylen Then Exit Do
 'get the next byte
 Get #Filenum, , tmp1
 'and assign it to datlen
 datlen = tmp1
 'if the frame is compressed
 If ImageType(CurrentType).Frame(CurrentFrame).CompressionType = 1 Then
  'if datlen is odd
  If (datlen And 1) = 1 Then
   datlen = datlen \ 2
   'get the next byte for the colour
   Get #Filenum, , tmp1
   'and draw a box filled with that colour, the size of datlen
   frmMain.Picture1.Line (xpos + StXPos, ypos + StYPos)-(xpos + datlen - 1 + StXPos, ypos + StYPos), LongPal(CLng(tmp1))
  Else
   'if datlen is even
   datlen = datlen \ 2
   'colour datlen number of pixels
   For i = xpos To xpos + datlen - 1
    'with the palette colour of tmp1
    Get #Filenum, , tmp1
    'and pset it on the picture
    frmMain.Picture1.PSet (i + StXPos, ypos + StYPos), LongPal(CLng(tmp1))
   Next i
  End If
 Else
  'otherwise, if it ain't compressed, draw the pixels
  For i = xpos To xpos + datlen - 1
    'of tmp1's colour
    Get #Filenum, , tmp1
   'on the screen
   frmMain.Picture1.PSet (i + StXPos, ypos + StYPos), LongPal(CLng(tmp1))
  Next i
 End If
 'increment xpos
 xpos = xpos + datlen
 'if xpos is less than xlength
 If xpos < ImageType(CurrentType).Frame(CurrentFrame).XLength Then
    'get the next byte
    Get #Filenum, , tmp1
    'and assign it to xpos
    xpos = xpos + tmp1
 End If
'and round we go
Loop
Exit Function

Errhand:
MsgBox "Could not draw image"
End Function

Public Function CloseFile()
'close the file
Close #Filenum
'and set the boolean to false
FileLoaded = False
End Function

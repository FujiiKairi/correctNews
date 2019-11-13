VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Sub ekisutoraNEWS()
    
'---�R�[�h1�b�C���^�[�l�b�g�ɐڑ����ău���E�U���J��---
    Dim objIE As InternetExplorer                               '���삷��IE������I�u�W�F�N�g������
    Set objIE = CreateObject("InternetExplorer.Application")    '��̃I�u�W�F�N�g��IE������
    objIE.Visible = False                                   '������悤��
    
'---�R�[�h2�b�C���^�[�l�b�g�̓���̃y�[�W���J��---
    objIE.navigate "https://www.sankei.com/politics/newslist/politics-n1.html#12"               '���̃y�[�W���J��
    Call IEWait(objIE)   'IE��ҋ@                              IEWait���Ă�
    Call WaitFor(3) '3�b��~                                    WaitFor�����

    Dim anchor As HTMLAnchorElement
    Dim numNews As Integer
    Dim i As Integer
    Dim pageIndex As Integer
'    i = 0
'    'url�����G�N�Z���ɏ�������
'    For Each elmListItem In objIE.document.getElementsByClassName("entry tertitary vertical")
'            Set anchor = elmListItem.getElementsByTagName("a")(0)
'            Worksheets("Sheet2").Cells(i + 2, 1).Value = anchor.href 'url���
'            i = i + 1
'    Next
   
    For pageIndex = 244 To 251
        objIE.navigate (Cells(pageIndex, 1).Value)
        'Sleep (10000)
        Do While objIE.Busy = True Or objIE.readyState < READYSTATE_COMPLETE
            DoEvents
        Loop
        Worksheets("Sheet2").Cells(pageIndex, 2).Value = objIE.document.getElementById("__r_article_title__").innerHTML '�L���^�C�g��
        Worksheets("Sheet2").Cells(pageIndex, 3).Value = objIE.document.getElementById("__r_publish_date__").innerHTML '�L���̓��t
        Sleep (1000)
        
        For Each elmListItem In objIE.document.getElementsByClassName("post_content")
            For Each objLink In elmListItem.getElementsByTagName("p")
                Worksheets("Sheet2").Cells(pageIndex, 4).Value = Worksheets("Sheet2").Cells(pageIndex, 4).Value + objLink.innerHTML
            Next
        Next
        
        'Sleep (10000)
    Next
        
        
'    For pageIndex = 101 To 150
'        objIE.navigate (Cells(pageIndex, 1).Value)
'        Sleep 5000
'        '�y�[�W�̃R�[�h�ԍ��̎擾
'        tmp = Split(Cells(pageIndex, 1).Value, "/")
'        Worksheets("imagedata").Cells(pageIndex, 3).Value = tmp(4)
'        '�t�H���_�쐬
'        FolderName = "C:\Users\mnuus\Downloads\�񎟃u���O\�摜\" & Worksheets("imagedata").Cells(pageIndex, 3).Value
'        If Dir(FolderName, vbDirectory) = "" Then   '�����̃t�H���_���Ȃ��ꍇ�t�H���_���쐬'
'            MkDir FolderName
'        End If
'        '��������^�O�̎擾
'        tagString = objIE.document.getElementsByClassName("tag-block")(0).getElementsByClassName("ga-event")(0).innerHTML
'        tmp = Split(tagString, "<")
'        tagString = tmp(0)
'        numTag = objIE.document.getElementsByClassName("tag-block")(0).getElementsByClassName("ga-event").Length
'        For i = 1 To numTag - 1
'            tmp = Split(objIE.document.getElementsByClassName("tag-block")(0).getElementsByClassName("ga-event")(i).innerHTML, "<")
'            tagString = tagString & "," & tmp(0)
'        Next
'        Worksheets("imagedata").Cells(pageIndex, 4).Value = tagString
'        '��������摜�ۑ�
'        For Each objLink In objIE.document.getElementsByClassName("gallery-block")
'            numPicture = objLink.getElementsByTagName("a").Length - 1
'        Next
'            'url��ۑ�����
'        For numPicIndex = 1 To numPicture
'            For Each objLink In objIE.document.getElementsByClassName("gallery-block")
'                Set anchor = objLink.getElementsByTagName("a")(numPicIndex - 1)
'                Worksheets("imagedata").Cells(pageIndex, 5 + numPicIndex - 1).Value = anchor.href 'url���
'
'            Next
'        Next
'        For numPicIndex = 0 To numPicture - 1
'              imgURL = Cells(pageIndex, 5 + numPicIndex).Value
'              '�摜�t�@�C����
'              Filename = Mid(imgURL, InStrRev(imgURL, "/") + 1)
'
'              '�摜�ۑ���(+�摜�t�@�C�����j
'              savePath = FolderName & "\" & numPicIndex & ".jpg"
'
'              '�L���b�V���N���A
'              cacheDel = DeleteUrlCacheEntry(imgURL)
'
'              '�摜�_�E�����[�h
'              result = URLDownloadToFile(0, imgURL, savePath, 0, 0)
'        Next
'
'    Next
'---�R�[�h6�bIE�����---
    MsgBox ("�I���܂���")
    objIE.Quit
    Set objIE = Nothing
ErrLabel:
    'msg = msg & "�G���[���������܂���"
    Resume Next
End Sub

Function OpenPage(ByVal url As String, ByRef objIE As Object)
    objIE.navigate (url)               '���̃y�[�W���J��
    Call WaitFor(3) '3�b��~                                    WaitFor�����
End Function

'---�R�[�h2-1�bIE��ҋ@����֐�---
Function IEWait(ByRef objIE As Object)                      '�I�u�W�F�N�g���Q�Ɠn��
    Do While objIE.Busy = True Or objIE.readyState <> 4     'busy�v���p�e�B��true��������readystate��4�iIE�I�u�W�F�N�g�̑S�f�[�^�ǂݍ��݊�����ԁj
        DoEvents
    Loop
End Function

'---�R�[�h2-2�b�w�肵���b������~����֐�---
Function WaitFor(ByVal second As Integer)
    Dim futureTime As Date
 
    futureTime = DateAdd("s", second, Now)                  'functime�����̎���+second(�����ł�3�b)�ɂ���
 
    While Now < futureTime
        DoEvents                                            '�L�����Z���{�^���Ȃǂ̃C�x���g���N���������ɂ��̏�����OS�ɂ킽��
    Wend                                                    'while do loop �Ɠ���
End Function













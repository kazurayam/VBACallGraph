VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "����Ǘ�"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cp As ComposedPersons

Public Sub LoadData(ByRef baseSheet As Worksheet)
    Set cp = New ComposedPersons
    Call cp.Initialize(baseSheet)
    Call cp.LoadData
    Call LoadIdList
    
    '����V�[�g��Range("B1")�� "$B$14" �̂悤�ȃA�h���X�����񂪏����Ă���B
    '���̃A�h���X�ɂ��ǂ�Person�I�u�W�F�N�g�ɑΉ����邩�𔻒肵�Ă�
    '����Person�̃f�[�^���t�H�[���ɓW�J����
    Dim aPerson As Person
    Set aPerson = cp.InterpreteSelectedRangeToPerson(baseSheet.Range("B1"))
    If Not aPerson Is Nothing Then
        LoadFields aPerson.ID
    End If
End Sub


'**
'* cp����ID�̃��X�g��ǂݍ���ŃR���{�{�b�N�XComboBoxId�ɃZ�b�g����
'*
Private Sub LoadIdList()
    ComboBoxId.List = cp.GetIdList()
    ComboBoxId.AddItem "New"
End Sub

'**
'* �C�x���g�v���V�[�W���F ComboBoxId_Change
'*
Private Sub ComboBoxId_Change()
    With ComboBoxId
        If IsValidId Then
            If Information.IsNumeric(.value) Then
                Call LoadFields(.value)
            Else
                Call ClearFields
            End If
        End If
    End With
End Sub

Private Property Get IsValidId() As Boolean
    IsValidId = False
    With ComboBoxId
        If (.value > 0 And .value <= cp.MaxId) Or _
            (.value = "New") Then
            IsValidId = True
        End If
    End With
End Property

'**
'* anId�Ŏw�肳�ꂽID��Person�I�u�W�F�N�g��Persons�R���N�V��������
'* ���o���āAUserForm�̊e�R���g���[���̒l�Ƃ��ăZ�b�g����
'*
Private Sub LoadFields(ByVal anID As Long)
    With cp.persons.Item(anID)
        ComboBoxId.value = anID
        TextBoxName.value = .Name
        Call SetGender(.Gender)
        TextBoxBirthday.value = .Birthday
        LabelAge.Caption = .Age
        CheckBoxActive.value = .Active
    End With
End Sub

'**
'* ���ʂ�\��������i�u�j�v�u���v�j�ɂ��ƂÂ���
'* �I�v�V�����{�^���̒l��ݒ肷��
'*
Private Sub SetGender(ByVal aGender As String)
    OptionButtonFemale.value = True
    If aGender = "�j" Then
        OptionButtonMale.value = True
    End If
End Sub

'**
'* UserForm�̊e�R���g���[���̒l������������
'*
Private Sub ClearFields()
    TextBoxName.value = ""
    OptionButtonMale.value = True
    TextBoxBirthday.value = ""
    LabelAge.Caption = ""
    CheckBoxActive.value = True
End Sub

'**
'* Event Procedure:
'* �t�H�[���́u�X�V�v�{�^�����������ꂽ�Ƃ��ɋN�������
'*
Private Sub CommandButtonUpdate_Click()
    '
    If CheckFields Then
        Dim p As Person: Set p = New Person
        If ComboBoxId.value = "New" Then
            p.ID = cp.MaxId + 1
        Else
            p.ID = CLng(ComboBoxId.value)
        End If
        p.Name = TextBoxName.text
        p.Birthday = TextBoxBirthday.value
        p.Gender = "��"
        If OptionButtonMale.value = True Then p.Gender = "�j"
        p.Active = CheckBoxActive.value
        
        '����Person�I�u�W�F�N�g�̏�Ԃ�ComposedPersons�ɋL�^����
        cp.UpdatePerson p
        ' �X�V���ꂽComposePersons�����[�N�V�[�g�̃e�[�u���ɔ��f����
        cp.ApplyData
        
        '�X�V���ꂽComposedPersons�ɓ�������悤�Ƀt�H�[���̕\�����X�V����
        Call LoadFields(p.ID)
        Call LoadIdList
        
    End If
End Sub

'**
'* �t�H�[����ɂ���e�R���g���[���̒l�����������͂���Ă��邩�ǂ�����
'* ���肷��
'*
Private Function CheckFields() As Boolean
    
    CheckFields = True
    
    If Not IsValidId Then
        MsgBox "�uID�v�Ƃ��Ă�1�ȏ��" & cp.MaxId & "�ȉ��̐��l�܂���""New""����͂��Ă�������", vbInformation
        CheckFields = False
    End If
    
    If Len(TextBoxName.text) = 0 Then
        MsgBox "�u���O�v����͂��Ă�������", vbInformation
        CheckFields = False
    End If
    
    If Not IsDate(TextBoxBirthday.value) Then
        MsgBox "�u�a�����v�ɓ��t����͂��Ă�������", vbInformation
        CheckFields = False
    End If
    
End Function


'**
'* Event Procedure:
'* �t�H�[���́u����v�{�^���������ꂽ�Ƃ��ɋN�������
'*
Private Sub CommandButtonClose_Click()
    Unload Me
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "msg_buddies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public buddy_name As String
Public buddy_ip As String
Public buddy_port As String
Public buddy_file As String
Public buddy_enabled As Integer

Public Function clear_all()
    buddy_name = ""
    buddy_ip = ""
    buddy_port = ""
    buddy_file = ""
    buddy_enabled = 0
    buddy_filter = 0
End Function

Public Function save(index As Variant)
    If buddy_name <> "" Then
        SaveSetting App.ProductName & " Buddies", Str(index), "name", buddy_name
        SaveSetting App.ProductName & " Buddies", Str(index), "ip", buddy_ip
        SaveSetting App.ProductName & " Buddies", Str(index), "port", buddy_port
        SaveSetting App.ProductName & " Buddies", Str(index), "file", buddy_file
        SaveSetting App.ProductName & " Buddies", Str(index), "enabled", buddy_enabled
    End If
End Function

Public Function load(index As Variant)
    If GetSetting(App.ProductName & " Buddies", Str(index), "name") <> "" Then
        buddy_name = GetSetting(App.ProductName & " Buddies", Str(index), "name")
        buddy_ip = GetSetting(App.ProductName & " Buddies", Str(index), "ip")
        buddy_port = GetSetting(App.ProductName & " Buddies", Str(index), "port")
        buddy_file = GetSetting(App.ProductName & " Buddies", Str(index), "file")
        buddy_enabled = Val(GetSetting(App.ProductName & " Buddies", Str(index), "enabled"))
    End If
End Function

Public Function delete(index As Variant)
    DeleteSetting App.ProductName & " Buddies", Str(index), "name"
End Function

Public Function start_buddy()
    frm_messenger.msa_buddy.Characters.load buddy_name, buddy_file
    frm_messenger.msa_buddy.Characters(buddy_name).Show
End Function

Public Function stop_buddy()
    'frm_messenger.msa_buddy.Characters.Unload
End Function

Attribute VB_Name = "Globals"
Global cliserv As String
Global win_top As Long
Global win_left As Long
Global remote_port As String
Global local_port As String
Global remote_ip As String
Global my_name As String
Global input_x As Long
Global input_y As Long
Global Const iam_server = "Server"
Global Const iam_client = "Client"
'********************************
'ADDINS DIRECTORIES
Global Const icons_path = "\Icons\"
Global Const sounds_path = "\Sounds\"
'SOUND EVENTS
Global Const noof_events = 9
Global snd_events(noof_events) As sound_event
'EVENT LIST
Global Const event_onload = 0
Global Const event_onunload = 1
Global Const event_onrx = 2
Global Const event_onsend = 3
Global Const event_onerror = 4
Global Const event_onconnect = 5
Global Const event_ondisconnect = 6
Global Const event_ontype = 7
Global Const event_onreturn = 8
'MESSAGE HTML FILE PATH
Global Const messages_path = "\Messages\"
Global Const messages_file = "tmpmsgs.htm"
Global Const start_message = 1
Global Const end_message = 2
'ICON DEFS
Global Const noof_icons = 24
Global msg_icons(noof_icons) As icon_defs
'ICONS LIST
Global Const icon_smile = 0
Global Const icon_sad = 1
Global Const icon_beer = 2
Global Const icon_disgust = 3
Global Const icon_love = 4
Global Const icon_smileo = 5
Global Const icon_smilep = 6
Global Const icon_wink = 7
Global Const icon_unlove = 8
Global Const icon_crooked = 9
Global Const icon_coctail = 10
Global Const icon_gift = 11
Global Const icon_smiled = 12
Global Const icon_email = 13
Global Const icon_man = 14
Global Const icon_woman = 15
Global Const icon_vampire = 16
Global Const icon_kiss = 17
Global Const icon_rose = 18
Global Const icon_star = 19
Global Const icon_sleep = 20
Global Const icon_hot = 21
Global Const icon_photo = 22
Global Const icon_cat = 23
'COLOUR VARIABLES
Global Const noof_colvars = 3
Global wm_colvars(noof_colvars) As colour_variable
'COLOUR VARIABLE LIST
Global Const colvar_rxmsg = 0
Global Const colvar_txmsg = 1
Global Const colvar_page = 2
'HELP STRING TABLE
Global Const help_colours = "This is a list of all changeable colours with Web Messenger.  You can give Web Messenger a more personal touch by choosing your favourite colours, or ones that are easier on your eyes."
Global Const help_icons = "This is a list of the available icons which you can use in your messages, simply type the required string stated below and it shall be replaced by a small picture.  For example, type the message "" I love (B) and hate (C) "".  The (B) should be replaced by a pint of beer and the (C) should be replaced by a Coctail glass."
Global Const help_sndevents = "This is a list of all sound events in Web Messenger.  You may change the sounds to the events and even test them by pressing the test button.  You can also enable or disable the sound by the check box along side."
Global Const help_main = "Welcome to Web Messenger.  Before you can begin your chat session you must decide who is going to be the server and who is going to be the client.  Once decided, complete the appropriate box on the right and then you may begin, simply press the begin button to do so."
Global Const help_buddies = "This is where you configure all of your buddies.  By adding a buddy it is much quicker to start chatting as you wont have to put the details in again!  Also you can use an MSAgent character for each buddy, this will come onto the screen when you connect and then speak to you as well as the written message.  You can download new characters from the internet, just search from MSAgent characters!"
'MESSAGE BUDDIES
Global Const max_buddies = 50
Public buddies(max_buddies) As msg_buddies
'********************************

Public Sub init_buddies()
    For init_buddy = 0 To max_buddies
        Set buddies(init_buddy) = New msg_buddies
        With buddies(init_buddy)
            .clear_all
        End With
    Next init_buddy
End Sub

Public Sub setup_colvars()
    For init_colvar = 0 To noof_colvars
        Set wm_colvars(init_colvar) = New colour_variable
        With wm_colvars(init_colvar)
            Select Case init_colvar
                Case Is = colvar_rxmsg
                    .variable_name = "colvar_rxmsg"
                    .variable_description = "Recieved messages"
                Case Is = colvar_txmsg
                    .variable_name = "colvar_txmsg"
                    .variable_description = "Sent messages"
                Case Is = colvar_page
                    .variable_name = "colvar_page"
                    .variable_description = "Message page colour"
            End Select
        End With
    Next init_colvar
End Sub

Public Sub setup_icons()
    For init_icons = 0 To noof_icons
        Set msg_icons(init_icons) = New icon_defs
        With msg_icons(init_icons)
            Select Case init_icons
                Case Is = icon_smile
                    .icon_filename = App.Path & icons_path & "smile.gif"
                    .icon_recogstr = ":)"
                    .icon_description = "Smiley face"
                Case Is = icon_sad
                    .icon_filename = App.Path & icons_path & "sad.gif"
                    .icon_recogstr = ":("
                    .icon_description = "Sad face"
                Case Is = icon_beer
                    .icon_filename = App.Path & icons_path & "beer.gif"
                    .icon_recogstr = "(B)"
                    .icon_description = "Pint of beer"
                Case Is = icon_disgust
                    .icon_filename = App.Path & icons_path & "disgust.gif"
                    .icon_recogstr = ":|"
                    .icon_description = "Disgust face"
                Case Is = icon_love
                    .icon_filename = App.Path & icons_path & "love.gif"
                    .icon_recogstr = "(L)"
                    .icon_description = "Heart"
                Case Is = icon_smileo
                    .icon_filename = App.Path & icons_path & "smileo.gif"
                    .icon_recogstr = ":o"
                    .icon_description = "Shocked Face"
                Case Is = icon_smilep
                    .icon_filename = App.Path & icons_path & "smilep.gif"
                    .icon_recogstr = ":p"
                    .icon_description = "Cheeky face"
                Case Is = icon_wink
                    .icon_filename = App.Path & icons_path & "wink.gif"
                    .icon_recogstr = ";)"
                    .icon_description = "Winking face"
                Case Is = icon_unlove
                    .icon_filename = App.Path & icons_path & "unlove.gif"
                    .icon_recogstr = "(U)"
                    .icon_description = "Broken heart"
                Case Is = icon_crooked
                    .icon_filename = App.Path & icons_path & "crooked.gif"
                    .icon_recogstr = ":/"
                    .icon_description = "Crooked face"
                Case Is = icon_coctail
                    .icon_filename = App.Path & icons_path & "coctail.gif"
                    .icon_recogstr = "(C)"
                    .icon_description = "Coctail glass"
                Case Is = icon_gift
                    .icon_filename = App.Path & icons_path & "gift.gif"
                    .icon_recogstr = "(G)"
                    .icon_description = "Gift"
                Case Is = icon_smiled
                    .icon_filename = App.Path & icons_path & "smiled.gif"
                    .icon_recogstr = ":D"
                    .icon_description = "Big smiley face"
                Case Is = icon_email
                    .icon_filename = App.Path & icons_path & "email.gif"
                    .icon_recogstr = "(E)"
                    .icon_description = "Email"
                Case Is = icon_man
                    .icon_filename = App.Path & icons_path & "man.gif"
                    .icon_recogstr = "(M)"
                    .icon_description = "Man"
                Case Is = icon_woman
                    .icon_filename = App.Path & icons_path & "woman.gif"
                    .icon_recogstr = "(W)"
                    .icon_description = "Woman"
                Case Is = icon_vampire
                    .icon_filename = App.Path & icons_path & "vampire.gif"
                    .icon_recogstr = "(V)"
                    .icon_description = "Vampire"
                Case Is = icon_kiss
                    .icon_filename = App.Path & icons_path & "kiss.gif"
                    .icon_recogstr = "(K)"
                    .icon_description = "Kissy lips"
                Case Is = icon_rose
                    .icon_filename = App.Path & icons_path & "rose.gif"
                    .icon_recogstr = "(R)"
                    .icon_description = "Rose"
                Case Is = icon_star
                    .icon_filename = App.Path & icons_path & "star.gif"
                    .icon_recogstr = "(S)"
                    .icon_description = "Star"
                Case Is = icon_sleep
                    .icon_filename = App.Path & icons_path & "zzz.gif"
                    .icon_recogstr = "(Z)"
                    .icon_description = "Sleeping moon"
                Case Is = icon_hot
                    .icon_filename = App.Path & icons_path & "hot.gif"
                    .icon_recogstr = "(H)"
                    .icon_description = "Sun with shades"
                Case Is = icon_photo
                    .icon_filename = App.Path & icons_path & "photo.gif"
                    .icon_recogstr = "(P)"
                    .icon_description = "Camera"
                Case Is = icon_cat
                    .icon_filename = App.Path & icons_path & "cat.gif"
                    .icon_recogstr = ":{"
                    .icon_description = "Cats face"
            End Select
        End With
    Next init_icons
End Sub

Public Sub setup_events()
    For init_events = 0 To noof_events
        Set snd_events(init_events) = New sound_event
        With snd_events(init_events)
            Select Case init_events
                Case Is = event_onload
                    .snd_name = "Web Messenger load"
                Case Is = event_onunload
                    .snd_name = "Web Messenger quit"
                Case Is = event_onrx
                    .snd_name = "Recieve message"
                Case Is = event_onsend
                    .snd_name = "Send Message"
                Case Is = event_onerror
                    .snd_name = "Error"
                Case Is = event_onconnect
                    .snd_name = "Connect"
                Case Is = event_ondisconnect
                    .snd_name = "Disconnect"
                Case Is = event_ontype
                    .snd_name = "Message type"
                Case Is = event_onreturn
                    .snd_name = "Message return"
            End Select
        End With
    Next init_events
End Sub

Public Sub save_window(window As String, save_top As Long, save_left As Long)
    SaveSetting App.ProductName, "windows", window, "SAVED"
    SaveSetting App.ProductName, "windows", window & " top", save_top
    SaveSetting App.ProductName, "windows", window & " left", save_left
End Sub

Public Sub load_window(window As String)
    'CHECK IF SETTINGS EXIST FIRST
    If GetSetting(App.ProductName, "windows", window) = "SAVED" Then
        'RETRIEVE SETTINGS
        win_top = Val(GetSetting(App.ProductName, "windows", window & " top"))
        win_left = Val(GetSetting(App.ProductName, "windows", window & " left"))
    Else
        win_top = 0
        win_left = 0
    End If
End Sub

Public Function increment_counter(counter As Integer, max As Integer) As Integer
    If counter < max Then
        counter = counter + 1
    Else
        counter = 0
    End If
    increment_counter = counter
End Function

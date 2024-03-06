Attribute VB_Name = "modInput"
Option Explicit

'//Mouse Button
Public Const VK_LBUTTON As Long = &H1
Public Const VK_RBUTTON As Long = &H2

'//System Key
Public Const VK_BACK As Long = &H8
Public Const VK_TAB As Long = &H9

'//Arrow Key
Public Const VK_LEFT As Long = &H25
Public Const VK_UP As Long = &H26
Public Const VK_RIGHT As Long = &H27
Public Const VK_DOWN As Long = &H28

Public UpKey As Boolean
Public DownKey As Boolean
Public LeftKey As Boolean
Public RightKey As Boolean
Public ShiftKey As Boolean
Public chkMoveKey As Boolean
Public UpMoveKey As Boolean
Public DownMoveKey As Boolean
Public LeftMoveKey As Boolean
Public RightMoveKey As Boolean
Public AtkKey As Boolean

'//In-Game Keys
Public Sub CheckKeys()
Dim i As Long
Dim Control As Long

    '//ControlKeys
    If GetKeyState(ControlKey(ControlEnum.KeyUp).cAsciiKey) >= 0 Then UpKey = False
    If GetKeyState(ControlKey(ControlEnum.KeyDown).cAsciiKey) >= 0 Then DownKey = False
    If GetKeyState(ControlKey(ControlEnum.KeyLeft).cAsciiKey) >= 0 Then LeftKey = False
    If GetKeyState(ControlKey(ControlEnum.KeyRight).cAsciiKey) >= 0 Then RightKey = False
    If GetKeyState(ControlKey(ControlEnum.KeyCheckMove).cAsciiKey) >= 0 Then chkMoveKey = False
    If GetKeyState(ControlKey(ControlEnum.KeyMoveUp).cAsciiKey) >= 0 Then UpMoveKey = False
    If GetKeyState(ControlKey(ControlEnum.KeyMoveDown).cAsciiKey) >= 0 Then DownMoveKey = False
    If GetKeyState(ControlKey(ControlEnum.KeyMoveLeft).cAsciiKey) >= 0 Then LeftMoveKey = False
    If GetKeyState(ControlKey(ControlEnum.KeyMoveRight).cAsciiKey) >= 0 Then RightMoveKey = False
    If GetKeyState(ControlKey(ControlEnum.KeyAttack).cAsciiKey) >= 0 Then AtkKey = False
        
    '//Constant
    If GetKeyState(vbKeyShift) >= 0 Then ShiftKey = False
    
    'If GetAsyncKeyState(VK_UP) >= 0 Then UpKey = False
    'If GetAsyncKeyState(VK_DOWN) >= 0 Then DownKey = False
    'If GetAsyncKeyState(VK_LEFT) >= 0 Then LeftKey = False
    'If GetAsyncKeyState(VK_RIGHT) >= 0 Then RightKey = False
End Sub

Public Sub CheckInputKeys()
    If GetKeyState(ControlKey(ControlEnum.KeyCheckMove).cAsciiKey) < 0 Then
        chkMoveKey = True
    Else
        chkMoveKey = False
    End If
    
    '//Move Key
    If GetKeyState(ControlKey(ControlEnum.KeyAttack).cAsciiKey) < 0 Then
        AtkKey = True
        UpMoveKey = True
        DownMoveKey = False
        RightMoveKey = False
        LeftMoveKey = False
    Else
        AtkKey = False
    End If
    If GetKeyState(ControlKey(ControlEnum.KeyMoveUp).cAsciiKey) < 0 Then
        UpMoveKey = True
        DownMoveKey = False
        RightMoveKey = False
        LeftMoveKey = False
    Else
        UpMoveKey = False
    End If
    If GetKeyState(ControlKey(ControlEnum.KeyMoveDown).cAsciiKey) < 0 Then
        UpMoveKey = False
        DownMoveKey = True
        RightMoveKey = False
        LeftMoveKey = False
    Else
        DownMoveKey = False
    End If
    If GetKeyState(ControlKey(ControlEnum.KeyMoveLeft).cAsciiKey) < 0 Then
        UpMoveKey = False
        DownMoveKey = False
        RightMoveKey = False
        LeftMoveKey = True
    Else
        LeftMoveKey = False
    End If
    If GetKeyState(ControlKey(ControlEnum.KeyMoveRight).cAsciiKey) < 0 Then
        UpMoveKey = False
        DownMoveKey = False
        RightMoveKey = True
        LeftMoveKey = False
    Else
        RightMoveKey = False
    End If
    
    'If GetKeyState(vbKeyUp) < 0 Then
    If GetKeyState(ControlKey(ControlEnum.KeyUp).cAsciiKey) < 0 Then
        UpKey = True
        DownKey = False
        LeftKey = False
        RightKey = False
        Exit Sub
    Else
        UpKey = False
    End If
    
    'If GetKeyState(vbKeyDown) < 0 Then
    If GetKeyState(ControlKey(ControlEnum.KeyDown).cAsciiKey) < 0 Then
        UpKey = False
        DownKey = True
        LeftKey = False
        RightKey = False
        Exit Sub
    Else
        DownKey = False
    End If
    
    'If GetKeyState(vbKeyLeft) < 0 Then
    If GetKeyState(ControlKey(ControlEnum.KeyLeft).cAsciiKey) < 0 Then
        UpKey = False
        DownKey = False
        LeftKey = True
        RightKey = False
        Exit Sub
    Else
        LeftKey = False
    End If
    
    'If GetKeyState(vbKeyRight) < 0 Then
    If GetKeyState(ControlKey(ControlEnum.KeyRight).cAsciiKey) < 0 Then
        UpKey = False
        DownKey = False
        LeftKey = False
        RightKey = True
        Exit Sub
    Else
        RightKey = False
    End If
    
    '//Constant
    If GetKeyState(vbKeyShift) < 0 Then
        ShiftKey = True
    Else
        ShiftKey = False
    End If
End Sub

Private Sub ResetButtonState(Optional ByVal Force As Boolean = False)
Dim i As Byte

    '//Reset all state of buttons
    For i = 1 To ButtonEnum.Button_Count - 1
        If Force Then
            Button(i).State = ButtonState.StateNormal
        Else
            If Button(i).State = ButtonState.StateHover Then
                Button(i).State = ButtonState.StateNormal
            End If
        End If
    Next
    
    '//Reset Mouse Icon
    If Not IsHovering Then MouseIcon = 0 '//Default
    
    '//Chatbox
    ChatScrollUp = False
    ChatScrollDown = False
    '//Pokedex Scroll
    PokedexScrollUp = False
    PokedexScrollDown = False
    
    ShopButtonState = 0
    ShopButtonHover = 0
End Sub

'//This handle the main form's key event
Public Sub FormKeyPress(KeyAscii As Integer)
Dim i As Long
Dim Slot As Long

    If Fade Then Exit Sub
    
    '//SelMenu
    If SelMenu.Visible Then Exit Sub
    
    '//Prioritize Inputbox
    If GUI(GuiEnum.GUI_CHOICEBOX).Visible Then Exit Sub
    
    '//zOrdering of gui
    If GUI(GuiEnum.GUI_INPUTBOX).Visible Then
        InputBoxKeyPress KeyAscii
    Else
        If Not IsLoading Then
            If GuiVisibleCount > 0 Then
                If CanShowGui(GuiZOrder(GuiVisibleCount)) Then
                    Select Case GuiZOrder(GuiVisibleCount)
                        Case GuiEnum.GUI_LOGIN: LoginKeyPress KeyAscii
                        Case GuiEnum.GUI_REGISTER: RegisterKeyPress KeyAscii
                        Case GuiEnum.GUI_CHARACTERCREATE: CharacterCreateKeyPress KeyAscii
                        Case GuiEnum.GUI_OPTION: OptionKeyPress KeyAscii
                        Case GuiEnum.GUI_CHATBOX: ChatboxKeyPress KeyAscii
                        Case GuiEnum.GUI_TRADE: TradeKeyPress KeyAscii
                    End Select
                End If
            End If
        End If
    End If
    
    If Not ChatOn Then
        If GameState = GameStateEnum.InGame Then
            For i = ControlEnum.KeyPokeSlot1 To ControlEnum.KeyPokeSlot6
                If KeyAscii = ControlKey(i).cAsciiKey Then
                    Slot = i - (ControlEnum.KeyPokeSlot1 - 1)
                    If Slot > 0 And Slot <= 6 Then
                        If PlayerPokemon(MyIndex).Num > 0 Then
                            '// Call Back
                            SendPlayerPokemonState 0, PlayerPokemon(MyIndex).Slot
                        Else
                            If SpawnTimer <= GetTickCount Then
                                If PlayerPokemons(Slot).Num > 0 Then
                                    If PlayerPokemons(Slot).CurHP > 0 Then
                                        If PlayerPokemons(Slot).Level <= (Player(MyIndex).Level + 10) Then
                                            SendPlayerPokemonState 1, Slot
                                            SpawnTimer = GetTickCount + 2000
                                        Else
                                            AddAlert "Not enough level", White
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End If
End Sub

'//This handle the main form's key event
Public Sub FormKeyUp(KeyCode As Integer, Shift As Integer)
Dim i As Long, zs As Long
Dim CanOpenMenu As Boolean
Dim Slot As Byte
Dim cIndex As Long

    If Fade Then Exit Sub
    
    '//SelMenu
    If SelMenu.Visible Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyEscape
            If Not GettingMap Then
                CanOpenMenu = True
                If GuiVisibleCount > 0 Then
                    Select Case GuiZOrder(GuiVisibleCount)
                        Case GuiEnum.GUI_INVENTORY
                            If GUI(GuiEnum.GUI_INVENTORY).Visible = True Then
                                GuiState GUI_INVENTORY, False
                                Button(ButtonEnum.Game_Bag).State = 0
                            End If
                            CanOpenMenu = False
                        Case GuiEnum.GUI_BADGE
                            If GUI(GuiEnum.GUI_BADGE).Visible = True Then
                                GuiState GUI_BADGE, False
                            End If
                            CanOpenMenu = False
                        Case GuiEnum.GUI_OPTION
                            If GUI(GuiEnum.GUI_OPTION).Visible = True Then
                                GuiState GUI_OPTION, False
                            End If
                            CanOpenMenu = False
                        Case GuiEnum.GUI_POKEDEX
                            If GUI(GuiEnum.GUI_POKEDEX).Visible = True Then
                                GuiState GUI_POKEDEX, False
                                Button(ButtonEnum.Game_Pokedex).State = 0
                            End If
                            CanOpenMenu = False
                        Case GuiEnum.GUI_TRAINER
                            If GUI(GuiEnum.GUI_TRAINER).Visible = True Then
                                GuiState GUI_TRAINER, False
                                Button(ButtonEnum.Game_Card).State = 0
                            End If
                            CanOpenMenu = False
                        Case GuiEnum.GUI_POKEMONSUMMARY
                            If GUI(GuiEnum.GUI_POKEMONSUMMARY).Visible = True Then
                                GuiState GUI_POKEMONSUMMARY, False
                            End If
                            CanOpenMenu = False
                    End Select
                End If
                    
                If CanOpenMenu Then
                    If Not GUI(GuiEnum.GUI_CHOICEBOX).Visible And Not GUI(GuiEnum.GUI_OPTION).Visible And Not GUI(GuiEnum.GUI_INPUTBOX).Visible Then
                        If GUI(GuiEnum.GUI_GLOBALMENU).Visible Then
                            GuiState GUI_GLOBALMENU, False
                        Else
                            GuiState GUI_GLOBALMENU, True
                        End If
                    End If
                End If
            End If
    End Select
    
    If GameState = GameStateEnum.InGame Then
        For i = ControlEnum.KeyHotbarSlot1 To ControlEnum.KeyHotbarSlot5
            If KeyCode = ControlKey(i).cAsciiKey Then
                Slot = i - (ControlEnum.KeyHotbarSlot1 - 1)
                If Slot > 0 And Slot <= MAX_HOTBAR Then
                    If Player(MyIndex).Hotbar(Slot) > 0 Then
                        SendUseHotbar Slot
                    End If
                End If
            End If
        Next
        
        If Not ChatOn Then
            For i = ControlEnum.KeyInventory To ControlEnum.KeyConvo4
                If KeyCode = ControlKey(i).cAsciiKey Then
                    Select Case i
                        Case ControlEnum.KeyInventory
                            If ShortKeyTimer <= GetTickCount Then
                                If GUI(GuiEnum.GUI_INVENTORY).Visible = False Then
                                    GuiState GuiEnum.GUI_INVENTORY, True
                                    '//Set to top most
                                    UpdateGuiOrder GUI_INVENTORY
                                Else
                                    GuiState GuiEnum.GUI_INVENTORY, False
                                    Button(ButtonEnum.Game_Bag).State = 0
                                End If
                                ShortKeyTimer = GetTickCount + 1000
                            End If
                        Case ControlEnum.KeyPokedex
                            If ShortKeyTimer <= GetTickCount Then
                                If GUI(GuiEnum.GUI_POKEDEX).Visible = False Then
                                    GuiState GuiEnum.GUI_POKEDEX, True
                                    '//Set to top most
                                    UpdateGuiOrder GUI_POKEDEX
                                Else
                                    GuiState GuiEnum.GUI_POKEDEX, False
                                    Button(ButtonEnum.Game_Pokedex).State = 0
                                End If
                                ShortKeyTimer = GetTickCount + 1000
                            End If
                        Case ControlEnum.KeyInteract
                            If ConvoNum > 0 Then
                                If Not ConvoShowButton Then
                                    If Len(ConvoText) > ConvoDrawTextLen Then
                                        ConvoDrawTextLen = Len(ConvoText)
                                        ConvoRenderText = Left$(ConvoText, ConvoDrawTextLen)
                                    Else
                                        '//Proceed to next convo
                                        If ConvoNoReply = YES Then
                                            '//Proceed to next
                                            SendProcessConvo
                                        Else
                                            ConvoShowButton = True
                                        End If
                                    End If
                                End If
                            Else
                                cIndex = FindFrontNPC
                                If cIndex > 0 Then
                                    SendConvo 1, cIndex
                                End If
                            End If
                        Case ControlEnum.KeyConvo1
                            If ConvoNum > 0 Then
                                If Len(ConvoText) > ConvoDrawTextLen Then
                                    ConvoDrawTextLen = Len(ConvoText)
                                    ConvoRenderText = Left$(ConvoText, ConvoDrawTextLen)
                                Else
                                    '//Proceed to next convo
                                    If ConvoNoReply = NO And ConvoShowButton Then
                                        '//Proceed to next
                                        SendProcessConvo 1
                                    Else
                                        ConvoShowButton = True
                                    End If
                                End If
                            End If
                        Case ControlEnum.KeyConvo2
                            If ConvoNum > 0 Then
                                If Len(ConvoText) > ConvoDrawTextLen Then
                                    ConvoDrawTextLen = Len(ConvoText)
                                    ConvoRenderText = Left$(ConvoText, ConvoDrawTextLen)
                                Else
                                    '//Proceed to next convo
                                    If ConvoNoReply = NO And ConvoShowButton Then
                                        '//Proceed to next
                                        SendProcessConvo 2
                                    Else
                                        ConvoShowButton = True
                                    End If
                                End If
                            End If
                        Case ControlEnum.KeyConvo3
                            If ConvoNum > 0 Then
                                If Len(ConvoText) > ConvoDrawTextLen Then
                                    ConvoDrawTextLen = Len(ConvoText)
                                    ConvoRenderText = Left$(ConvoText, ConvoDrawTextLen)
                                Else
                                    '//Proceed to next convo
                                    If ConvoNoReply = NO And ConvoShowButton Then
                                        '//Proceed to next
                                        SendProcessConvo 3
                                    Else
                                        ConvoShowButton = True
                                    End If
                                End If
                            End If
                        Case ControlEnum.KeyConvo4
                            If ConvoNum > 0 Then
                                If Len(ConvoText) > ConvoDrawTextLen Then
                                    ConvoDrawTextLen = Len(ConvoText)
                                    ConvoRenderText = Left$(ConvoText, ConvoDrawTextLen)
                                Else
                                    '//Proceed to next convo
                                    If ConvoNoReply = NO And ConvoShowButton Then
                                        '//Proceed to next
                                        SendProcessConvo 4
                                    Else
                                        ConvoShowButton = True
                                    End If
                                End If
                            End If
                    End Select
                End If
            Next
        End If
                
    End If
    
    '//zOrdering of gui
    If Not IsLoading Then
        If GuiVisibleCount > 0 Then
            If CanShowGui(GuiZOrder(GuiVisibleCount)) Then
                Select Case GuiZOrder(GuiVisibleCount)
                    Case GuiEnum.GUI_OPTION: OptionKeyUp KeyCode, Shift
                End Select
            End If
        End If
    End If
End Sub

'//This handle the main form's key event
Public Sub FormMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim DidClick As Boolean
Dim x2 As Long, Y2 As Long
Dim textX As Long, textY As Long
Dim PreventAction As Boolean

    If Fade Then Exit Sub
    
    ChatOn = False
    EditInputMoney = False
    editKey = 0
    EditTab = False
    MyChat = vbNullString
    
    '//SelMenu
    If SelMenu.Visible Then
        If SelMenuLogic(Buttons) Then
            Exit Sub
        End If
    End If
    
    '//GUI Priority
    '1st = Choice Box/Input Box
    '2nd = Global Menu
    '3rd = Other

    '//Choice Box must be above all gui if visible
    If GUI(GuiEnum.GUI_CHOICEBOX).Visible Then
        If CursorX >= GUI(GuiEnum.GUI_CHOICEBOX).X And CursorX <= GUI(GuiEnum.GUI_CHOICEBOX).X + GUI(GuiEnum.GUI_CHOICEBOX).Width And CursorY >= GUI(GuiEnum.GUI_CHOICEBOX).Y And CursorY <= GUI(GuiEnum.GUI_CHOICEBOX).Y + GUI(GuiEnum.GUI_CHOICEBOX).Height Then
            If Not DidClick Then
                ChoiceBoxMouseDown Buttons, Shift, X, Y
                DidClick = True
            End If
        End If
    ElseIf GUI(GuiEnum.GUI_INPUTBOX).Visible Then
        If CursorX >= GUI(GuiEnum.GUI_INPUTBOX).X And CursorX <= GUI(GuiEnum.GUI_INPUTBOX).X + GUI(GuiEnum.GUI_INPUTBOX).Width And CursorY >= GUI(GuiEnum.GUI_INPUTBOX).Y And CursorY <= GUI(GuiEnum.GUI_INPUTBOX).Y + GUI(GuiEnum.GUI_INPUTBOX).Height Then
            If Not DidClick Then
                InputBoxMouseDown Buttons, Shift, X, Y
                DidClick = True
            End If
        End If
    Else
        '//Global Menu must be above all gui except choice box
        If GUI(GuiEnum.GUI_GLOBALMENU).Visible Then
            If CursorX >= GUI(GuiEnum.GUI_GLOBALMENU).X And CursorX <= GUI(GuiEnum.GUI_GLOBALMENU).X + GUI(GuiEnum.GUI_GLOBALMENU).Width And CursorY >= GUI(GuiEnum.GUI_GLOBALMENU).Y And CursorY <= GUI(GuiEnum.GUI_GLOBALMENU).Y + GUI(GuiEnum.GUI_GLOBALMENU).Height Then
                If Not DidClick Then
                    GlobalMenuMouseDown Buttons, Shift, X, Y
                    DidClick = True
                End If
            End If
        ElseIf GUI(GuiEnum.GUI_OPTION).Visible Then
            If CursorX >= GUI(GuiEnum.GUI_OPTION).X And CursorX <= GUI(GuiEnum.GUI_OPTION).X + GUI(GuiEnum.GUI_OPTION).Width And CursorY >= GUI(GuiEnum.GUI_OPTION).Y And CursorY <= GUI(GuiEnum.GUI_OPTION).Y + GUI(GuiEnum.GUI_OPTION).Height Then
                If Not DidClick Then
                    OptionMouseDown Buttons, Shift, X, Y
                    DidClick = True
                End If
            End If
        Else
            If GUI(GuiEnum.GUI_CONVO).Visible Then
                If Not DidClick Then
                    '//Handle Convo
                    ConvoMouseDown Buttons, Shift, X, Y
                    DidClick = True
                End If
            Else
                '//zOrdering of gui
                If GuiVisibleCount > 0 Then
                    For i = GuiVisibleCount To 1 Step -1
                        If CanShowGui(GuiZOrder(i)) Then
                            If GuiZOrder(i) > 0 Then
                                If CursorX >= GUI(GuiZOrder(i)).X And CursorX <= GUI(GuiZOrder(i)).X + GUI(GuiZOrder(i)).Width And CursorY >= GUI(GuiZOrder(i)).Y And CursorY <= GUI(GuiZOrder(i)).Y + GUI(GuiZOrder(i)).Height Then
                                    Select Case GuiZOrder(i)
                                        Case GuiEnum.GUI_LOGIN
                                            If Not DidClick Then
                                                LoginMouseDown Buttons, Shift, X, Y
                                                DidClick = True
                                                Exit For
                                            End If
                                        Case GuiEnum.GUI_REGISTER
                                            If Not DidClick Then
                                                RegisterMouseDown Buttons, Shift, X, Y
                                                DidClick = True
                                                Exit For
                                            End If
                                        Case GuiEnum.GUI_CHARACTERSELECT
                                            If Not DidClick Then
                                                CharacterSelectMouseDown Buttons, Shift, X, Y
                                                DidClick = True
                                                Exit For
                                            End If
                                        Case GuiEnum.GUI_CHARACTERCREATE
                                            If Not DidClick Then
                                                CharacterCreateMouseDown Buttons, Shift, X, Y
                                                DidClick = True
                                                Exit For
                                            End If
                                        Case GuiEnum.GUI_CHATBOX
                                            If Not DidClick Then
                                                ChatBoxMouseDown Buttons, Shift, X, Y
                                                DidClick = True
                                                Exit For
                                            End If
                                        Case GuiEnum.GUI_INVENTORY
                                            If Not DidClick Then
                                                InventoryMouseDown Buttons, Shift, X, Y
                                                DidClick = True
                                                Exit For
                                            End If
                                        Case GuiEnum.GUI_MOVEREPLACE
                                            If Not DidClick Then
                                                MoveReplaceMouseDown Buttons, Shift, X, Y
                                                DidClick = True
                                                Exit For
                                            End If
                                        Case GuiEnum.GUI_TRAINER
                                            If Not DidClick Then
                                                TrainerMouseDown Buttons, Shift, X, Y
                                                DidClick = True
                                                Exit For
                                            End If
                                        Case GuiEnum.GUI_INVSTORAGE
                                            If Not DidClick Then
                                                InvStorageMouseDown Buttons, Shift, X, Y
                                                DidClick = True
                                                Exit For
                                            End If
                                        Case GuiEnum.GUI_POKEMONSTORAGE
                                            If Not DidClick Then
                                                PokemonStorageMouseDown Buttons, Shift, X, Y
                                                DidClick = True
                                                Exit For
                                            End If
                                        Case GuiEnum.GUI_SHOP
                                            If Not DidClick Then
                                                ShopMouseDown Buttons, Shift, X, Y
                                                DidClick = True
                                                Exit For
                                            End If
                                        Case GuiEnum.GUI_TRADE
                                            If Not DidClick Then
                                                TradeMouseDown Buttons, Shift, X, Y
                                                DidClick = True
                                                Exit For
                                            End If
                                        Case GuiEnum.GUI_POKEDEX
                                            If Not DidClick Then
                                                PokedexMouseDown Buttons, Shift, X, Y
                                                DidClick = True
                                                Exit For
                                            End If
                                        Case GuiEnum.GUI_POKEMONSUMMARY
                                            If Not DidClick Then
                                                PokemonSummaryMouseDown Buttons, Shift, X, Y
                                                DidClick = True
                                                Exit For
                                            End If
                                        Case GuiEnum.GUI_RELEARN
                                            If Not DidClick Then
                                                RelearnMouseDown Buttons, Shift, X, Y
                                                DidClick = True
                                                Exit For
                                            End If
                                        Case GuiEnum.GUI_BADGE
                                            If Not DidClick Then
                                                BadgeMouseDown Buttons, Shift, X, Y
                                                DidClick = True
                                                Exit For
                                            End If
                                        Case GuiEnum.GUI_SLOTMACHINE
                                            If Not DidClick Then
                                                SlotMachineMouseDown Buttons, Shift, X, Y
                                                DidClick = True
                                                Exit For
                                            End If
                                        Case GuiEnum.GUI_RANK
                                            If Not DidClick Then
                                                RankMouseDown Buttons, Shift, X, Y
                                                DidClick = True
                                                Exit For
                                            End If
                                    End Select
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        End If
    End If
    
    Select Case GameState
        Case GameStateEnum.InMenu
            If Not DidClick And Not GUI(GuiEnum.GUI_CHOICEBOX).Visible And Not GUI(GuiEnum.GUI_OPTION).Visible And Not GUI(GuiEnum.GUI_GLOBALMENU).Visible And Not GUI(GuiEnum.GUI_INPUTBOX).Visible Then
                '//Loop through all items
                For i = ButtonEnum.Menu_Website To ButtonEnum.Menu_ChangePass
                    If CanShowButton(i) Then
                        If CursorX >= Button(i).X And CursorX <= Button(i).X + Button(i).Width And CursorY >= Button(i).Y And CursorY <= Button(i).Y + Button(i).Height Then
                            If Button(i).State = ButtonState.StateHover Then
                                Button(i).State = ButtonState.StateClick
                            End If
                        End If
                    End If
                Next
                
                '//Credit Button
                textX = Screen_Width - 80
                textY = Screen_Height - 40
                If CursorX >= textX And CursorX <= textX + 80 And CursorY >= textY And CursorY <= textY + 40 Then
                    If CreditVisible Then
                        CreditState = 1
                    Else
                        CreditVisible = True
                        CreditOffset = 0
                        CreditState = 0
                        
                        For i = 0 To CreditTextCount
                            Credit(i).Y = Credit(i).StartY
                        Next
                    End If
                End If
            End If
        Case GameStateEnum.InGame
            If Not DidClick And Not GUI(GuiEnum.GUI_CHOICEBOX).Visible And Not GUI(GuiEnum.GUI_OPTION).Visible And Not GUI(GuiEnum.GUI_GLOBALMENU).Visible And Not GUI(GuiEnum.GUI_INPUTBOX).Visible And Not GUI(GuiEnum.GUI_CONVO).Visible Then
                If Buttons = vbRightButton Then
                    If Editor <> EDITOR_MAP Then
                        For i = 1 To MAX_PLAYER_POKEMON
                            If PlayerPokemon(MyIndex).Num > 0 Then
                                If PlayerPokemon(MyIndex).Slot = i Then
                                    If PlayerPokemons(i).Num > 0 Then
                                        x2 = Screen_Width - 34 - 5 - ((i - 1) * 40)
                                        Y2 = 62 ' + 52 + ((i - 1) * 40)
                                            
                                        If CursorX >= x2 And CursorX <= x2 + 34 And CursorY >= Y2 And CursorY <= Y2 + 37 Then
                                            SelPoke = i
                                            If PlayerPokemons(i).CurHP > 0 Then
                                                OpenSelMenu SelMenuType.PlayerPokes
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If PlayerPokemons(i).Num > 0 Then
                                    If PlayerPokemons(i).CurHP > 0 Then
                                        x2 = Screen_Width - 34 - 5 - ((i - 1) * 40)
                                        Y2 = 62 ' + 52 + ((i - 1) * 40)
                                            
                                        If CursorX >= x2 And CursorX <= x2 + 34 And CursorY >= Y2 And CursorY <= Y2 + 37 Then
                                            SelPoke = i
                                            OpenSelMenu SelMenuType.SpawnPokes
                                        End If
                                    Else
                                        x2 = Screen_Width - 34 - 5 - ((i - 1) * 40)
                                        Y2 = 62 ' + 52 + ((i - 1) * 40)
                                            
                                        If CursorX >= x2 And CursorX <= x2 + 34 And CursorY >= Y2 And CursorY <= Y2 + 37 Then
                                            SelPoke = i
                                            OpenSelMenu SelMenuType.RevivePokes
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
                
                For i = 1 To MAX_HOTBAR
                    x2 = Screen_Width - 42 - 170 - ((i - 1) * 45)
                    Y2 = 5 '62 + 37 + 5
                            
                    If CursorX >= x2 And CursorX <= x2 + 42 And CursorY >= Y2 And CursorY <= Y2 + 45 Then
                        If Buttons = vbRightButton Then
                            '//Remove Hotbar
                            SendHotbarUpdate i
                        End If
                    End If
                Next
            
                '//Editor Map
                If Buttons = vbLeftButton Or Buttons = vbRightButton Then
                    If Editor = EDITOR_MAP Then
                        MapEditorMouseDown Buttons
                    End If
                End If
            
                '//Admin Warp
                If Buttons = vbRightButton Then
                    If ShiftKey Then
                        If Editor <> EDITOR_MAP Then
                            If Player(MyIndex).Access > 0 Then
                                If PlayerPokemon(MyIndex).Num <= 0 Then
                                    AdminWarp curTileX, curTileY
                                End If
                            End If
                        End If
                    End If
                End If
            
                If Not Editor = EDITOR_MAP Then
                    If Not DidClick And Not GUI(GuiEnum.GUI_CHOICEBOX).Visible And Not GUI(GuiEnum.GUI_OPTION).Visible And Not GUI(GuiEnum.GUI_GLOBALMENU).Visible And Not GUI(GuiEnum.GUI_INPUTBOX).Visible Then
                        '//Loop through all items
                        For i = ButtonEnum.Game_Pokedex To ButtonEnum.Game_Evolve
                            If CanShowButton(i) Then
                                If CursorX >= Button(i).X And CursorX <= Button(i).X + Button(i).Width And CursorY >= Button(i).Y And CursorY <= Button(i).Y + Button(i).Height Then
                                    PreventAction = False
                                    Select Case i
                                        Case ButtonEnum.Game_Pokedex
                                            If GUI(GuiEnum.GUI_POKEDEX).Visible Then
                                                PreventAction = True
                                            End If
                                        Case ButtonEnum.Game_Bag
                                            If GUI(GuiEnum.GUI_INVENTORY).Visible Then
                                                PreventAction = True
                                            End If
                                        Case ButtonEnum.Game_Card
                                            If GUI(GuiEnum.GUI_TRAINER).Visible Then
                                                PreventAction = True
                                            End If
                                        Case ButtonEnum.Game_Clan
                                        Case ButtonEnum.Game_Task
                                        Case ButtonEnum.Game_Menu
                                            If GUI(GuiEnum.GUI_GLOBALMENU).Visible Then
                                                PreventAction = True
                                            End If
                                    End Select
                                    
                                    If Not PreventAction Then
                                        If Button(i).State = ButtonState.StateHover Then
                                            Button(i).State = ButtonState.StateClick
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            
                SearchMouseDown Buttons
            End If
    End Select
End Sub

'//This handle the main form's key event
Public Sub FormMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim PreventAction As Boolean

    '//Get Cursor Location
    CursorX = (Screen_Width / frmMain.scaleWidth) * X
    CursorY = (Screen_Height / frmMain.scaleHeight) * Y
    
    '//Make sure that the cursor position is always inside the screen
    If CursorX < 0 Then CursorX = 0
    If CursorY < 0 Then CursorY = 0
    If CursorX > Screen_Width - 16 Then CursorX = Screen_Width - 16
    If CursorY > Screen_Height - 16 Then CursorY = Screen_Height - 16
    
    '//Get Current Tile Location based on Cursor's Location
    curTileX = TileView.Left + ((CursorX + Camera.Left) \ TILE_X) 'CursorX \ Pic_X
    curTileY = TileView.top + ((CursorY + Camera.top) \ TILE_Y) 'CursorY \ Pic_Y

    '//Cursor Logic
    If InitCursorTimer Then
        If oldCursorX <> CursorX Or oldCursorY <> CursorY Then
            oldCursorX = CursorX
            oldCursorY = CursorY
            CursorTimer = GetTickCount + 20000
            CanShowCursor = True
        End If
    End If
    
    If InvItemDesc > 0 Then
        i = IsInvItem(CursorX, CursorY)
        If Not i = InvItemDesc Then
            InvItemDesc = 0
            InvItemDescTimer = 0
            InvItemDescShow = False
        End If
    End If
    
    WindowPriority = 0
    
    If Fade Then Exit Sub
    
    '//Reset button
    ResetButtonState
    
    '//SelMenu
    If SelMenu.Visible Then Exit Sub
    
    '//GUI Priority
    '1st = Choice Box/Input Box
    '2nd = Global Menu
    '3rd = Other
    
    '//Choice Box must be above all gui if visible
    If GUI(GuiEnum.GUI_CHOICEBOX).Visible Then
        ChoiceBoxMouseMove Buttons, Shift, X, Y
    ElseIf GUI(GuiEnum.GUI_INPUTBOX).Visible Then
        InputBoxMouseMove Buttons, Shift, X, Y
    Else
        If GUI(GuiEnum.GUI_GLOBALMENU).Visible Then
            GlobalMenuMouseMove Buttons, Shift, X, Y
        ElseIf GUI(GuiEnum.GUI_OPTION).Visible Then
            OptionMouseMove Buttons, Shift, X, Y
        Else
            If GUI(GuiEnum.GUI_CONVO).Visible Then
                ConvoMouseMove Buttons, Shift, X, Y
            Else
                '//zOrdering of gui
                If GuiVisibleCount > 0 Then
                    If CanShowGui(GuiZOrder(GuiVisibleCount)) Then
                        Select Case GuiZOrder(GuiVisibleCount)
                            Case GuiEnum.GUI_LOGIN:             LoginMouseMove Buttons, Shift, X, Y
                            Case GuiEnum.GUI_REGISTER:          RegisterMouseMove Buttons, Shift, X, Y
                            Case GuiEnum.GUI_CHARACTERSELECT:   CharacterSelectMouseMove Buttons, Shift, X, Y
                            Case GuiEnum.GUI_CHARACTERCREATE:   CharacterCreateMouseMove Buttons, Shift, X, Y
                            Case GuiEnum.GUI_CHATBOX:           ChatBoxMouseMove Buttons, Shift, X, Y
                            Case GuiEnum.GUI_INVENTORY
                                InventoryMouseMove Buttons, Shift, X, Y
                                If GUI(GuiEnum.GUI_INVSTORAGE).Visible Then
                                    InvStorageMouseMove Buttons, Shift, X, Y
                                End If
                            Case GuiEnum.GUI_MOVEREPLACE:       MoveReplaceMouseMove Buttons, Shift, X, Y
                            Case GuiEnum.GUI_TRAINER:           TrainerMouseMove Buttons, Shift, X, Y
                            Case GuiEnum.GUI_INVSTORAGE
                                InvStorageMouseMove Buttons, Shift, X, Y
                                If GUI(GuiEnum.GUI_INVENTORY).Visible Then
                                    InventoryMouseMove Buttons, Shift, X, Y
                                End If
                            Case GuiEnum.GUI_POKEMONSTORAGE:    PokemonStorageMouseMove Buttons, Shift, X, Y
                            Case GuiEnum.GUI_SHOP:              ShopMouseMove Buttons, Shift, X, Y
                            Case GuiEnum.GUI_TRADE:             TradeMouseMove Buttons, Shift, X, Y
                            Case GuiEnum.GUI_POKEDEX:           PokedexMouseMove Buttons, Shift, X, Y
                            Case GuiEnum.GUI_POKEMONSUMMARY:    PokemonSummaryMouseMove Buttons, Shift, X, Y
                            Case GuiEnum.GUI_RELEARN:           RelearnMouseMove Buttons, Shift, X, Y
                            Case GuiEnum.GUI_BADGE:             BadgeMouseMove Buttons, Shift, X, Y
                            Case GuiEnum.GUI_SLOTMACHINE:       SlotMachineMouseMove Buttons, Shift, X, Y
                            Case GuiEnum.GUI_RANK:              RankMouseMove Buttons, Shift, X, Y
                        End Select
                    End If
                End If
            End If
        End If
    End If
    
    Select Case GameState
        Case GameStateEnum.InMenu
            IsHovering = False
            '//Loop through all items
            If Not GUI(GuiEnum.GUI_CHOICEBOX).Visible And Not GUI(GuiEnum.GUI_GLOBALMENU).Visible And Not GUI(GuiEnum.GUI_OPTION).Visible And Not GUI(GuiEnum.GUI_INPUTBOX).Visible Then
                For i = ButtonEnum.Menu_Website To ButtonEnum.Menu_ChangePass
                    If CanShowButton(i) Then
                        If CursorX >= Button(i).X And CursorX <= Button(i).X + Button(i).Width And CursorY >= Button(i).Y And CursorY <= Button(i).Y + Button(i).Height Then
                            If Button(i).State = ButtonState.StateNormal Then
                                Button(i).State = ButtonState.StateHover
            
                                IsHovering = True
                                MouseIcon = 1 '//Select
                            End If
                        End If
                    End If
                Next
            End If
        Case GameStateEnum.InGame
            IsHovering = False
            
            '//Editor Map
            If Buttons = vbLeftButton Or Buttons = vbRightButton Then
                If Editor = EDITOR_MAP Then
                    MapEditorMouseDown Buttons
                End If
            End If
            
            '//Loop through all items
            If Not Editor = EDITOR_MAP Then
                If Not GUI(GuiEnum.GUI_CHOICEBOX).Visible And Not GUI(GuiEnum.GUI_GLOBALMENU).Visible And Not GUI(GuiEnum.GUI_OPTION).Visible And Not GUI(GuiEnum.GUI_INPUTBOX).Visible And Not GUI(GuiEnum.GUI_CONVO).Visible Then
                    For i = ButtonEnum.Game_Pokedex To ButtonEnum.Game_Evolve
                        If CanShowButton(i) Then
                            If CursorX >= Button(i).X And CursorX <= Button(i).X + Button(i).Width And CursorY >= Button(i).Y And CursorY <= Button(i).Y + Button(i).Height Then
                                PreventAction = False
                                Select Case i
                                    Case ButtonEnum.Game_Pokedex
                                        If GUI(GuiEnum.GUI_POKEDEX).Visible Then
                                            PreventAction = True
                                        End If
                                    Case ButtonEnum.Game_Bag
                                        If GUI(GuiEnum.GUI_INVENTORY).Visible Then
                                            PreventAction = True
                                        End If
                                    Case ButtonEnum.Game_Card
                                        If GUI(GuiEnum.GUI_TRAINER).Visible Then
                                            PreventAction = True
                                        End If
                                    Case ButtonEnum.Game_Clan
                                    Case ButtonEnum.Game_Task
                                    Case ButtonEnum.Game_Menu
                                        If GUI(GuiEnum.GUI_GLOBALMENU).Visible Then
                                            PreventAction = True
                                        End If
                                End Select
                                
                                If Not PreventAction Then
                                    If Button(i).State = ButtonState.StateNormal Then
                                        Button(i).State = ButtonState.StateHover
                    
                                        IsHovering = True
                                        MouseIcon = 1 '//Select
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End If
            
            SearchMouseMove Buttons
    End Select
End Sub

'//This handle the main form's key event
Public Sub FormMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim PreventAction As Boolean
Dim x2 As Long, Y2 As Long

    '//Chat Scroll
    ChatHold = False
    GuiPathEdit = False

    If Fade Then Exit Sub
    
    '//SelMenu
    If SelMenu.Visible Then Exit Sub

    '//GUI Priority
    '1st = Choice Box/Input Box
    '2nd = Global Menu
    '3rd = Other
    
    If GameState = GameStateEnum.InGame Then
        If DragInvSlot > 0 Then
            For i = 1 To MAX_HOTBAR
                x2 = Screen_Width - 42 - 170 - ((i - 1) * 45)
                Y2 = 5 '62 + 37 + 5
                    
                If CursorX >= x2 And CursorX <= x2 + 42 And CursorY >= Y2 And CursorY <= Y2 + 45 Then
                    SendHotbarUpdate i, DragInvSlot
                End If
            Next
        End If
    End If
    
    '//Choice Box must be above all gui if visible
    If GUI(GuiEnum.GUI_CHOICEBOX).Visible Then
        ChoiceBoxMouseUp Buttons, Shift, X, Y
    ElseIf GUI(GuiEnum.GUI_INPUTBOX).Visible Then
        InputBoxMouseUp Buttons, Shift, X, Y
    Else
        If GUI(GuiEnum.GUI_GLOBALMENU).Visible Then
            GlobalMenuMouseUp Buttons, Shift, X, Y
        ElseIf GUI(GuiEnum.GUI_OPTION).Visible Then
            OptionMouseUp Buttons, Shift, X, Y
        Else
            If GUI(GuiEnum.GUI_CONVO).Visible Then
                ConvoMouseUp Buttons, Shift, X, Y
            Else
                '//zOrdering of gui
                If GuiVisibleCount > 0 Then
                    If CanShowGui(GuiZOrder(GuiVisibleCount)) Then
                        If CursorX >= GUI(GuiZOrder(GuiVisibleCount)).X And CursorX <= GUI(GuiZOrder(GuiVisibleCount)).X + GUI(GuiZOrder(GuiVisibleCount)).Width And CursorY >= GUI(GuiZOrder(GuiVisibleCount)).Y And CursorY <= GUI(GuiZOrder(GuiVisibleCount)).Y + GUI(GuiZOrder(GuiVisibleCount)).Height Then
                            Select Case GuiZOrder(GuiVisibleCount)
                                Case GuiEnum.GUI_LOGIN:             LoginMouseUp Buttons, Shift, X, Y
                                Case GuiEnum.GUI_REGISTER:          RegisterMouseUp Buttons, Shift, X, Y
                                Case GuiEnum.GUI_CHARACTERSELECT:   CharacterSelectMouseUp Buttons, Shift, X, Y
                                Case GuiEnum.GUI_CHARACTERCREATE:   CharacterCreateMouseUp Buttons, Shift, X, Y
                                Case GuiEnum.GUI_CHATBOX:           ChatBoxMouseUp Buttons, Shift, X, Y
                                Case GuiEnum.GUI_INVENTORY:         InventoryMouseUp Buttons, Shift, X, Y
                                Case GuiEnum.GUI_MOVEREPLACE:       MoveReplaceMouseUp Buttons, Shift, X, Y
                                Case GuiEnum.GUI_TRAINER:           TrainerMouseUp Buttons, Shift, X, Y
                                Case GuiEnum.GUI_INVSTORAGE:        InvStorageMouseUp Buttons, Shift, X, Y
                                Case GuiEnum.GUI_POKEMONSTORAGE:    PokemonStorageMouseUp Buttons, Shift, X, Y
                                Case GuiEnum.GUI_SHOP:              ShopMouseUp Buttons, Shift, X, Y
                                Case GuiEnum.GUI_TRADE:             TradeMouseUp Buttons, Shift, X, Y
                                Case GuiEnum.GUI_POKEDEX:           PokedexMouseUp Buttons, Shift, X, Y
                                Case GuiEnum.GUI_POKEMONSUMMARY:    PokemonSummaryMouseUp Buttons, Shift, X, Y
                                Case GuiEnum.GUI_RELEARN:           RelearnMouseUp Buttons, Shift, X, Y
                                Case GuiEnum.GUI_BADGE:             BadgeMouseUp Buttons, Shift, X, Y
                                Case GuiEnum.GUI_SLOTMACHINE:       SlotMachineMouseUp Buttons, Shift, X, Y
                                Case GuiEnum.GUI_RANK:              RankMouseUp Buttons, Shift, X, Y
                            End Select
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    DragInvSlot = 0
    DragStorageSlot = 0
    DragPokeSlot = 0
    
    Select Case GameState
        Case GameStateEnum.InMenu
            If Not GUI(GuiEnum.GUI_CHOICEBOX).Visible And Not GUI(GuiEnum.GUI_GLOBALMENU).Visible And Not GUI(GuiEnum.GUI_OPTION).Visible And Not GUI(GuiEnum.GUI_INPUTBOX).Visible Then
                '//Loop through all items
                For i = ButtonEnum.Menu_Website To ButtonEnum.Menu_ChangePass
                    If CanShowButton(i) Then
                        If CursorX >= Button(i).X And CursorX <= Button(i).X + Button(i).Width And CursorY >= Button(i).Y And CursorY <= Button(i).Y + Button(i).Height Then
                            If Button(i).State = ButtonState.StateClick Then
                                Button(i).State = ButtonState.StateNormal
                                '//Do function of the button
                                Select Case i
                                    Case ButtonEnum.Menu_Website
                                        
                                    Case ButtonEnum.Menu_Register
                                        '//Temporary Closed for Closed Beta
                                        
                                        If Not GUI(GuiEnum.GUI_REGISTER).Visible Then
                                            GuiState GUI_LOGIN, False
                                            GuiState GUI_REGISTER, True, True
                                            CurTextbox = 1
                                            User = vbNullString
                                            Pass = vbNullString
                                            Pass2 = vbNullString
                                            Email = vbNullString
                                        End If
                                    Case ButtonEnum.Menu_ChangePass
                                        If Not GUI(GuiEnum.GUI_CHOICEBOX).Visible Then
                                            OpenInputBox "Enter your new password", IB_NEWPASSWORD
                                        End If
                                End Select
                            End If
                        End If
                    End If
                Next
            End If
        Case GameStateEnum.InGame
            If Not Editor = EDITOR_MAP Then
                If Not GUI(GuiEnum.GUI_CHOICEBOX).Visible And Not GUI(GuiEnum.GUI_GLOBALMENU).Visible And Not GUI(GuiEnum.GUI_OPTION).Visible And Not GUI(GuiEnum.GUI_INPUTBOX).Visible And Not GUI(GuiEnum.GUI_CONVO).Visible Then
                    '//Loop through all items
                    For i = ButtonEnum.Game_Pokedex To ButtonEnum.Game_Evolve
                        If CanShowButton(i) Then
                            If CursorX >= Button(i).X And CursorX <= Button(i).X + Button(i).Width And CursorY >= Button(i).Y And CursorY <= Button(i).Y + Button(i).Height Then
                                PreventAction = False
                                Select Case i
                                    Case ButtonEnum.Game_Pokedex
                                        If GUI(GuiEnum.GUI_POKEDEX).Visible Then
                                            PreventAction = True
                                        End If
                                    Case ButtonEnum.Game_Bag
                                        If GUI(GuiEnum.GUI_INVENTORY).Visible Then
                                            PreventAction = True
                                        End If
                                    Case ButtonEnum.Game_Card
                                        If GUI(GuiEnum.GUI_TRAINER).Visible Then
                                            PreventAction = True
                                        End If
                                    Case ButtonEnum.Game_Clan
                                    Case ButtonEnum.Game_Task
                                    Case ButtonEnum.Game_Menu
                                        If GUI(GuiEnum.GUI_GLOBALMENU).Visible Then
                                            PreventAction = True
                                        End If
                                End Select
                                
                                If Not PreventAction Then
                                    If Button(i).State = ButtonState.StateClick Then
                                        Button(i).State = ButtonState.StateNormal
                                        '//Do function of the button
                                        Select Case i
                                            Case ButtonEnum.Game_Pokedex
                                                If GUI(GuiEnum.GUI_POKEDEX).Visible = False Then
                                                    GuiState GUI_POKEDEX, True
                                                End If
                                            Case ButtonEnum.Game_Bag
                                                If GUI(GuiEnum.GUI_INVENTORY).Visible = False Then
                                                    GuiState GUI_INVENTORY, True
                                                End If
                                            Case ButtonEnum.Game_Card
                                                If GUI(GuiEnum.GUI_TRAINER).Visible = False Then
                                                    GuiState GUI_TRAINER, True
                                                End If
                                            Case ButtonEnum.Game_Task
                                            Case ButtonEnum.Game_Clan
                                            Case ButtonEnum.Game_Menu
                                                If Not GUI(GuiEnum.GUI_CHOICEBOX).Visible And Not GUI(GuiEnum.GUI_OPTION).Visible And Not GUI(GuiEnum.GUI_INPUTBOX).Visible Then
                                                    If GUI(GuiEnum.GUI_GLOBALMENU).Visible Then
                                                        GuiState GUI_GLOBALMENU, False
                                                    Else
                                                        GuiState GUI_GLOBALMENU, True
                                                    End If
                                                End If
                                            Case ButtonEnum.Game_Evolve
                                                OpenSelMenu SelMenuType.Evolve
                                            Case Else
                                        End Select
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            End If
    End Select
    
    '//Reset button
    ResetButtonState True
    '//reset dragging Gui
    If GuiVisibleCount > 0 Then
        For i = 1 To GuiVisibleCount
            If GuiZOrder(i) > 0 Then
                GUI(GuiZOrder(i)).InDrag = False
                GUI(GuiZOrder(i)).OldMouseX = 0
                GUI(GuiZOrder(i)).OldMouseY = 0
            End If
        Next
    End If
End Sub

' ***********
' ** Login **
' ***********
Private Sub LoginKeyPress(KeyAscii As Integer)
Dim FoundError As Boolean

    '//Make sure it's visible
    If Not GUI(GuiEnum.GUI_LOGIN).Visible Then Exit Sub
    
    If GuiVisibleCount <= 0 Then Exit Sub
    If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_LOGIN Then Exit Sub

    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        If CurTextbox = 1 Then
            CurTextbox = 2
        ElseIf CurTextbox = 2 Then
            If KeyAscii = vbKeyReturn Then
                '//Check if we properly input the required form
                FoundError = False
                
                If WaitTimer > GetTickCount Then
                    AddAlert "Please wait for a few second before trying again", White
                    FoundError = True
                End If
                
                '//Textbox 1
                If Not FoundError And Not CheckNameInput(User, False, (NAME_LENGTH - 1)) Then
                    CurTextbox = 1
                    AddAlert "Invalid Username", White
                    FoundError = True
                End If
                
                If Not FoundError And Not CheckNameInput(Pass, False, (NAME_LENGTH - 1)) Then
                    CurTextbox = 2
                    AddAlert "Invalid Password", White
                    FoundError = True
                End If
                
                '//No Error Found
                If Not FoundError Then
                    '//Send account information
                    Menu_State MENU_STATE_LOGIN
                    '//Prevent Spamming
                    WaitTimer = GetTickCount + 5000
                End If
            Else
                CurTextbox = 1
            End If
        End If
    End If
    
    Select Case CurTextbox
        Case 1: If (isNameLegal(KeyAscii, True) And Len(User) < (NAME_LENGTH - 1)) Or KeyAscii = vbKeyBack Then User = InputText(User, KeyAscii)
        Case 2: If (isNameLegal(KeyAscii, True) And Len(Pass) < (NAME_LENGTH - 1)) Or KeyAscii = vbKeyBack Then Pass = InputText(Pass, KeyAscii)
    End Select
End Sub

Private Sub LoginMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte

    With GUI(GuiEnum.GUI_LOGIN)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_LOGIN
        
        '//Loop through all items
        If Not ServerList Then
            For i = ButtonEnum.Login_Confirm To ButtonEnum.Login_Confirm
                If CanShowButton(i) Then
                    If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateHover Then
                            Button(i).State = ButtonState.StateClick
                        End If
                    End If
                End If
            Next
        End If
        
        '//Clicking Textbox
        If CursorX >= .X + 86 And CursorX <= .X + 86 + 189 And CursorY >= .Y + 45 And CursorY <= .Y + 45 + 21 Then
            CurTextbox = 1
        ElseIf CursorX >= .X + 86 And CursorX <= .X + 86 + 189 And CursorY >= .Y + 78 And CursorY <= .Y + 78 + 21 Then
            CurTextbox = 2
        End If
        
        '//Checkbox
        If CursorX >= .X + 65 And CursorX <= .X + 65 + 17 And CursorY >= .Y + 140 And CursorY <= .Y + 140 + 17 Then
            If GameSetting.SavePass = YES Then
                GameSetting.SavePass = NO
            Else
                GameSetting.SavePass = YES
            End If
        End If
        
        If ServerList Then
            For i = 1 To MAX_SERVER_LIST
                If CursorX >= .X + 20 + 110 And CursorX <= .X + 20 + 110 + 140 And CursorY >= .Y + 112 + ((20 * MAX_SERVER_LIST) + ((i - 2) * 20)) And CursorY <= .Y + 112 + ((20 * MAX_SERVER_LIST) + ((i - 2) * 20)) + 20 Then
                    CurServerList = i
                    LoadServerList CurServerList
                    Exit For
                End If
            Next
            ServerList = False
        Else
            If CursorX >= .X + 111 And CursorX <= .X + 111 + 162 And CursorY >= .Y + 111 And CursorY <= .Y + 111 + 21 Then
                ServerList = True
            End If
        End If
        
        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Private Sub LoginMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte
Dim tmpX As Long, tmpY As Long

    With GUI(GuiEnum.GUI_LOGIN)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_LOGIN Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        If Not ServerList Then
            For i = ButtonEnum.Login_Confirm To ButtonEnum.Login_Confirm
                If CanShowButton(i) Then
                    If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateNormal Then
                            Button(i).State = ButtonState.StateHover
                            
                            IsHovering = True
                            MouseIcon = 1 '//Select
                        End If
                    End If
                End If
            Next
        End If
        
        '//Hovering Textbox
        If CursorX >= .X + 86 And CursorX <= .X + 86 + 189 And CursorY >= .Y + 45 And CursorY <= .Y + 45 + 21 Then
            IsHovering = True
            MouseIcon = 2 '//I-Beam
        ElseIf CursorX >= .X + 86 And CursorX <= .X + 86 + 189 And CursorY >= .Y + 78 And CursorY <= .Y + 78 + 21 Then
            IsHovering = True
            MouseIcon = 2 '//I-Beam
        End If
        
        '//Checkbox
        If CursorX >= .X + 65 And CursorX <= .X + 65 + 17 And CursorY >= .Y + 140 And CursorY <= .Y + 140 + 17 Then
            IsHovering = True
            MouseIcon = 1 '//Select
        End If
        
        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Private Sub LoginMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte
Dim FoundError As Boolean

    With GUI(GuiEnum.GUI_LOGIN)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_LOGIN Then Exit Sub
        
        '//Loop through all items
        If Not ServerList Then
            For i = ButtonEnum.Login_Confirm To ButtonEnum.Login_Confirm
                If CanShowButton(i) Then
                    If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateClick Then
                            Button(i).State = ButtonState.StateNormal
                            '//Do function of the button
                            Select Case i
                                Case ButtonEnum.Login_Confirm
                                    '//Check if we properly input the required form
                                    FoundError = False
                                    
                                    If WaitTimer > GetTickCount Then
                                        AddAlert "Please wait for a few second before trying again", White
                                        FoundError = True
                                    End If
                                    
                                    '//Textbox 1
                                    If Not FoundError And Not CheckNameInput(User, False, (NAME_LENGTH - 1)) Then
                                        CurTextbox = 1
                                        AddAlert "Invalid Username", White
                                        FoundError = True
                                    End If
                                    
                                    If Not FoundError And Not CheckNameInput(Pass, False, (NAME_LENGTH - 1)) Then
                                        CurTextbox = 2
                                        AddAlert "Invalid Password", White
                                        FoundError = True
                                    End If
                                    
                                    '//No Error Found
                                    If Not FoundError Then
                                        '//Send account information
                                        Menu_State MENU_STATE_LOGIN
                                        '//Prevent Spamming
                                        WaitTimer = GetTickCount + 5000
                                    End If
                            End Select
                        End If
                    End If
                End If
            Next
        End If
        
        '//Check for dragging
        .InDrag = False
    End With
End Sub

' **************
' ** Register **
' **************
Private Sub RegisterKeyPress(KeyAscii As Integer)
Dim FoundError As Boolean

    '//Make sure it's visible
    If Not GUI(GuiEnum.GUI_REGISTER).Visible Then Exit Sub
    
    If GuiVisibleCount <= 0 Then Exit Sub
    If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_REGISTER Then Exit Sub

    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        If CurTextbox = 1 Then
            CurTextbox = 2
        ElseIf CurTextbox = 2 Then
            CurTextbox = 3
        ElseIf CurTextbox = 3 Then
            CurTextbox = 4
        ElseIf CurTextbox = 4 Then
            If KeyAscii = vbKeyReturn Then
                '//Check if we properly input the required form
                FoundError = False
                
                If WaitTimer > GetTickCount Then
                    AddAlert "Please wait for a few second before trying again", White
                    FoundError = True
                End If
                
                '//Textbox 1
                If Not FoundError And Not CheckNameInput(User, False, (NAME_LENGTH - 1)) Then
                    CurTextbox = 1
                    AddAlert "Your username must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                    FoundError = True
                End If
                
                If Not FoundError And Not CheckNameInput(Pass, False, (NAME_LENGTH - 1)) Then
                    CurTextbox = 2
                    AddAlert "Your password must be between " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                    FoundError = True
                End If
                
                If Not FoundError And (Pass <> Pass2) Then
                    CurTextbox = 3
                    AddAlert "Password did not match", White
                    FoundError = True
                End If
                
                If Not FoundError And Not CheckNameInput(Email, False, (TEXT_LENGTH - 1), True) Then
                    CurTextbox = 4
                    AddAlert "Invalid email", White
                    FoundError = True
                End If
                
                '//No Error Found
                If Not FoundError Then
                    '//Send account information
                    Menu_State MENU_STATE_REGISTER
                    '//Prevent Spamming
                    WaitTimer = GetTickCount + 5000
                End If
            Else
                CurTextbox = 1
            End If
        End If
    End If
    
    Select Case CurTextbox
        Case 1: If (isNameLegal(KeyAscii, True) And Len(User) < (NAME_LENGTH - 1)) Or KeyAscii = vbKeyBack Then User = InputText(User, KeyAscii)
        Case 2: If (isNameLegal(KeyAscii, True) And Len(Pass) < (NAME_LENGTH - 1)) Or KeyAscii = vbKeyBack Then Pass = InputText(Pass, KeyAscii)
        Case 3: If (isNameLegal(KeyAscii, True) And Len(Pass2) < (NAME_LENGTH - 1)) Or KeyAscii = vbKeyBack Then Pass2 = InputText(Pass2, KeyAscii)
        Case 4: If (isStringLegal(KeyAscii, True) And Len(Email) < (TEXT_LENGTH - 1)) Or KeyAscii = vbKeyBack Then Email = InputText(Email, KeyAscii)
    End Select
End Sub

Private Sub RegisterMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte

    With GUI(GuiEnum.GUI_REGISTER)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_REGISTER
        
        '//Loop through all items
        For i = ButtonEnum.Register_Confirm To ButtonEnum.Register_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
        
        '//Clicking Textbox
        If CursorX >= .X + 86 And CursorX <= .X + 86 + 189 And CursorY >= .Y + 45 And CursorY <= .Y + 45 + 21 Then
            CurTextbox = 1
        ElseIf CursorX >= .X + 86 And CursorX <= .X + 86 + 189 And CursorY >= .Y + 78 And CursorY <= .Y + 78 + 21 Then
            CurTextbox = 2
        ElseIf CursorX >= .X + 86 And CursorX <= .X + 86 + 189 And CursorY >= .Y + 99 And CursorY <= .Y + 99 + 21 Then
            CurTextbox = 3
        ElseIf CursorX >= .X + 86 And CursorX <= .X + 86 + 189 And CursorY >= .Y + 166 And CursorY <= .Y + 166 + 21 Then
            CurTextbox = 4
        End If
        
        '//Checking of Checkbox
        If CursorX >= .X + 65 And CursorX <= .X + 65 + 17 And CursorY >= .Y + 137 And CursorY <= .Y + 137 + 17 Then
            If ShowPass = YES Then
                ShowPass = NO
            Else
                ShowPass = YES
            End If
        End If
        
        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Private Sub RegisterMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte
Dim tmpX As Long, tmpY As Long

    With GUI(GuiEnum.GUI_REGISTER)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_REGISTER Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        For i = ButtonEnum.Register_Confirm To ButtonEnum.Register_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
                        
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next
        
        '//Hovering Textbox
        If CursorX >= .X + 86 And CursorX <= .X + 86 + 189 And CursorY >= .Y + 45 And CursorY <= .Y + 45 + 21 Then
            IsHovering = True
            MouseIcon = 2 '//I-Beam
        ElseIf CursorX >= .X + 86 And CursorX <= .X + 86 + 189 And CursorY >= .Y + 78 And CursorY <= .Y + 78 + 21 Then
            IsHovering = True
            MouseIcon = 2 '//I-Beam
        ElseIf CursorX >= .X + 86 And CursorX <= .X + 86 + 189 And CursorY >= .Y + 99 And CursorY <= .Y + 99 + 21 Then
            IsHovering = True
            MouseIcon = 2 '//I-Beam
        ElseIf CursorX >= .X + 86 And CursorX <= .X + 86 + 189 And CursorY >= .Y + 166 And CursorY <= .Y + 166 + 21 Then
            IsHovering = True
            MouseIcon = 2 '//I-Beam
        End If
        
        '//Checkbox
        If CursorX >= .X + 65 And CursorX <= .X + 65 + 17 And CursorY >= .Y + 137 And CursorY <= .Y + 137 + 17 Then
            IsHovering = True
            MouseIcon = 1 '//Select
        End If
        
        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Private Sub RegisterMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte
Dim FoundError As Boolean

    With GUI(GuiEnum.GUI_REGISTER)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_REGISTER Then Exit Sub
        
        '//Loop through all items
        For i = ButtonEnum.Register_Confirm To ButtonEnum.Register_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        '//Do function of the button
                        Select Case i
                            Case ButtonEnum.Register_Confirm
                                '//Check if we properly input the required form
                                FoundError = False
                                
                                If WaitTimer > GetTickCount Then
                                    AddAlert "Please wait for a few second before trying again", White
                                    FoundError = True
                                End If
                                
                                '//Textbox 1
                                If Not FoundError And Not CheckNameInput(User, False, (NAME_LENGTH - 1)) Then
                                    CurTextbox = 1
                                    AddAlert "Your username must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                                    FoundError = True
                                End If
                                
                                If Not FoundError And Not CheckNameInput(Pass, False, (NAME_LENGTH - 1)) Then
                                    CurTextbox = 2
                                    AddAlert "Your password must be between " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                                    FoundError = True
                                End If
                                
                                If Not FoundError And (Pass <> Pass2) Then
                                    CurTextbox = 3
                                    AddAlert "Password did not match", White
                                    FoundError = True
                                End If
                                
                                If Not FoundError And Not CheckNameInput(Email, False, (TEXT_LENGTH - 1), True) Then
                                    CurTextbox = 4
                                    AddAlert "Invalid email", White
                                    FoundError = True
                                End If
                                
                                '//No Error Found
                                If Not FoundError Then
                                    '//Send account information
                                    Menu_State MENU_STATE_REGISTER
                                    '//Prevent Spamming
                                    WaitTimer = GetTickCount + 5000
                                End If
                            Case ButtonEnum.Register_Close
                                GuiState GUI_REGISTER, False
                                GuiState GUI_LOGIN, True, True
                                CurTextbox = 1
                        End Select
                    End If
                End If
            End If
        Next
        
        '//Check for dragging
        .InDrag = False
    End With
End Sub

' *********************
' ** CharacterSelect **
' *********************
Private Sub CharacterSelectMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_CHARACTERSELECT)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_CHARACTERSELECT
        
        '//Loop through all items
        For i = ButtonEnum.Character_SwitchLeft To ButtonEnum.Character_Delete
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
        
        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Private Sub CharacterSelectMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim tmpX As Long, tmpY As Long

    With GUI(GuiEnum.GUI_CHARACTERSELECT)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_CHARACTERSELECT Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        For i = ButtonEnum.Character_SwitchLeft To ButtonEnum.Character_Delete
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
                        
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next

        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Private Sub CharacterSelectMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_CHARACTERSELECT)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_CHARACTERSELECT Then Exit Sub
        
        '//Loop through all items
        For i = ButtonEnum.Character_SwitchLeft To ButtonEnum.Character_Delete
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        '//Do function of the button
                        Select Case i
                            Case ButtonEnum.Character_SwitchLeft
                                If CurChar > 1 Then
                                    CurChar = CurChar - 1
                                End If
                            Case ButtonEnum.Character_SwitchRight
                                If CurChar < MAX_PLAYERCHAR Then
                                    CurChar = CurChar + 1
                                End If
                            Case ButtonEnum.Character_New
                                '//Reset Character Create Data
                                CharName = vbNullString
                                SelGender = 0

                                GuiState GUI_CHARACTERSELECT, False
                                GuiState GUI_CHARACTERCREATE, True
                            Case ButtonEnum.Character_Use
                                If WaitTimer > GetTickCount Then
                                    AddAlert "Please wait for a few second before trying again", White
                                Else
                                    Menu_State MENU_STATE_USECHAR
                                    '//Prevent Spamming
                                    WaitTimer = GetTickCount + 5000
                                End If
                            Case ButtonEnum.Character_Delete
                                OpenChoiceBox "Are you sure that you want to delete this character?", CB_CHARDEL
                        End Select
                    End If
                End If
            End If
        Next
        
        '//Check for dragging
        .InDrag = False
    End With
End Sub

' *********************
' ** CharacterCreate **
' *********************
Private Sub CharacterCreateKeyPress(KeyAscii As Integer)
    '//Make sure it's visible
    If Not GUI(GuiEnum.GUI_CHARACTERCREATE).Visible Then Exit Sub
    
    If GuiVisibleCount <= 0 Then Exit Sub
    If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_CHARACTERCREATE Then Exit Sub
    
    If (isNameLegal(KeyAscii, True) And Len(CharName) < (NAME_LENGTH - 1)) Or KeyAscii = vbKeyBack Then CharName = InputText(CharName, KeyAscii)
End Sub

Private Sub CharacterCreateMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_CHARACTERCREATE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_CHARACTERCREATE
        
        '//Loop through all items
        For i = ButtonEnum.CharCreate_Confirm To ButtonEnum.CharCreate_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
        
        '//Gender
        If CursorY >= .Y + 78 And CursorY <= .Y + 78 + 87 Then
            If CursorX >= .X + 26 And CursorX <= .X + 26 + 114 Then
                SelGender = GENDER_MALE
            ElseIf CursorX >= .X + 142 And CursorX <= .X + 142 + 114 Then
                SelGender = GENDER_FEMALE
            End If
        End If
        
        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Private Sub CharacterCreateMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim tmpX As Long, tmpY As Long

    With GUI(GuiEnum.GUI_CHARACTERCREATE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_CHARACTERCREATE Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        For i = ButtonEnum.CharCreate_Confirm To ButtonEnum.CharCreate_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
        
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next
        
        '//Gender
        If CursorY >= .Y + 78 And CursorY <= .Y + 78 + 87 Then
            If CursorX >= .X + 26 And CursorX <= .X + 26 + 114 Then
                IsHovering = True
                MouseIcon = 1 '//Select
            ElseIf CursorX >= .X + 142 And CursorX <= .X + 142 + 114 Then
                IsHovering = True
                MouseIcon = 1 '//Select
            End If
        End If

        '//Hovering Textbox
        If CursorX >= .X + 86 And CursorX <= .X + 86 + 167 And CursorY >= .Y + 45 And CursorY <= .Y + 45 + 22 Then
            IsHovering = True
            MouseIcon = 2 '//I-Beam
        End If

        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Private Sub CharacterCreateMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim FoundError As Boolean

    With GUI(GuiEnum.GUI_CHARACTERCREATE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_CHARACTERCREATE Then Exit Sub
        
        '//Loop through all items
        For i = ButtonEnum.CharCreate_Confirm To ButtonEnum.CharCreate_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        '//Do function of the button
                        Select Case i
                            Case ButtonEnum.CharCreate_Confirm
                                '//Check if we properly input the required form
                                FoundError = False
                                
                                If WaitTimer > GetTickCount Then
                                    AddAlert "Please wait for a few second before trying again", White
                                    FoundError = True
                                End If

                                If Not FoundError And Not CheckNameInput(CharName, False, (NAME_LENGTH - 1)) Then
                                    CurTextbox = 1
                                    AddAlert "Your character name must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers, spaces, and _ allowed", White
                                    FoundError = True
                                End If
                                
                                '//Some Bug Error
                                If CurChar = 0 Then
                                    GuiState GUI_CHARACTERCREATE, False
                                    GuiState GUI_CHARACTERSELECT, True
                                End If
    
                                '//No Error Found
                                If Not FoundError Then
                                    '//Send account information
                                    Menu_State MENU_STATE_ADDCHAR
                                    '//Prevent Spamming
                                    WaitTimer = GetTickCount + 5000
                                End If
                            Case ButtonEnum.CharCreate_Close
                                GuiState GUI_CHARACTERCREATE, False
                                GuiState GUI_CHARACTERSELECT, True
                        End Select
                    End If
                End If
            End If
        Next
        
        '//Check for dragging
        .InDrag = False
    End With
End Sub

' ***************
' ** ChoiceBox **
' ***************
Private Sub ChoiceBoxMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_CHOICEBOX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_CHOICEBOX
        
        '//Loop through all items
        For i = ButtonEnum.ChoiceBox_Yes To ButtonEnum.ChoiceBox_No
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Sub ChoiceBoxMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_CHOICEBOX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If GuiVisibleCount <= 0 Then Exit Sub

        IsHovering = False

        '//Loop through all items
        For i = ButtonEnum.ChoiceBox_Yes To ButtonEnum.ChoiceBox_No
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
        
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Sub ChoiceBoxMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim z As Long

    With GUI(GuiEnum.GUI_CHOICEBOX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If GuiVisibleCount <= 0 Then Exit Sub

        '//Loop through all items
        For i = ButtonEnum.ChoiceBox_Yes To ButtonEnum.ChoiceBox_No
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        '//Do function of the button
                        Select Case i
                            Case ButtonEnum.ChoiceBox_Yes
                                Select Case ChoiceBoxType
                                    Case CB_EXIT
                                        If GameState = GameStateEnum.InMenu Then
                                            '//Exit
                                            CloseChoiceBox
                                            UnloadMain
                                        ElseIf GameState = GameStateEnum.InGame Then
                                            GettingMap = True
                                            CloseChoiceBox
                                            InitFade 0, FadeIn, 5
                                        End If
                                    Case CB_CHARDEL
                                        Menu_State MENU_STATE_DELCHAR
                                        CloseChoiceBox
                                    Case CB_RETURNMENU
                                        If GameState = GameStateEnum.InMenu Then
                                            CloseChoiceBox
                                            ResetMenu
                                        ElseIf GameState = GameStateEnum.InGame Then
                                            GettingMap = True
                                            CloseChoiceBox
                                            InitFade 0, FadeIn, 6
                                        End If
                                    Case CB_SAVESETTING
                                        '//Save setting
                                        SaveSettingConfiguration
                                        '//Save Key Input
                                        For z = 1 To ControlEnum.Control_Count - 1
                                            ControlKey(z).cAsciiKey = TmpKey(z)
                                        Next
                                        SaveControlKey
                                        
                                        CloseChoiceBox
                                        If GUI(GuiEnum.GUI_OPTION).Visible Then
                                            GuiState GUI_OPTION, False
                                        End If
                                    Case CB_EVOLVE
                                        '//Do Evolve
                                        SendEvolvePoke EvolveSelect
                                        CloseChoiceBox
                                    Case CB_REQUEST
                                        '//Accept Request
                                        SendRequestState 1
                                        CloseChoiceBox
                                    Case CB_RELEASE
                                        '//release pokemon
                                        SendReleasePokemon ReleaseStorageSlot, ReleaseStorageData
                                        ReleaseStorageSlot = 0
                                        ReleaseStorageData = 0
                                        CloseChoiceBox
                                    Case CB_BUYSLOT
                                        '//Buy Slot
                                        SendBuyStorageSlot BuySlotType, BuySlotData
                                        BuySlotType = 0
                                        BuySlotData = 0
                                        CloseChoiceBox
                                    Case CB_FLY
                                        SendFlyToBadge FlyBadgeSlot
                                        FlyBadgeSlot = 0
                                        CloseChoiceBox
                                End Select
                            Case ButtonEnum.ChoiceBox_No
                                Select Case ChoiceBoxType
                                    Case CB_SAVESETTING
                                        CloseChoiceBox
                                        If GUI(GuiEnum.GUI_OPTION).Visible Then
                                            GuiState GUI_OPTION, False
                                        End If
                                    Case CB_REQUEST
                                        '//Decline request
                                        SendRequestState 2
                                        CloseChoiceBox
                                    Case CB_RELEASE
                                        ReleaseStorageSlot = 0
                                        ReleaseStorageData = 0
                                        CloseChoiceBox
                                    Case CB_BUYSLOT
                                        BuySlotType = 0
                                        BuySlotData = 0
                                        CloseChoiceBox
                                    Case CB_FLY
                                        FlyBadgeSlot = 0
                                        CloseChoiceBox
                                    Case Else
                                        CloseChoiceBox
                                End Select
                        End Select
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Sub CloseChoiceBox()
    GuiState GUI_CHOICEBOX, False
    ChoiceBoxText = vbNullString
    ChoiceBoxType = 0
    EvolveSelect = 0
End Sub

' ***************
' ** GlobalMenu **
' ***************
Private Sub GlobalMenuMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_GLOBALMENU)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_GLOBALMENU
        
        '//Loop through all items
        For i = ButtonEnum.GlobalMenu_Return To ButtonEnum.GlobalMenu_Exit
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Sub GlobalMenuMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_GLOBALMENU)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If GuiVisibleCount <= 0 Then Exit Sub

        IsHovering = False

        '//Loop through all items
        For i = ButtonEnum.GlobalMenu_Return To ButtonEnum.GlobalMenu_Exit
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
            
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Sub GlobalMenuMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_GLOBALMENU)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If GuiVisibleCount <= 0 Then Exit Sub

        '//Loop through all items
        For i = ButtonEnum.GlobalMenu_Return To ButtonEnum.GlobalMenu_Exit
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        '//Do function of the button
                        Select Case i
                            Case ButtonEnum.GlobalMenu_Return
                                GuiState GUI_GLOBALMENU, False
                            Case ButtonEnum.GlobalMenu_Option
                                GuiState GUI_GLOBALMENU, False
                                GuiState GUI_OPTION, True
                                InitSettingConfiguration
                            Case ButtonEnum.GlobalMenu_Back
                                If GameState = GameStateEnum.InMenu Then
                                    GuiState GUI_GLOBALMENU, False
                                ElseIf GameState = GameStateEnum.InGame Then
                                    GuiState GUI_GLOBALMENU, False
                                    OpenChoiceBox "Are you sure you want to return to main menu?", CB_RETURNMENU
                                End If
                            Case ButtonEnum.GlobalMenu_Exit
                                GuiState GUI_GLOBALMENU, False
                                OpenChoiceBox "Are you sure you want to exit the game?", CB_EXIT
                        End Select
                    End If
                End If
            End If
        Next
    End With
End Sub

' ***************
' ** Option **
' ***************
Private Sub OptionKeyPress(KeyAscii As Integer)
    '//Make sure it's visible
    If Not GUI(GuiEnum.GUI_OPTION).Visible Then Exit Sub
    
    If GuiVisibleCount <= 0 Then Exit Sub
    If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_OPTION Then Exit Sub
    
    If GuiPathEdit Then
        GuiPath = InputText(GuiPath, KeyAscii)
        setDidChange = True
    End If
End Sub

Private Sub OptionKeyUp(KeyCode As Integer, Shift As Integer)
    '//Make sure it's visible
    If Not GUI(GuiEnum.GUI_OPTION).Visible Then Exit Sub
    
    If GuiVisibleCount <= 0 Then Exit Sub
    If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_OPTION Then Exit Sub
    
    If editKey > 0 And editKey <= ControlEnum.Control_Count - 1 Then
        If Not InvalidInput(KeyCode) Then
            If Not CheckSameKey(KeyCode) Then
                TmpKey(editKey) = KeyCode
                '//Exit editing
                editKey = 0
                setDidChange = True
            Else
                AddAlert "Key Input already in used", White
            End If
        Else
            AddAlert "Invalid Key Input", White
        End If
    End If
End Sub

Private Sub OptionMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim tmpX As Long, tmpY As Long
Dim Count As Long

    With GUI(GuiEnum.GUI_OPTION)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_OPTION
        
        '//Loop through all items
        For i = ButtonEnum.Option_Close To ButtonEnum.Option_sSoundDown
            If setWindow <> i Then
                If CanShowButton(i) Then
                    If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateHover Then
                            Button(i).State = ButtonState.StateClick
                        End If
                    End If
                End If
            End If
        Next
        
        GuiPathEdit = False
        
        '//Window
        Select Case setWindow
            Case ButtonEnum.Option_Video
                '//Fullscreen
                tmpX = 105: tmpY = 45
                If CursorX >= .X + tmpX And CursorX <= .X + tmpX + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                    If isFullscreen = YES Then
                        isFullscreen = NO
                    Else
                        isFullscreen = YES
                    End If
                    setDidChange = True
                End If
            Case ButtonEnum.Option_Sound
                tmpX = 105: tmpY = 45
                For i = 1 To MAX_VOLUME
                    If CursorX >= .X + tmpX + 125 + ((8 + 3) * (i - 1)) And CursorX <= .X + tmpX + 125 + ((8 + 3) * (i - 1)) + 9 Then
                        '//Background Music
                        If CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 20 Then
                            BGVolume = i
                            setDidChange = True
                        ElseIf CursorY >= .Y + 27 + tmpY And CursorY <= .Y + 27 + tmpY + 20 Then
                        '//Sound Effect
                            SEVolume = i
                            setDidChange = True
                        End If
                    End If
                Next
            Case ButtonEnum.Option_Game
                '//Show Fps
                tmpX = 105: tmpY = 70
                If CursorX >= .X + tmpX And CursorX <= .X + tmpX + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                    If FPSvisible = YES Then
                        FPSvisible = NO
                    Else
                        FPSvisible = YES
                    End If
                    setDidChange = True
                End If
                
                '//Show Ping
                tmpX = 105: tmpY = 90
                If CursorX >= .X + tmpX And CursorX <= .X + tmpX + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                    If PingVisible = YES Then
                        PingVisible = NO
                    Else
                        PingVisible = YES
                    End If
                    setDidChange = True
                End If
                
                '//Skip Boot Up
                tmpX = 105: tmpY = 110
                If CursorX >= .X + tmpX And CursorX <= .X + tmpX + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                    If tSkipBootUp = YES Then
                        tSkipBootUp = NO
                    Else
                        tSkipBootUp = YES
                    End If
                    setDidChange = True
                End If
                
                '//Name Visible
                tmpX = 105: tmpY = 130
                If CursorX >= .X + tmpX And CursorX <= .X + tmpX + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                    If Namevisible = YES Then
                        Namevisible = NO
                    Else
                        Namevisible = YES
                    End If
                    setDidChange = True
                End If
            Case ButtonEnum.Option_Control
                '//Control Key
                For i = 1 To MAX_CONTROL_PREV
                    Count = CurControlKey + (i - 1)
                    If Count > 0 And Count <= ControlEnum.Control_Count - 1 Then
                        If CursorX >= .X + 212 And CursorX <= .X + 212 + 122 And CursorY >= .Y + 44 + ((25 + 5) * (i - 1)) And CursorY <= .Y + 44 + ((25 + 5) * (i - 1)) + 24 Then
                            editKey = Count
                        End If
                    End If
                Next
        End Select
        
        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Private Sub OptionMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim tmpX As Long, tmpY As Long
Dim Count As Long

    With GUI(GuiEnum.GUI_OPTION)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If GuiVisibleCount <= 0 Then Exit Sub

        IsHovering = False

        '//Loop through all items
        For i = ButtonEnum.Option_Close To ButtonEnum.Option_sSoundDown
            If setWindow <> i Then
                If CanShowButton(i) Then
                    If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateNormal Then
                            Button(i).State = ButtonState.StateHover
                
                            IsHovering = True
                            MouseIcon = 1 '//Select
                        End If
                    End If
                End If
            End If
        Next
        
        '//Window
        Select Case setWindow
            Case ButtonEnum.Option_Video
                '//Fullscreen
                tmpX = 105: tmpY = 45
                If CursorX >= .X + tmpX And CursorX <= .X + tmpX + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                    IsHovering = True
                    MouseIcon = 1 '//Select
                End If
            Case ButtonEnum.Option_Sound
                tmpX = 105: tmpY = 45
                For i = 1 To MAX_VOLUME
                    If CursorX >= .X + tmpX + 125 + ((8 + 3) * (i - 1)) And CursorX <= .X + tmpX + 125 + ((8 + 3) * (i - 1)) + 9 Then
                        '//Background Music
                        If CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 20 Then
                            IsHovering = True
                            MouseIcon = 1 '//Select
                        ElseIf CursorY >= .Y + 27 + tmpY And CursorY <= .Y + 27 + tmpY + 20 Then
                        '//Sound Effect
                            IsHovering = True
                            MouseIcon = 1 '//Select
                        End If
                    End If
                Next
            Case ButtonEnum.Option_Game
                '//Gui Path
                If CursorX >= .X + 165 And CursorX <= .X + 165 + 180 And CursorY >= .Y + 45 And CursorY <= .Y + 45 + 18 Then
                    IsHovering = True
                    MouseIcon = 2 '//I-Beam
                End If
                
                '//FPS
                tmpX = 105: tmpY = 70
                If CursorX >= .X + tmpX And CursorX <= .X + tmpX + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                    IsHovering = True
                    MouseIcon = 1 '//Select
                End If
                
                '//Ping
                tmpX = 105: tmpY = 90
                If CursorX >= .X + tmpX And CursorX <= .X + tmpX + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                    IsHovering = True
                    MouseIcon = 1 '//Select
                End If
                
                '//Skip BootUp
                tmpX = 105: tmpY = 110
                If CursorX >= .X + tmpX And CursorX <= .X + tmpX + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                    IsHovering = True
                    MouseIcon = 1 '//Select
                End If
                
                '//Show Name
                tmpX = 105: tmpY = 130
                If CursorX >= .X + tmpX And CursorX <= .X + tmpX + 17 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 17 Then
                    IsHovering = True
                    MouseIcon = 1 '//Select
                End If
                
                '//Language
                tmpX = 105: tmpY = 160
                For i = 1 To MAX_LANGUAGE
                    If CursorX >= .X + tmpX + 80 + ((i - 1) * 55) And CursorX <= .X + tmpX + 80 + ((i - 1) * 55) + 45 And CursorY >= .Y + tmpY And CursorY <= .Y + tmpY + 25 Then
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                Next
            Case ButtonEnum.Option_Control
                '//Control Key
                For i = 1 To MAX_CONTROL_PREV
                    Count = CurControlKey + (i - 1)
                    If Count > 0 And Count <= ControlEnum.Control_Count - 1 Then
                        If CursorX >= .X + 212 And CursorX <= .X + 212 + 122 And CursorY >= .Y + 44 + ((25 + 5) * (i - 1)) And CursorY <= .Y + 44 + ((25 + 5) * (i - 1)) + 24 Then
                            IsHovering = True
                            MouseIcon = 1 '//Select
                        End If
                    End If
                Next
        End Select
        
        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Private Sub OptionMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long, z As Long

    With GUI(GuiEnum.GUI_OPTION)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If GuiVisibleCount <= 0 Then Exit Sub

        '//Loop through all items
        For i = ButtonEnum.Option_Close To ButtonEnum.Option_sSoundDown
            If setWindow <> i Then
                If CanShowButton(i) Then
                    If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateClick Then
                            Button(i).State = ButtonState.StateNormal
                            Select Case i
                                Case ButtonEnum.Option_Close
                                    If setDidChange Then
                                        OpenChoiceBox "Do you want to save this settings?", CB_SAVESETTING
                                    Else
                                        GuiState GUI_OPTION, False
                                    End If
                                Case ButtonEnum.Option_Video, ButtonEnum.Option_Sound, ButtonEnum.Option_Game, ButtonEnum.Option_Control
                                    setWindow = i
                                Case ButtonEnum.Option_cTabUp
                                    If CurControlKey > 1 Then
                                        CurControlKey = CurControlKey - 1
                                    End If
                                Case ButtonEnum.Option_cTabDown
                                    If CurControlKey + (MAX_CONTROL_PREV - 1) < ControlEnum.Control_Count - 1 Then
                                        CurControlKey = CurControlKey + 1
                                    End If
                                Case ButtonEnum.Option_sMusicUp
                                    If BGVolume < MAX_VOLUME Then
                                        BGVolume = BGVolume + 1
                                        setDidChange = True
                                    End If
                                Case ButtonEnum.Option_sMusicDown
                                    If BGVolume > 0 Then
                                        BGVolume = BGVolume - 1
                                        setDidChange = True
                                    End If
                                Case ButtonEnum.Option_sSoundUp
                                    If SEVolume < MAX_VOLUME Then
                                        SEVolume = SEVolume + 1
                                        setDidChange = True
                                    End If
                                Case ButtonEnum.Option_sSoundDown
                                    If SEVolume > 0 Then
                                        SEVolume = SEVolume - 1
                                        setDidChange = True
                                    End If
                            End Select
                        End If
                    End If
                End If
            End If
        Next
        
        '//Window
        Select Case setWindow
            Case ButtonEnum.Option_Video

            Case ButtonEnum.Option_Sound
                
            Case ButtonEnum.Option_Game
                '//Gui Path
                If CursorX >= .X + 165 And CursorX <= .X + 165 + 180 And CursorY >= .Y + 45 And CursorY <= .Y + 45 + 18 Then
                    GuiPathEdit = True
                End If
                
                '//Language
                For i = 1 To MAX_LANGUAGE
                    If CursorX >= .X + 105 + 80 + ((i - 1) * 55) And CursorX <= .X + 105 + 80 + ((i - 1) * 55) + 45 And CursorY >= .Y + 160 And CursorY <= .Y + 160 + 25 Then
                        tmpCurLanguage = (i - 1)
                        setDidChange = True
                    End If
                Next
            Case ButtonEnum.Option_Control
            
        End Select
        
        '//Check for dragging
        .InDrag = False
    End With
End Sub

' *********************
' ** ChatBox **
' *********************
Private Sub ChatboxKeyPress(KeyAscii As Integer)
Dim i As Long
Dim cacheMsg As String

    '//Make sure it's visible
    If Not GUI(GuiEnum.GUI_CHATBOX).Visible Then Exit Sub
    
    If GuiVisibleCount <= 0 Then Exit Sub
    If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_CHATBOX Then Exit Sub

    If KeyAscii = vbKeyReturn Then
        If Not ChatOn Then
            ChatOn = True
            EditTab = False
            MyChat = vbNullString
        Else
            '//SendChat
            If Len(Trim$(MyChat)) > 0 Then
                HandleChatMsg MyChat
                ChatOn = False
                MyChat = vbNullString
            Else
                ChatOn = False
                MyChat = vbNullString
            End If
        End If
    ElseIf KeyAscii = vbKeyTab Then
        If Not EditTab Then
            EditTab = True
            ChatOn = False
            MyChat = vbNullString
        Else
            EditTab = False
            ChatOn = True
            MyChat = vbNullString
        End If
    End If
    
    If ChatOn Then
        If KeyAscii = vbKeySpace Then
            If Left$(MyChat, 1) = "@" Then
                MyChat = Mid$(MyChat, 2, Len(MyChat) - 1)
                ChatTab = vbNullString
        
                ' Get the desired player from the user text
                For i = 1 To Len(MyChat)
                    If Mid$(MyChat, i, 1) <> Space(1) Then
                        ChatTab = ChatTab & Mid$(MyChat, i, 1)
                    Else
                        Exit For
                    End If
                Next
            
                ' Make sure they are actually sending something
                If Len(MyChat) - i > 0 Then
                    MyChat = Mid$(MyChat, i + 1, Len(MyChat) - i)
                Else
                    MyChat = vbNullString
                End If
                
                Exit Sub
            End If
            
            If Left$(MyChat, 1) = "/" Then
                cacheMsg = LCase(MyChat)
                
                Select Case cacheMsg
                    Case "/map"
                        ChatTab = "/map"
                        MyChat = vbNullString
                        cacheMsg = vbNullString
                        Exit Sub
                    Case "/all"
                        ChatTab = "/all"
                        MyChat = vbNullString
                        cacheMsg = vbNullString
                        Exit Sub
                End Select
            End If
        End If
        
        If Len(MyChat) < MAX_CHAT_TEXT Or KeyAscii = vbKeyBack Then MyChat = InputText(MyChat, KeyAscii)
    ElseIf EditTab Then
        If Len(ChatTab) < (NAME_LENGTH - 1) Or KeyAscii = vbKeyBack Then ChatTab = InputText(ChatTab, KeyAscii)
    End If
End Sub

Private Sub ChatBoxMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_CHATBOX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_CHATBOX
        
        '//Loop through all items
        For i = ButtonEnum.Chatbox_Send To ButtonEnum.Chatbox_ScrollDown
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                    Select Case i
                        Case ButtonEnum.Chatbox_ScrollUp
                            ChatScrollUp = True
                            ChatScrollDown = False
                            ChatScrollTimer = GetTickCount
                        Case ButtonEnum.Chatbox_ScrollDown
                            ChatScrollUp = False
                            ChatScrollDown = True
                            ChatScrollTimer = GetTickCount
                    End Select
                End If
            End If
        Next
        
        If CursorX >= .X + 59 And CursorX <= .X + 59 + 314 And CursorY >= .Y + 144 And CursorY <= .Y + 144 + 19 Then
            ChatOn = True
            EditTab = False
        End If
        If CursorX >= .X + 6 And CursorX <= .X + 6 + 46 And CursorY >= .Y + 144 And CursorY <= .Y + 144 + 19 Then
            EditTab = True
            ChatOn = False
            MyChat = vbNullString
        End If
        
        '//Chat Scroll
        If totalChatLines > MaxChatLine Then
            ' Chat scroll
            If CursorX >= .X + chatScrollX And CursorX <= .X + chatScrollX + chatScrollW And CursorY >= .Y + chatScrollTop + (chatScrollL - chatScrollY) And CursorY <= .Y + chatScrollTop + (chatScrollL - chatScrollY) + chatScrollH Then
                ChatHold = True
            End If
        End If
        
        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 137 And .OldMouseX >= 19 Then
            .InDrag = True
        End If
    End With
End Sub

Private Sub ChatBoxMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim MaxY As Long
Dim tmpX As Long, tmpY As Long

    With GUI(GuiEnum.GUI_CHATBOX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_CHATBOX Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        For i = ButtonEnum.Chatbox_Send To ButtonEnum.Chatbox_ScrollDown
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
        
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next
        
        '//Hovering Textbox
        If CursorX >= .X + 59 And CursorX <= .X + 59 + 314 And CursorY >= .Y + 144 And CursorY <= .Y + 144 + 19 Then
            IsHovering = True
            MouseIcon = 2 '//I-Beam
        End If
        If CursorX >= .X + 6 And CursorX <= .X + 6 + 46 And CursorY >= .Y + 144 And CursorY <= .Y + 144 + 19 Then
            IsHovering = True
            MouseIcon = 2 '//I-Beam
        End If
        
        '//Hovering Chatscroll
        If CursorX >= .X + chatScrollX And CursorX <= .X + chatScrollX + chatScrollW And CursorY >= .Y + chatScrollTop + (chatScrollL - chatScrollY) And CursorY <= .Y + chatScrollTop + (chatScrollL - chatScrollY) + chatScrollH Then
            If totalChatLines > MaxChatLine Then
                IsHovering = True
                MouseIcon = 1 '//Select
            End If
        End If
  
        '//Chat Scroll
        If totalChatLines > MaxChatLine Then
            '//Scroll moving
            If ChatHold Then
                '//Upward
                If CursorY < .Y + chatScrollTop + (chatScrollL - chatScrollY) + (chatScrollH / 2) Then
                    If chatScrollY < chatScrollL Then
                        chatScrollY = (CursorY - (.Y + chatScrollTop + chatScrollL) - (chatScrollH / 2)) * -1
                        If chatScrollY >= chatScrollL Then chatScrollY = chatScrollL
                    End If
                End If
                '//Downward
                If CursorY > .Y + chatScrollTop + (chatScrollL - chatScrollY) + chatScrollH - (chatScrollH / 2) Then
                    If chatScrollY > 0 Then
                        chatScrollY = (CursorY - (.Y + chatScrollTop + chatScrollL) - chatScrollH + (chatScrollH / 2)) * -1
                        If chatScrollY <= 0 Then chatScrollY = 0
                    End If
                End If
                
                MaxY = totalChatLines
                If MaxY < MaxChatLine Then MaxY = MaxChatLine
                
                ChatScroll = (chatScrollY / (chatScrollL / (MaxY - 7))) + 7
                If ChatScroll < MaxChatLine Then ChatScroll = MaxChatLine
                If ChatScroll > MaxY Then ChatScroll = MaxY
                '//update the array
                UpdateChatArray
            End If
        End If
        
        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
            
            UpdateChatArray
        End If
    End With
End Sub

Private Sub ChatBoxMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim FoundError As Boolean

    With GUI(GuiEnum.GUI_CHATBOX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_CHATBOX Then Exit Sub
        
        '//Loop through all items
        For i = ButtonEnum.Chatbox_Send To ButtonEnum.Chatbox_ScrollDown
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        '//SendChat
                        If Len(Trim$(MyChat)) > 0 Then
                            HandleChatMsg MyChat
                        End If
                    End If
                End If
            End If
        Next
        
        '//Chat Scroll
        ChatHold = False
        
        '//Check for dragging
        .InDrag = False
    End With
End Sub

' ***************
' ** Inventory **
' ***************
Private Sub InventoryMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_INVENTORY)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_INVENTORY
        
        '//Loop through all items
        For i = ButtonEnum.Inventory_Close To ButtonEnum.Inventory_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
        
        If Not SelMenu.Visible And InvUseSlot = 0 Then
            If Buttons = vbRightButton Then
                '//Inv
                i = IsInvItem(CursorX, CursorY)
                If i > 0 Then
                    OpenSelMenu SelMenuType.Inv, i
                End If
            Else
                '//Disable Drag when intrade
                If TradeIndex = 0 Then
                    '//Inv
                    i = IsInvItem(CursorX, CursorY)
                    If i > 0 Then
                        DragInvSlot = i
                        WindowPriority = GuiEnum.GUI_INVENTORY
                    End If
                End If
            End If
        End If
        
        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Private Sub InventoryMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmpX As Long, tmpY As Long
Dim i As Long

    With GUI(GuiEnum.GUI_INVENTORY)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If CursorX >= .X And CursorX <= .X + .Width And CursorY >= .Y And CursorY <= .Y + .Height Then
        Else
            Exit Sub
        End If
        
        If DragInvSlot > 0 Or DragStorageSlot > 0 Then
            If WindowPriority = 0 Then
                WindowPriority = GuiEnum.GUI_INVENTORY
                If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_INVENTORY Then
                    UpdateGuiOrder GUI_INVENTORY
                End If
            End If
        End If
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_INVENTORY Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        For i = ButtonEnum.Inventory_Close To ButtonEnum.Inventory_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
        
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next
        
        '//Inv
        i = IsInvItem(CursorX, CursorY)
        If i > 0 Then
            IsHovering = True
            MouseIcon = 1 '//Select
            
            If Not InvItemDesc = i Then
                InvItemDesc = i
                InvItemDescTimer = GetTickCount
                InvItemDescShow = False
            End If
        End If

        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Private Sub InventoryMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_INVENTORY)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_INVENTORY Then Exit Sub
        
        '//Loop through all items
        For i = ButtonEnum.Inventory_Close To ButtonEnum.Inventory_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        Select Case i
                            Case ButtonEnum.Inventory_Close
                                If GUI(GuiEnum.GUI_INVENTORY).Visible Then
                                    GuiState GUI_INVENTORY, False
                                End If
                        End Select
                    End If
                End If
            End If
        Next
        
        '//Replace item
        If TradeIndex = 0 Then
            If DragInvSlot > 0 Then
                i = IsInvSlot(CursorX, CursorY)
                If i > 0 Then
                    SendSwitchInvSlot DragInvSlot, i
                End If
            End If
            DragInvSlot = 0
        End If
        
        '//Replace item
        If DragStorageSlot > 0 Then
            i = IsInvSlot(CursorX, CursorY)
            If i > 0 Then
                '//Check if value is greater than 1
                If PlayerInvStorage(InvCurSlot).Data(DragStorageSlot).value > 1 Then
                    If Not GUI(GuiEnum.GUI_CHOICEBOX).Visible Then
                        OpenInputBox "Enter amount", IB_WITHDRAW, DragStorageSlot, i
                    End If
                Else
                    '//Send Withdraw
                    SendWithdrawItemTo InvCurSlot, DragStorageSlot, i
                End If
            End If
        End If
        DragStorageSlot = 0
        
        '//Check for dragging
        .InDrag = False
    End With
End Sub

'////////////////////////////
'/////// Search Input ///////
'////////////////////////////
Private Sub SearchMouseDown(Buttons As Integer)
Dim i As Long
Dim InUsed As Byte
Dim Data1 As Long, Data2 As Long, Data3 As Long

    '//Usage of item
    If InvUseDataType > 0 Then
        InUsed = NO
        Data1 = 0: Data2 = 0: Data3 = 0
        Select Case InvUseDataType
            Case ItemTypeEnum.Pokeball
                '//Check if there's pokemon on tile
                For i = 1 To Pokemon_HighIndex
                    If MapPokemon(i).Num > 0 Then
                        If MapPokemon(i).Map = Player(MyIndex).Map Then
                            If curTileX = MapPokemon(i).X And curTileY = MapPokemon(i).Y Then
                                If curTileX >= Player(MyIndex).X - 4 And curTileX <= Player(MyIndex).X + 4 Then
                                    If curTileY >= Player(MyIndex).Y - 4 And curTileY <= Player(MyIndex).Y + 4 Then
                                        '//Catch Poke
                                        InUsed = YES
                                        Data1 = i
                                        Exit For
                                    Else
                                        AddAlert "Not in range", White
                                    End If
                                Else
                                    AddAlert "Not in range", White
                                End If
                            End If
                        End If
                    End If
                Next
                For i = 1 To Player_HighIndex
                    If PlayerPokemon(i).Num > 0 Then
                        If Player(i).Map = Player(MyIndex).Map Then
                            If curTileX = PlayerPokemon(i).X And curTileY = PlayerPokemon(i).Y Then
                                If curTileX >= Player(MyIndex).X - 4 And curTileX <= Player(MyIndex).X + 4 Then
                                    If curTileY >= Player(MyIndex).Y - 4 And curTileY <= Player(MyIndex).Y + 4 Then
                                        AddAlert "You cannot catch this Pokemon", White
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            Case ItemTypeEnum.Medicine
                If Item(PlayerInv(InvUseSlot).Num).Data1 = 4 Then '//Revive
                    For i = 1 To MAX_PLAYER_POKEMON
                        If PlayerPokemons(i).Num > 0 Then
                            curTileX = Screen_Width - 34 - 5 - ((i - 1) * 40)
                            curTileY = 62 ' + 52 + ((i - 1) * 40)
                                            
                            If CursorX >= curTileX And CursorX <= curTileX + 34 And CursorY >= curTileY And CursorY <= curTileY + 37 Then
                                '//Catch Poke
                                InUsed = YES
                                Data1 = i
                                Exit For
                            End If
                        End If
                    Next
                End If
        End Select
        
        SendGotData InUsed, Data1
        InvUseDataType = 0
        InvUseSlot = 0
        Exit Sub
    End If
    
    For i = 1 To Pokemon_HighIndex
        If MapPokemon(i).Num > 0 Then
            If MapPokemon(i).Map = Player(MyIndex).Map Then
                If curTileX = MapPokemon(i).X And curTileY = MapPokemon(i).Y Then
                    If curTileX >= Player(MyIndex).X - 4 And curTileX <= Player(MyIndex).X + 4 Then
                        If curTileY >= Player(MyIndex).Y - 4 And curTileY <= Player(MyIndex).Y + 4 Then
                            '//Scan Pokedex
                            OpenSelMenu SelMenuType.PokedexMapPokemon, i
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    For i = 1 To Npc_HighIndex
        If MapNpc(i).Num > 0 Then
            If curTileX = MapNpc(i).X And curTileY = MapNpc(i).Y Then
                '//Make sure in range
                If curTileX >= Player(MyIndex).X - 4 And curTileX <= Player(MyIndex).X + 4 Then
                    If curTileY >= Player(MyIndex).Y - 4 And curTileY <= Player(MyIndex).Y + 4 Then
                        If Npc(MapNpc(i).Num).Convo > 0 Then
                            OpenSelMenu SelMenuType.NPCChat, i
                        End If
                    Else
                        AddAlert "Not in range", White
                    End If
                Else
                    AddAlert "Not in range", White
                End If
            End If
        End If
    Next
    '//Player
    If Buttons = vbRightButton Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If Player(i).Map = Player(MyIndex).Map Then
                    If PlayerPokemon(i).Num > 0 Then
                        If curTileX = PlayerPokemon(i).X And curTileY = PlayerPokemon(i).Y Then
                            '//Make sure in range
                            If curTileX >= Player(MyIndex).X - 4 And curTileX <= Player(MyIndex).X + 4 Then
                                If curTileY >= Player(MyIndex).Y - 4 And curTileY <= Player(MyIndex).Y + 4 Then
                                    OpenSelMenu SelMenuType.PokedexPlayerPokemon, i
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                    If curTileX = Player(i).X And curTileY = Player(i).Y Then
                        '//Make sure in range
                        If curTileX >= Player(MyIndex).X - 4 And curTileX <= Player(MyIndex).X + 4 Then
                            If curTileY >= Player(MyIndex).Y - 4 And curTileY <= Player(MyIndex).Y + 4 Then
                                OpenSelMenu SelMenuType.PlayerMenu, i
                                Exit For
                            Else
                                AddAlert "Not in range", White
                            End If
                        Else
                            AddAlert "Not in range", White
                        End If
                    End If
                End If
            End If
        Next
    End If
    
    If curTileX >= Player(MyIndex).X - 1 And curTileX <= Player(MyIndex).X + 1 Then
        If curTileY >= Player(MyIndex).Y - 1 And curTileY <= Player(MyIndex).Y + 1 Then
            If Editor = 0 Then
                If Map.Tile(curTileX, curTileY).Attribute = MapAttribute.BothStorage Then
                    If Not GUI(GuiEnum.GUI_INVSTORAGE).Visible And Not GUI(GuiEnum.GUI_POKEMONSTORAGE).Visible Then
                        OpenSelMenu SelMenuType.Storage
                    End If
                ElseIf Map.Tile(curTileX, curTileY).Attribute = MapAttribute.InvStorage Then
                    If Not GUI(GuiEnum.GUI_INVSTORAGE).Visible And Not GUI(GuiEnum.GUI_POKEMONSTORAGE).Visible Then
                        SendOpenStorage 1
                    End If
                ElseIf Map.Tile(curTileX, curTileY).Attribute = MapAttribute.PokemonStorage Then
                    If Not GUI(GuiEnum.GUI_INVSTORAGE).Visible And Not GUI(GuiEnum.GUI_POKEMONSTORAGE).Visible Then
                        SendOpenStorage 2
                    End If
                ElseIf Map.Tile(curTileX, curTileY).Attribute = MapAttribute.ConvoTile Then
                    If ConvoNum = 0 Then
                        OpenSelMenu SelMenuType.ConvoTileCheck, Map.Tile(curTileX, curTileY).Data1
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub SearchMouseMove(Buttons As Integer)
Dim i As Long

    IsHovering = False
    
    If MyIndex = 0 Then Exit Sub
    If GettingMap Then Exit Sub

    '//Player
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Player(i).Map = Player(MyIndex).Map Then
                If curTileX = Player(i).X And curTileY = Player(i).Y Then
                    IsHovering = True
                    MouseIcon = 1 '//Select
                End If
                If i <> MyIndex Then
                    '//Player Pokemon
                    If PlayerPokemon(i).Num > 0 Then
                        If curTileX = PlayerPokemon(i).X And curTileY = PlayerPokemon(i).Y Then
                            IsHovering = True
                            MouseIcon = 1 '//Select
                        End If
                    End If
                End If
            End If
        End If
    Next
    For i = 1 To Npc_HighIndex
        If MapNpc(i).Num > 0 Then
            If curTileX = MapNpc(i).X And curTileY = MapNpc(i).Y Then
                IsHovering = True
                MouseIcon = 1 '//Select
            End If
        End If
    Next
    For i = 1 To Pokemon_HighIndex
        If MapPokemon(i).Num > 0 Then
            If MapPokemon(i).Map = Player(MyIndex).Map Then
                If curTileX = MapPokemon(i).X And curTileY = MapPokemon(i).Y Then
                    IsHovering = True
                    MouseIcon = 1 '//Select
                End If
            End If
        End If
    Next
        
    If curTileX >= Player(MyIndex).X - 1 And curTileX <= Player(MyIndex).X + 1 Then
        If curTileY >= Player(MyIndex).Y - 1 And curTileY <= Player(MyIndex).Y + 1 Then
            If Editor = 0 Then
                If Map.Tile(curTileX, curTileY).Attribute = MapAttribute.BothStorage Or Map.Tile(curTileX, curTileY).Attribute = MapAttribute.InvStorage Or Map.Tile(curTileX, curTileY).Attribute = MapAttribute.PokemonStorage Then
                    If Not GUI(GuiEnum.GUI_INVSTORAGE).Visible And Not GUI(GuiEnum.GUI_POKEMONSTORAGE).Visible Then
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                ElseIf Map.Tile(curTileX, curTileY).Attribute = MapAttribute.ConvoTile Then
                    IsHovering = True
                    MouseIcon = 1 '//Select
                End If
            End If
        End If
    End If
End Sub

' **************
' ** InputBox **
' **************
Private Sub InputBoxKeyPress(KeyAscii As Integer)
    '//Make sure it's visible
    If Not GUI(GuiEnum.GUI_INPUTBOX).Visible Then Exit Sub
    
    Select Case InputBoxType
        Case IB_NEWPASSWORD, IB_PASSWORDCONFIRM, IB_OLDPASSWORD
            If (isNameLegal(KeyAscii, True) And Len(InputBoxText) < (InputBoxLen - 1)) Or KeyAscii = vbKeyBack Then InputBoxText = InputText(InputBoxText, KeyAscii)
        Case IB_WITHDRAW
            If IsNumeric(KeyAscii) Then
                InputBoxText = InputText(InputBoxText, KeyAscii)
                If Val(InputBoxText) > PlayerInvStorage(InvCurSlot).Data(InputBoxData1).value Then
                    InputBoxText = PlayerInvStorage(InvCurSlot).Data(InputBoxData1).value
                End If
            End If
        Case IB_DEPOSIT
            If IsNumeric(KeyAscii) Then
                InputBoxText = InputText(InputBoxText, KeyAscii)
                If Val(InputBoxText) > PlayerInv(InputBoxData1).value Then
                    InputBoxText = PlayerInv(InputBoxData1).value
                End If
            End If
        Case IB_BUYITEM
            If IsNumeric(KeyAscii) Then
                InputBoxText = InputText(InputBoxText, KeyAscii)
                If (Shop(ShopNum).ShopItem(InputBoxData1).Price * Val(InputBoxText)) > Player(MyIndex).Money Then
                    InputBoxText = Round(Player(MyIndex).Money / Shop(ShopNum).ShopItem(InputBoxData1).Price, 0)
                End If
            End If
        Case IB_SELLITEM, IB_ADDTRADE
            If IsNumeric(KeyAscii) Then
                InputBoxText = InputText(InputBoxText, KeyAscii)
                If Val(InputBoxText) > PlayerInv(InputBoxData1).value Then
                    InputBoxText = PlayerInv(InputBoxData1).value
                End If
            End If
    End Select
End Sub

Private Sub InputBoxMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_INPUTBOX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_INPUTBOX
        
        '//Loop through all items
        For i = ButtonEnum.InputBox_Okay To ButtonEnum.InputBox_Cancel
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Sub InputBoxMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_INPUTBOX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If GuiVisibleCount <= 0 Then Exit Sub

        IsHovering = False

        '//Loop through all items
        For i = ButtonEnum.InputBox_Okay To ButtonEnum.InputBox_Cancel
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
        
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next
        
        '//Textbox
        If CursorX >= .X + 22 And CursorX <= .X + 22 + 223 And CursorY >= .Y + 34 And CursorY <= .Y + 34 + 19 Then
            IsHovering = True
            MouseIcon = 2 '//I-Beam
        End If
    End With
End Sub

Private Sub InputBoxMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_INPUTBOX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If GuiVisibleCount <= 0 Then Exit Sub

        '//Loop through all items
        For i = ButtonEnum.InputBox_Okay To ButtonEnum.InputBox_Cancel
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        '//Do function of the button
                        Select Case i
                            Case ButtonEnum.InputBox_Okay
                                Select Case InputBoxType
                                    Case IB_NEWPASSWORD
                                        NewPassword = InputBoxText
                                        InputBoxHeader = "Confirm your password"
                                        InputBoxType = IB_PASSWORDCONFIRM
                                        InputBoxText = vbNullString
                                    Case IB_PASSWORDCONFIRM
                                        If NewPassword = InputBoxText Then
                                            InputBoxHeader = "Enter your old password"
                                            InputBoxType = IB_OLDPASSWORD
                                            InputBoxText = vbNullString
                                        Else
                                            AddAlert "Password doesn't match", White
                                        End If
                                    Case IB_OLDPASSWORD
                                        OldPassword = InputBoxText
                                        '//Send Change Pass Data
                                        SendChangePassword NewPassword, OldPassword
                                        CloseInputBox
                                    Case IB_WITHDRAW
                                        If IsNumeric(InputBoxText) Then
                                            SendWithdrawItemTo InvCurSlot, InputBoxData1, InputBoxData2, Val(InputBoxText)
                                        End If
                                        CloseInputBox
                                    Case IB_DEPOSIT
                                        If IsNumeric(InputBoxText) Then
                                            SendDepositItemTo InvCurSlot, InputBoxData2, InputBoxData1, Val(InputBoxText)
                                        End If
                                        CloseInputBox
                                    Case IB_BUYITEM
                                        If IsNumeric(InputBoxText) Then
                                            '//Send Buy Item
                                            SendBuyItem InputBoxData1, Val(InputBoxText)
                                        End If
                                        CloseInputBox
                                    Case IB_SELLITEM
                                        If IsNumeric(InputBoxText) Then
                                            '//Send Sell Item
                                            SendSellItem InputBoxData1, Val(InputBoxText)
                                        End If
                                        CloseInputBox
                                    Case IB_ADDTRADE
                                        If IsNumeric(InputBoxText) Then
                                            '//Send Add Trade Item
                                            SendAddTrade 1, InputBoxData1, Val(InputBoxText)
                                        End If
                                        CloseInputBox
                                End Select
                            Case ButtonEnum.InputBox_Cancel
                                Select Case InputBoxType
                                    Case Else
                                        CloseInputBox
                                End Select
                        End Select
                    End If
                End If
            End If
        Next
    End With
End Sub

Public Sub CloseInputBox()
    GuiState GUI_INPUTBOX, False
    InputBoxHeader = vbNullString
    InputBoxText = vbNullString
    InputBoxType = 0
    
    '//Password
    NewPassword = vbNullString
    OldPassword = vbNullString
End Sub

' ***************
' ** MoveReplace **
' ***************
Private Sub MoveReplaceMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_MOVEREPLACE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_MOVEREPLACE
        
        '//Loop through all items
        For i = ButtonEnum.MoveReplace_Slot1 To ButtonEnum.MoveReplace_Cancel
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next

        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Private Sub MoveReplaceMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmpX As Long, tmpY As Long
Dim i As Long

    With GUI(GuiEnum.GUI_MOVEREPLACE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_MOVEREPLACE Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        For i = ButtonEnum.MoveReplace_Slot1 To ButtonEnum.MoveReplace_Cancel
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
        
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next
        
        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Private Sub MoveReplaceMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim MoveSlot As Byte

    With GUI(GuiEnum.GUI_MOVEREPLACE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_MOVEREPLACE Then Exit Sub
        
        '//Loop through all items
        For i = ButtonEnum.MoveReplace_Slot1 To ButtonEnum.MoveReplace_Cancel
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        Select Case i
                            Case ButtonEnum.MoveReplace_Slot1 To ButtonEnum.MoveReplace_Slot4
                                MoveSlot = i - (ButtonEnum.MoveReplace_Slot1 - 1)
                                SendReplaceNewMove MoveSlot
                                GuiState GUI_MOVEREPLACE, False
                            Case ButtonEnum.MoveReplace_Cancel
                                SendReplaceNewMove 0
                                GuiState GUI_MOVEREPLACE, False
                        End Select
                    End If
                End If
            End If
        Next
        
        '//Check for dragging
        .InDrag = False
    End With
End Sub

' ***************
' ** Trainer **
' ***************
Private Sub TrainerMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_TRAINER)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_TRAINER
        
        '//Loop through all items
        For i = ButtonEnum.Trainer_Close To ButtonEnum.Trainer_Badge
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next

        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Private Sub TrainerMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmpX As Long, tmpY As Long
Dim i As Long

    With GUI(GuiEnum.GUI_TRAINER)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_TRAINER Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        For i = ButtonEnum.Trainer_Close To ButtonEnum.Trainer_Badge
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
        
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next
        
        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Private Sub TrainerMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_TRAINER)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_TRAINER Then Exit Sub
        
        '//Loop through all items
        For i = ButtonEnum.Trainer_Close To ButtonEnum.Trainer_Badge
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        Select Case i
                            Case ButtonEnum.Trainer_Close
                                If GUI(GuiEnum.GUI_TRAINER).Visible = True Then
                                    GuiState GUI_TRAINER, False
                                End If
                            Case ButtonEnum.Trainer_Badge
                                If GUI(GuiEnum.GUI_BADGE).Visible = False Then
                                    GuiState GUI_BADGE, True
                                End If
                        End Select
                    End If
                End If
            End If
        Next
        
        '//Check for dragging
        .InDrag = False
    End With
End Sub

' ***************
' ** InvStorage **
' ***************
Private Sub InvStorageMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_INVSTORAGE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_INVSTORAGE
        
        '//Loop through all items
        For i = ButtonEnum.InvStorage_Close To ButtonEnum.InvStorage_Slot5
            If i <> InvCurSlot Then
                If CanShowButton(i) Then
                    If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateHover Then
                            Button(i).State = ButtonState.StateClick
                        End If
                    End If
                End If
            End If
        Next
        
        If Not SelMenu.Visible Then
            If Buttons = vbRightButton Then
                '//Inv
                i = IsInvStorageItem(CursorX, CursorY)
                If i > 0 Then
                    OpenSelMenu SelMenuType.InvStorage, i
                End If
            Else
                '//Inv
                i = IsInvStorageItem(CursorX, CursorY)
                If i > 0 Then
                    DragStorageSlot = i
                    WindowPriority = GuiEnum.GUI_INVSTORAGE
                End If
            End If
        End If
        
        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Private Sub InvStorageMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmpX As Long, tmpY As Long
Dim i As Long

    With GUI(GuiEnum.GUI_INVSTORAGE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        If CursorX >= .X And CursorX <= .X + .Width And CursorY >= .Y And CursorY <= .Y + .Height Then
        Else
            Exit Sub
        End If
        
        If DragInvSlot > 0 Or DragStorageSlot > 0 Then
            If WindowPriority = 0 Then
                WindowPriority = GuiEnum.GUI_INVSTORAGE
                If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_INVSTORAGE Then
                    UpdateGuiOrder GUI_INVSTORAGE
                End If
            End If
        End If
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_INVSTORAGE Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        For i = ButtonEnum.InvStorage_Close To ButtonEnum.InvStorage_Slot5
            If i <> InvCurSlot Then
                If CanShowButton(i) Then
                    If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateNormal Then
                            Button(i).State = ButtonState.StateHover
            
                            IsHovering = True
                            MouseIcon = 1 '//Select
                        End If
                    End If
                End If
            End If
        Next
        
        '//Inv
        i = IsInvStorageItem(CursorX, CursorY)
        If i > 0 Then
            IsHovering = True
            MouseIcon = 1 '//Select
        End If
        
        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Private Sub InvStorageMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim SlotNum As Long
Dim Amount As Long

    With GUI(GuiEnum.GUI_INVSTORAGE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_INVSTORAGE Then Exit Sub
        
        '//Loop through all items
        For i = ButtonEnum.InvStorage_Close To ButtonEnum.InvStorage_Slot5
            If i <> InvCurSlot Then
                If CanShowButton(i) Then
                    If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateClick Then
                            Button(i).State = ButtonState.StateNormal
                            Select Case i
                                Case ButtonEnum.InvStorage_Slot1 To ButtonEnum.InvStorage_Slot5
                                    SlotNum = ((i + 1) - ButtonEnum.InvStorage_Slot1)
                                    If PlayerInvStorage(SlotNum).Unlocked = YES Then
                                        '//Switch Slot
                                        InvCurSlot = SlotNum
                                    Else
                                        Amount = 100000 * (SlotNum - 2)
                                        '//Buy Slot
                                        BuySlotType = 1 '//Item
                                        BuySlotData = SlotNum
                                        OpenChoiceBox "Do you want to buy this slot for $" & Amount & "?", CB_BUYSLOT
                                    End If
                                Case ButtonEnum.InvStorage_Close
                                    SendOpenStorage 0
                            End Select
                        End If
                    End If
                End If
            End If
        Next
        
        '//Replace item
        If DragStorageSlot > 0 Then
            i = IsInvStorageSlot(CursorX, CursorY)
            If i > 0 Then
                SendSwitchStorageSlot InvCurSlot, DragStorageSlot, i
            End If
        End If
        DragStorageSlot = 0
        
        '//Replace item
        If DragInvSlot > 0 Then
            i = IsInvStorageSlot(CursorX, CursorY)
            If i > 0 Then
                '//Check if value is greater than 1
                If PlayerInv(DragInvSlot).value > 1 Then
                    If Not GUI(GuiEnum.GUI_CHOICEBOX).Visible Then
                        OpenInputBox "Enter amount", IB_DEPOSIT, DragInvSlot, i
                    End If
                Else
                    '//Deposit
                    SendDepositItemTo InvCurSlot, i, DragInvSlot
                End If
            End If
        End If
        DragInvSlot = 0
        
        '//Check for dragging
        .InDrag = False
    End With
End Sub

' ***************
' ** PokemonStorage **
' ***************
Private Sub PokemonStorageMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim SlotNum As Long

    With GUI(GuiEnum.GUI_POKEMONSTORAGE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_POKEMONSTORAGE
        
        '//Loop through all items
        For i = ButtonEnum.PokemonStorage_Close To ButtonEnum.PokemonStorage_Slot5
            SlotNum = ((i + 1) - ButtonEnum.PokemonStorage_Slot1)
            If SlotNum <> PokemonCurSlot Then
                If CanShowButton(i) Then
                    If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateHover Then
                            Button(i).State = ButtonState.StateClick
                        End If
                    End If
                End If
            End If
        Next
        
        If Not SelMenu.Visible Then
            If Buttons = vbRightButton Then
                '//Inv
                i = IsPokeStorage(CursorX, CursorY)
                If i > 0 Then
                    OpenSelMenu SelMenuType.PokeStorage, i
                End If
            Else
                i = IsPokeStorage(CursorX, CursorY)
                If i > 0 Then
                    DragPokeSlot = i
                End If
            End If
        End If
        
        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Private Sub PokemonStorageMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmpX As Long, tmpY As Long
Dim i As Long
Dim SlotNum As Long

    With GUI(GuiEnum.GUI_POKEMONSTORAGE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_POKEMONSTORAGE Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        For i = ButtonEnum.PokemonStorage_Close To ButtonEnum.PokemonStorage_Slot5
            SlotNum = ((i + 1) - ButtonEnum.PokemonStorage_Slot1)
            If SlotNum <> PokemonCurSlot Then
                If CanShowButton(i) Then
                    If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateNormal Then
                            Button(i).State = ButtonState.StateHover
            
                            IsHovering = True
                            MouseIcon = 1 '//Select
                        End If
                    End If
                End If
            End If
        Next

        i = IsPokeStorage(CursorX, CursorY)
        If i > 0 Then
            IsHovering = True
            MouseIcon = 1 '//Select
        End If
        
        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Private Sub PokemonStorageMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim SlotNum As Long
Dim Amount As Long

    With GUI(GuiEnum.GUI_POKEMONSTORAGE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_POKEMONSTORAGE Then Exit Sub
        
        '//Loop through all items
        For i = ButtonEnum.PokemonStorage_Close To ButtonEnum.PokemonStorage_Slot5
            SlotNum = ((i + 1) - ButtonEnum.PokemonStorage_Slot1)
            If SlotNum <> PokemonCurSlot Then
                If CanShowButton(i) Then
                    If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateClick Then
                            Button(i).State = ButtonState.StateNormal
                            Select Case i
                                Case ButtonEnum.PokemonStorage_Slot1 To ButtonEnum.PokemonStorage_Slot5
                                    If PlayerPokemonStorage(SlotNum).Unlocked = YES Then
                                        '//Switch Slot
                                        PokemonCurSlot = SlotNum
                                    Else
                                        Amount = 100000 * (SlotNum - 2)
                                        '//Buy Slot
                                        BuySlotType = 2 '//Pokemon
                                        BuySlotData = SlotNum
                                        OpenChoiceBox "Do you want to buy this slot for $" & Amount & "?", CB_BUYSLOT
                                    End If
                                Case ButtonEnum.PokemonStorage_Close
                                    SendOpenStorage 0
                            End Select
                        End If
                    End If
                End If
            End If
        Next
        
        '//Replace item
        If DragPokeSlot > 0 Then
            i = IsPokeStorageSlot(CursorX, CursorY)
            If i > 0 Then
                SendSwitchStoragePokeSlot PokemonCurSlot, DragPokeSlot, i
            End If
        End If
        DragPokeSlot = 0
        
        '//Check for dragging
        .InDrag = False
    End With
End Sub

' ***************
' ** Convo **
' ***************
Private Sub ConvoMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_CONVO)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Loop through all items
        If ConvoShowButton Then
            For i = ButtonEnum.Convo_Reply1 To ButtonEnum.Convo_Reply3
                If CanShowButton(i) Then
                    If CursorX >= Button(i).X And CursorX <= Button(i).X + Button(i).Width And CursorY >= Button(i).Y And CursorY <= Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateHover Then
                            Button(i).State = ButtonState.StateClick
                        End If
                    End If
                End If
            Next
        End If
        
        '//Skip Scrolling Text
        If Not ConvoShowButton Then
            If ConvoNum > 0 Then
                If CursorX >= .X And CursorX <= .X + .Width And CursorY >= .Y And CursorY <= .Y + .Height Then
                    If Len(ConvoText) > ConvoDrawTextLen Then
                        ConvoDrawTextLen = Len(ConvoText)
                        ConvoRenderText = Left$(ConvoText, ConvoDrawTextLen)
                    Else
                        '//Proceed to next convo
                        If ConvoNoReply = YES Then
                            '//Proceed to next
                            SendProcessConvo
                        Else
                            '//Show Choice
                            ConvoShowButton = True
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub ConvoMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_CONVO)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        If ConvoShowButton Then
            For i = ButtonEnum.Convo_Reply1 To ButtonEnum.Convo_Reply3
                If CanShowButton(i) Then
                    If CursorX >= Button(i).X And CursorX <= Button(i).X + Button(i).Width And CursorY >= Button(i).Y And CursorY <= Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateNormal Then
                            Button(i).State = ButtonState.StateHover
                
                            IsHovering = True
                            MouseIcon = 1 '//Select
                        End If
                    End If
                End If
            Next
        End If
        
        If Not ConvoShowButton Then
            If ConvoNum > 0 Then
                If CursorX >= .X And CursorX <= .X + .Width And CursorY >= .Y And CursorY <= .Y + .Height Then
                    IsHovering = True
                    MouseIcon = 1 '//Select
                End If
            End If
        End If
    End With
End Sub

Private Sub ConvoMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_CONVO)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Loop through all items
        If ConvoShowButton Then
            For i = ButtonEnum.Convo_Reply1 To ButtonEnum.Convo_Reply3
                If CanShowButton(i) Then
                    If CursorX >= Button(i).X And CursorX <= Button(i).X + Button(i).Width And CursorY >= Button(i).Y And CursorY <= Button(i).Y + Button(i).Height Then
                        If Button(i).State = ButtonState.StateClick Then
                            Button(i).State = ButtonState.StateNormal
                            Select Case i
                                Case ButtonEnum.Convo_Reply1 To ButtonEnum.Convo_Reply3
                                    SendProcessConvo ((i + 1) - ButtonEnum.Convo_Reply1)
                            End Select
                        End If
                    End If
                End If
            Next
        End If
    End With
End Sub

' ***************
' ** Shop **
' ***************
Private Sub ShopMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim DrawX As Long, DrawY As Long

    With GUI(GuiEnum.GUI_SHOP)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_SHOP
        
        '//Loop through all items
        For i = ButtonEnum.Shop_Close To ButtonEnum.Shop_ScrollDown
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
        
        '//Item
        For i = ShopAddY To ShopAddY + 8
            If i > 0 And i <= MAX_SHOP_ITEM Then
                If Shop(ShopNum).ShopItem(i).Num > 0 Then
                    DrawX = .X + (31 + ((4 + 127) * (((((i + 1) - ShopAddY) - 1) Mod 3))))
                    DrawY = .Y + (42 + ((4 + 78) * ((((i + 1) - ShopAddY) - 1) \ 3)))
                    
                    '//Button
                    If CursorX >= DrawX + 12 And CursorX <= DrawX + 12 + 103 And CursorY >= DrawY + 44 And CursorY <= DrawY + 44 + 25 Then
                        ShopButtonHover = i
                        If ShopButtonState = 1 Then
                            ShopButtonState = 2 '//Click
                        End If
                        '//Buy Item
                        If Item(Shop(ShopNum).ShopItem(i).Num).Stock = YES Then
                            '//Add Input
                            OpenInputBox "Enter amount", IB_BUYITEM, i
                        Else
                            '//Buy Item
                            SendBuyItem i
                        End If
                    End If
                    
                    '//Icon
                    If CursorX >= DrawX + 9 And CursorX <= DrawX + 9 + 32 And CursorY >= DrawY + 6 And CursorY <= DrawY + 6 + 32 Then
                        
                    End If
                End If
            End If
        Next
        
        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Private Sub ShopMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmpX As Long, tmpY As Long
Dim i As Long
Dim DrawX As Long, DrawY As Long

    With GUI(GuiEnum.GUI_SHOP)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_SHOP Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        For i = ButtonEnum.Shop_Close To ButtonEnum.Shop_ScrollDown
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
            
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next

        '//Item
        For i = ShopAddY To ShopAddY + 8
            If i > 0 And i <= MAX_SHOP_ITEM Then
                If Shop(ShopNum).ShopItem(i).Num > 0 Then
                    DrawX = .X + (31 + ((4 + 127) * (((((i + 1) - ShopAddY) - 1) Mod 3))))
                    DrawY = .Y + (42 + ((4 + 78) * ((((i + 1) - ShopAddY) - 1) \ 3)))
                    
                    '//Button
                    If X >= DrawX + 12 And X <= DrawX + 12 + 103 And Y >= DrawY + 44 And Y <= DrawY + 44 + 25 Then
                        ShopButtonHover = i
                        If ShopButtonState = 0 Then
                            ShopButtonState = 1 '//Hover
                        End If
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                    
                    '//Icon
                    If X >= DrawX + 9 And X <= DrawX + 9 + 32 And Y >= DrawY + 6 And Y <= DrawY + 6 + 32 Then
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next

        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Private Sub ShopMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_SHOP)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_SHOP Then Exit Sub
        
        '//Loop through all items
        For i = ButtonEnum.Shop_Close To ButtonEnum.Shop_ScrollDown
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        Select Case i
                            Case ButtonEnum.Shop_Close
                                SendCloseShop
                            Case ButtonEnum.Shop_ScrollUp
                                If ShopAddY > 3 Then
                                    ShopAddY = ShopAddY - 3
                                End If
                            Case ButtonEnum.Shop_ScrollDown
                                If ShopAddY + 8 < ShopCountItem Then
                                    ShopAddY = ShopAddY + 3
                                End If
                        End Select
                    End If
                End If
            End If
        Next
        
        '//Check for dragging
        .InDrag = False
    End With
End Sub

' ***************
' ** TRADE **
' ***************
Private Sub TradeKeyPress(KeyAscii As Integer)
    '//Make sure it's visible
    If Not GUI(GuiEnum.GUI_TRADE).Visible Then Exit Sub
    
    If EditInputMoney Then
        If IsNumeric(KeyAscii) Then
            TradeInputMoney = InputText(TradeInputMoney, KeyAscii)
            If TradeInputMoney = vbNullString Then
                TradeInputMoney = 0
            End If
            TradeInputMoney = Round(Val(TradeInputMoney), 0)
            If Val(TradeInputMoney) > Player(MyIndex).Money Then
                TradeInputMoney = Player(MyIndex).Money
            End If
        End If
    End If
End Sub

Private Sub TradeMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim DrawX As Long, DrawY As Long

    With GUI(GuiEnum.GUI_TRADE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_TRADE
        
        '//Loop through all items
        For i = ButtonEnum.Trade_Close To ButtonEnum.Trade_AddMoney
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
        
        '//Trade Item
        If Buttons = vbRightButton Then
            i = IsTradeYourItem(CursorX, CursorY)
            If i > 0 Then
                CheckingTrade = 1
                OpenSelMenu SelMenuType.TradeItem, i
            End If
            
            i = IsTradeTheirItem(CursorX, CursorY)
            If i > 0 Then
                CheckingTrade = 2
                OpenSelMenu SelMenuType.TradeItem, i
            End If
        End If
        
        If YourTrade.TradeSet = NO Then
            If YourTrade.TradeMoney <> Val(TradeInputMoney) Then
                If CursorX >= .X + 66 And CursorX <= .X + 66 + 112 And CursorY >= .Y + 279 And CursorY <= .Y + 279 + 19 Then
                    EditInputMoney = True
                End If
            Else
                If CursorX >= .X + 66 And CursorX <= .X + 66 + 135 And CursorY >= .Y + 279 And CursorY <= .Y + 279 + 19 Then
                    EditInputMoney = True
                End If
            End If
        End If
        
        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Private Sub TradeMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmpX As Long, tmpY As Long
Dim i As Long

    With GUI(GuiEnum.GUI_TRADE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_TRADE Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        For i = ButtonEnum.Trade_Close To ButtonEnum.Trade_AddMoney
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
            
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next

        If YourTrade.TradeMoney <> Val(TradeInputMoney) Then
            If CursorX >= .X + 66 And CursorX <= .X + 66 + 112 And CursorY >= .Y + 279 And CursorY <= .Y + 279 + 19 Then
                IsHovering = True
                MouseIcon = 2 '//I-Beam
            End If
        Else
            If CursorX >= .X + 66 And CursorX <= .X + 66 + 135 And CursorY >= .Y + 279 And CursorY <= .Y + 279 + 19 Then
                IsHovering = True
                MouseIcon = 2 '//I-Beam
            End If
        End If

        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Private Sub TradeMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_TRADE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_TRADE Then Exit Sub

        '//Loop through all items
        For i = ButtonEnum.Trade_Close To ButtonEnum.Trade_AddMoney
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        Select Case i
                            Case ButtonEnum.Trade_Close
                                SendTradeState 0
                            Case ButtonEnum.Trade_Accept
                                SendTradeState 1
                            Case ButtonEnum.Trade_Decline
                                SendTradeState 0
                            Case ButtonEnum.Trade_Set
                                If YourTrade.TradeSet = NO Then
                                    SendSetTradeState YES
                                Else
                                    SendSetTradeState NO
                                End If
                            Case ButtonEnum.Trade_AddMoney
                                If YourTrade.TradeSet = NO Then
                                    If IsNumeric(TradeInputMoney) Then
                                        SendTradeUpdateMoney Val(TradeInputMoney)
                                    End If
                                End If
                        End Select
                    End If
                End If
            End If
        Next

        '//Check for dragging
        .InDrag = False
    End With
End Sub

' ***************
' ** Pokedex **
' ***************
Private Sub PokedexMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_POKEDEX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_POKEDEX
        
        '//Loop through all items
        For i = ButtonEnum.Pokedex_Close To ButtonEnum.Pokedex_ScrollDown
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                    Select Case i
                        Case ButtonEnum.Pokedex_ScrollUp
                            PokedexScrollUp = True
                            PokedexScrollDown = False
                            PokedexScrollTimer = GetTickCount
                        Case ButtonEnum.Pokedex_ScrollDown
                            PokedexScrollUp = False
                            PokedexScrollDown = True
                            PokedexScrollTimer = GetTickCount
                    End Select
                End If
            End If
        Next
        
        '//Check for scroll
        If CursorX >= .X + 7 And CursorX <= .X + 7 + 19 And CursorY >= .Y + PokedexScrollStartY + ((PokedexScrollEndY - PokedexScrollSize) - PokedexScrollY) And CursorY <= .Y + PokedexScrollStartY + ((PokedexScrollEndY - PokedexScrollSize) - PokedexScrollY) + PokedexScrollSize Then
            PokedexScrollHold = True
        End If
        
        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Private Sub PokedexMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmpX As Long, tmpY As Long
Dim i As Long

    With GUI(GuiEnum.GUI_POKEDEX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_POKEDEX Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        For i = ButtonEnum.Pokedex_Close To ButtonEnum.Pokedex_ScrollDown
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
            
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next
        
        '//Check for scroll
        If CursorX >= .X + 7 And CursorX <= .X + 7 + 19 And CursorY >= .Y + PokedexScrollStartY + ((PokedexScrollEndY - PokedexScrollSize) - PokedexScrollY) And CursorY <= .Y + PokedexScrollStartY + ((PokedexScrollEndY - PokedexScrollSize) - PokedexScrollY) + PokedexScrollSize Then
            IsHovering = True
            MouseIcon = 1 '//Select
        End If
        
        i = IsPokedexSlot(X, Y)
        If i > 0 Then
            IsHovering = True
            MouseIcon = 1 '//Select
            If Not PokedexInfoIndex = i + 1 Then
                PokedexInfoIndex = i + 1
                PokedexShowTimer = GetTickCount + 1000
            End If
        Else
            PokedexInfoIndex = 0
        End If
  
        '//Scroll moving
        If PokedexScrollHold Then
            '//Upward
            If CursorY < .Y + PokedexScrollStartY + ((PokedexScrollEndY - PokedexScrollSize) - PokedexScrollY) + (PokedexScrollSize / 2) Then
                If PokedexScrollY < PokedexScrollEndY - PokedexScrollSize Then
                    PokedexScrollY = (CursorY - (.Y + PokedexScrollStartY + (PokedexScrollEndY - PokedexScrollSize)) - (PokedexScrollSize / 2)) * -1
                    If PokedexScrollY >= PokedexScrollEndY - PokedexScrollSize Then PokedexScrollY = PokedexScrollEndY - PokedexScrollSize
                End If
            End If
            '//Downward
            If CursorY > .Y + PokedexScrollStartY + ((PokedexScrollEndY - PokedexScrollSize) - PokedexScrollY) + PokedexScrollSize - (PokedexScrollSize / 2) Then
                If PokedexScrollY > 0 Then
                    PokedexScrollY = (CursorY - (.Y + PokedexScrollStartY + (PokedexScrollEndY - PokedexScrollSize)) - PokedexScrollSize + (PokedexScrollSize / 2)) * -1
                    If PokedexScrollY <= 0 Then PokedexScrollY = 0
                End If
            End If
             
            PokedexScrollCount = (132 - PokedexScrollY)
            PokedexViewCount = ((PokedexScrollCount / MaxPokedexViewLine) / (132 / MaxPokedexViewLine)) * MaxPokedexViewLine
        End If
        
        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Private Sub PokedexMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_POKEDEX)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_POKEDEX Then Exit Sub

        '//Loop through all items
        For i = ButtonEnum.Pokedex_Close To ButtonEnum.Pokedex_ScrollDown
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        Select Case i
                            Case ButtonEnum.Pokedex_Close
                                If GUI(GuiEnum.GUI_POKEDEX).Visible Then
                                    GuiState GUI_POKEDEX, False
                                End If
                        End Select
                    End If
                End If
            End If
        Next
        
        '//Pokedex Scroll
        PokedexScrollHold = False

        '//Check for dragging
        .InDrag = False
    End With
End Sub

' ***************
' ** Pokemon Summary **
' ***************
Private Sub PokemonSummaryMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_POKEMONSUMMARY)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_POKEMONSUMMARY
        
        '//Loop through all items
        For i = ButtonEnum.PokemonSummary_Close To ButtonEnum.PokemonSummary_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
        
        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Private Sub PokemonSummaryMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmpX As Long, tmpY As Long
Dim i As Long

    With GUI(GuiEnum.GUI_POKEMONSUMMARY)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_POKEMONSUMMARY Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        For i = ButtonEnum.PokemonSummary_Close To ButtonEnum.PokemonSummary_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
            
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next

        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Private Sub PokemonSummaryMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_POKEMONSUMMARY)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_POKEMONSUMMARY Then Exit Sub
        
        '//Loop through all items
        For i = ButtonEnum.PokemonSummary_Close To ButtonEnum.PokemonSummary_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        Select Case i
                            Case ButtonEnum.PokemonSummary_Close
                                If GUI(GuiEnum.GUI_POKEMONSUMMARY).Visible Then
                                    GuiState GUI_POKEMONSUMMARY, False
                                End If
                                SummaryType = 0
                                SummarySlot = 0
                                SummaryData = 0
                        End Select
                    End If
                End If
            End If
        Next

        '//Check for dragging
        .InDrag = False
    End With
End Sub

' *************
' ** Relearn **
' *************
Private Sub RelearnMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim CanHover As Boolean, MoveNum As Long, MN As Long
Dim x2 As Long

    With GUI(GuiEnum.GUI_RELEARN)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_RELEARN
        
        '//Loop through all items
        For i = ButtonEnum.Relearn_Close To ButtonEnum.Relearn_ScrollUp
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
        
        If GUI(GuiEnum.GUI_MOVEREPLACE).Visible = False Then
            If MoveRelearnPokeNum > 0 Then
                For i = 1 To 5
                    CanHover = True
                    MoveNum = i + MoveRelearnCurPos
                    If MoveNum >= 0 And MoveNum <= MoveRelearnMaxIndex Then
                        If Pokemon(MoveRelearnPokeNum).Moveset(MoveNum).MoveNum > 0 Then
                            MN = Pokemon(MoveRelearnPokeNum).Moveset(MoveNum).MoveNum
                            '//Check if pokemon already learned the move or pokemon doesn't have enough level
                            If MoveRelearnPokeSlot > 0 Then
                                If PlayerPokemons(MoveRelearnPokeSlot).Num > 0 Then
                                    For x2 = 1 To MAX_MOVESET
                                        If PlayerPokemons(MoveRelearnPokeSlot).Moveset(x2).Num = MN Then
                                            CanHover = False
                                        End If
                                    Next
                                    If PlayerPokemons(MoveRelearnPokeSlot).Level < Pokemon(MoveRelearnPokeNum).Moveset(MoveNum).MoveLevel Then
                                        CanHover = False
                                    End If
                                    
                                    If CanHover Then
                                        If CursorX >= .X + 36 And CursorX <= .X + 36 + 198 And CursorY >= .Y + 46 + ((i - 1) * 48) And CursorY <= .Y + 46 + ((i - 1) * 48) + 42 Then
                                            SendRelearnMove MoveNum, MoveRelearnPokeSlot, MoveRelearnPokeNum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End If
        
        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Private Sub RelearnMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmpX As Long, tmpY As Long
Dim i As Long
Dim CanHover As Boolean, MoveNum As Long, MN As Long

    With GUI(GuiEnum.GUI_RELEARN)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_RELEARN Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        For i = ButtonEnum.Relearn_Close To ButtonEnum.Relearn_ScrollUp
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
            
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next
        
        If MoveRelearnPokeNum > 0 Then
            For i = 1 To 5
                CanHover = True
                MoveNum = i + MoveRelearnCurPos
                If MoveNum >= 0 And MoveNum <= MoveRelearnMaxIndex Then
                    If Pokemon(MoveRelearnPokeNum).Moveset(MoveNum).MoveNum > 0 Then
                        MN = Pokemon(MoveRelearnPokeNum).Moveset(MoveNum).MoveNum
                        '//Check if pokemon already learned the move or pokemon doesn't have enough level
                        If MoveRelearnPokeSlot > 0 Then
                            If PlayerPokemons(MoveRelearnPokeSlot).Num > 0 Then
                                For X = 1 To MAX_MOVESET
                                    If PlayerPokemons(MoveRelearnPokeSlot).Moveset(X).Num = MN Then
                                        CanHover = False
                                    End If
                                Next
                                If PlayerPokemons(MoveRelearnPokeSlot).Level < Pokemon(MoveRelearnPokeNum).Moveset(MoveNum).MoveLevel Then
                                    CanHover = False
                                End If
                                
                                If CanHover Then
                                    If CursorX >= .X + 36 And CursorX <= .X + 36 + 198 And CursorY >= .Y + 46 + ((i - 1) * 48) And CursorY <= .Y + 46 + ((i - 1) * 48) + 42 Then
                                        IsHovering = True
                                        MouseIcon = 1 '//Select
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If

        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Private Sub RelearnMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_RELEARN)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_RELEARN Then Exit Sub
        
        '//Loop through all items
        For i = ButtonEnum.Relearn_Close To ButtonEnum.Relearn_ScrollUp
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        Select Case i
                            Case ButtonEnum.Relearn_Close
                                If GUI(GuiEnum.GUI_RELEARN).Visible Then
                                    GuiState GUI_RELEARN, False
                                End If
                            Case ButtonEnum.Relearn_ScrollDown
                                If MoveRelearnCurPos < (MoveRelearnMaxIndex - 4) Then
                                    MoveRelearnCurPos = MoveRelearnCurPos + 1
                                End If
                            Case ButtonEnum.Relearn_ScrollUp
                                If MoveRelearnCurPos > 0 Then
                                    MoveRelearnCurPos = MoveRelearnCurPos - 1
                                End If
                        End Select
                    End If
                End If
            End If
        Next

        '//Check for dragging
        .InDrag = False
    End With
End Sub

' ***************
' ** Badge **
' ***************
Private Sub BadgeMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim PosX As Long, PosY As Long

    With GUI(GuiEnum.GUI_BADGE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_BADGE
        
        '//Loop through all items
        For i = ButtonEnum.Badge_Close To ButtonEnum.Badge_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
        
        '//Badge
        For i = 1 To MAX_BADGE
            If Player(MyIndex).Badge(i) > 0 Then
                PosX = .X + (84 + ((1 + 20) * (((i - 1) Mod 8))))
                PosY = .Y + (42 + ((10 + 20) * ((i - 1) \ 8)))

                '//Draw Icon
                'RenderTexture Tex_Gui(.Pic), PosX, PosY, TexX, TexY, 20, 20, 20, 20
                If CursorX >= PosX And CursorX <= PosX + 20 And CursorY >= PosY And CursorY <= PosY + 20 Then
                    FlyBadgeSlot = i
                    OpenChoiceBox "Do you want to fly to this badge's location?", CB_FLY
                    Exit For
                End If
            End If
        Next
        
        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Private Sub BadgeMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmpX As Long, tmpY As Long
Dim i As Long

    With GUI(GuiEnum.GUI_BADGE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_BADGE Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        For i = ButtonEnum.Badge_Close To ButtonEnum.Badge_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
            
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next
        
        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Private Sub BadgeMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_BADGE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_BADGE Then Exit Sub
        
        '//Loop through all items
        For i = ButtonEnum.Badge_Close To ButtonEnum.Badge_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        Select Case i
                            Case ButtonEnum.Badge_Close
                                If GUI(GuiEnum.GUI_BADGE).Visible Then
                                    GuiState GUI_BADGE, False
                                End If
                        End Select
                    End If
                End If
            End If
        Next

        '//Check for dragging
        .InDrag = False
    End With
End Sub

' ***************
' ** SlotMachine **
' ***************
Private Sub SlotMachineMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_SLOTMACHINE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_SLOTMACHINE
        
        '//Loop through all items
        For i = ButtonEnum.SlotMachine_Close To ButtonEnum.SlotMachine_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
        
        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Private Sub SlotMachineMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmpX As Long, tmpY As Long
Dim i As Long

    With GUI(GuiEnum.GUI_SLOTMACHINE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_SLOTMACHINE Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        For i = ButtonEnum.SlotMachine_Close To ButtonEnum.SlotMachine_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
            
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next
        
        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Private Sub SlotMachineMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_SLOTMACHINE)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_SLOTMACHINE Then Exit Sub
        
        '//Loop through all items
        For i = ButtonEnum.SlotMachine_Close To ButtonEnum.SlotMachine_Close
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        Select Case i
                            Case ButtonEnum.SlotMachine_Close
                                '//Close Slot Machine
                                
                        End Select
                    End If
                End If
            End If
        Next

        '//Check for dragging
        .InDrag = False
    End With
End Sub

Private Sub HandleChatMsg(ByVal chatText As String)
Dim chatMsg As String
Dim Command() As String
Dim cacheChatTab As String
Dim motdText As String
Dim i As Long

    chatMsg = chatText
    
    '//First, let the program check if we input any sign keys and check if they are valid command
    If Left$(chatMsg, 1) = "/" Then
        Command = Split(chatMsg, Space(1))
        Command(0) = LCase(Command(0))
        
        Select Case Command(0)
            '//////////////////////////////
            '/////// Player Command ///////
            '//////////////////////////////
            Case "/help"
                '//Normal Players command
                AddText "- Chat Command -", Pink
                AddText "[space] = you must press 'Spacebar' to trigger the command", White
                AddText "/map[space] = Map Message", White
                AddText "/all[space] = Global Message", White
                AddText "@playername[space] = Whisper", White
                AddText "/online = check who's online", White
                AddText "- Action Key -", Pink
                AddText "[" & Trim$(GetKeyCodeName(ControlKey(ControlEnum.KeyUp).cAsciiKey)) & ", " _
                        & Trim$(GetKeyCodeName(ControlKey(ControlEnum.KeyLeft).cAsciiKey)) & ", " _
                        & Trim$(GetKeyCodeName(ControlKey(ControlEnum.KeyDown).cAsciiKey)) & ", " _
                        & Trim$(GetKeyCodeName(ControlKey(ControlEnum.KeyRight).cAsciiKey)) & "] = Movement key", White
                AddText "[" & Trim$(GetKeyCodeName(ControlKey(ControlEnum.KeyAttack).cAsciiKey)) & "] = Attack key", White
                AddText "Hold [" & Trim$(GetKeyCodeName(ControlKey(ControlEnum.KeyCheckMove).cAsciiKey)) & "] and press [" _
                        & Trim$(GetKeyCodeName(ControlKey(ControlEnum.KeyMoveUp).cAsciiKey)) & ", " _
                        & Trim$(GetKeyCodeName(ControlKey(ControlEnum.KeyMoveLeft).cAsciiKey)) & ", " _
                        & Trim$(GetKeyCodeName(ControlKey(ControlEnum.KeyMoveDown).cAsciiKey)) & ", " _
                        & Trim$(GetKeyCodeName(ControlKey(ControlEnum.KeyMoveRight).cAsciiKey)) & "] to change set move", White
                '//Moderator Command
                If Player(MyIndex).Access >= ACCESS_MODERATOR Then
                    AddText "- Moderator Command -", Pink
                    AddText "/warpto map# = Warp to specific map", White
                    AddText "/warptome playername = Warp player to your position", White
                    AddText "/warpmeto playername = Warp yourself to player's position", White
                    AddText "/loc = View game statistic", White
                End If
                '//Mapper command
                If Player(MyIndex).Access >= ACCESS_MAPPER Then
                    AddText "- Mapper Command -", Pink
                    AddText "[Note: Mapper can use the commands of Moderator]", White
                    AddText "/editmap", White
                End If
                '//Developer command
                If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
                    AddText "- Developer Command -", Pink
                    AddText "[Note: Developer can use the commands of Moderator and Mapper]", White
                    AddText "/editnpc, /editpokemon, /edititem, /editanimation, /editmove, /editspawn, /editconversation, /editshop, /editquest", White
                    AddText "/getitem itemnum itemval = Get specific item [For testing purpose only]", White
                End If
                '//Administrator command
                If Player(MyIndex).Access >= ACCESS_CREATOR Then
                    AddText "- Administrator Command -", Pink
                    AddText "[Note: Administrator can use all commands]", White
                    AddText "/setaccess playername access = Change specific player's access", White
                    AddText "/motd msg = Change MOTD", White
                End If
            Case "/online"
                SendWhosOnline
            Case "/rank"
                GuiState GUI_RANK, True
            '/////////////////////////////////
            '/////// Moderator Command ///////
            '/////////////////////////////////
            Case "/kick"
                If Player(MyIndex).Access >= ACCESS_MODERATOR Then
                    If UBound(Command) < 1 Then
                        AddText "Usage: /kick playername", BrightRed
                        GoTo continue
                    End If
                    
                    SendKickPlayer (Command(1))
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/ban"
                If Player(MyIndex).Access >= ACCESS_MODERATOR Then
                    If UBound(Command) < 1 Then
                        AddText "Usage: /ban playername", BrightRed
                        GoTo continue
                    End If
                    
                    SendBanPlayer (Command(1))
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/mute"
                If Player(MyIndex).Access >= ACCESS_MODERATOR Then
                    If UBound(Command) < 1 Then
                        AddText "Usage: /mute playername", BrightRed
                        GoTo continue
                    End If
                    
                    SendMutePlayer (Command(1))
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/unmute"
                If Player(MyIndex).Access >= ACCESS_MODERATOR Then
                    If UBound(Command) < 1 Then
                        AddText "Usage: /unmute playername", BrightRed
                        GoTo continue
                    End If
                    
                    SendUnmutePlayer (Command(1))
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/warpto"
                If Player(MyIndex).Access >= ACCESS_MODERATOR Then
                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpto map#", BrightRed
                        GoTo continue
                    End If
                    
                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /warpto map#", BrightRed
                        GoTo continue
                    End If
                    
                    If PlayerPokemon(MyIndex).Num > 0 Then
                        AddText "Cannot warp", BrightRed
                        GoTo continue
                    End If
                    
                    SendWarpTo Val(Command(1))
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/warptome"
                If Player(MyIndex).Access >= ACCESS_MODERATOR Then
                    If UBound(Command) < 1 Then
                        AddText "Usage: /warptome playername", BrightRed
                        GoTo continue
                    End If
                    
                    If PlayerPokemon(MyIndex).Num > 0 Then
                        AddText "Cannot warp", BrightRed
                        GoTo continue
                    End If
                    
                    SendWarpToMe (Command(1))
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/warpmeto"
                If Player(MyIndex).Access >= ACCESS_MODERATOR Then
                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpmeto playername", BrightRed
                        GoTo continue
                    End If
                    
                    If PlayerPokemon(MyIndex).Num > 0 Then
                        AddText "Cannot warp", BrightRed
                        GoTo continue
                    End If
                    
                    SendWarpMeTo (Command(1))
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/loc"
                If Player(MyIndex).Access >= ACCESS_MODERATOR Then
                    ShowLoc = Not ShowLoc
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            '//////////////////////////////
            '/////// Mapper Command ///////
            '//////////////////////////////
            Case "/editmap"
                If Player(MyIndex).Access >= ACCESS_MAPPER Then
                    If GameSetting.Fullscreen = YES Then
                        AddText "You cannot open any editor in fullscreen mode", BrightRed
                        GoTo continue
                    Else
                        SendRequestEditMap
                    End If
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            '/////////////////////////////////
            '/////// Developer Command ///////
            '/////////////////////////////////
            Case "/editnpc"
                If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
                    If GameSetting.Fullscreen = YES Then
                        AddText "You cannot open any editor in fullscreen mode", BrightRed
                        GoTo continue
                    Else
                        SendRequestEditNpc
                    End If
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/editpokemon"
                If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
                    If GameSetting.Fullscreen = YES Then
                        AddText "You cannot open any editor in fullscreen mode", BrightRed
                        GoTo continue
                    Else
                        SendRequestEditPokemon
                    End If
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/edititem"
                If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
                    If GameSetting.Fullscreen = YES Then
                        AddText "You cannot open any editor in fullscreen mode", BrightRed
                        GoTo continue
                    Else
                        SendRequestEditItem
                    End If
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/editanimation"
                If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
                    If GameSetting.Fullscreen = YES Then
                        AddText "You cannot open any editor in fullscreen mode", BrightRed
                        GoTo continue
                    Else
                        SendRequestEditAnimation
                    End If
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/editmove"
                If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
                    If GameSetting.Fullscreen = YES Then
                        AddText "You cannot open any editor in fullscreen mode", BrightRed
                        GoTo continue
                    Else
                        SendRequestEditPokemonMove
                    End If
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/editspawn"
                If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
                    If GameSetting.Fullscreen = YES Then
                        AddText "You cannot open any editor in fullscreen mode", BrightRed
                        GoTo continue
                    Else
                        SendRequestEditSpawn
                    End If
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/editconversation"
                If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
                    If GameSetting.Fullscreen = YES Then
                        AddText "You cannot open any editor in fullscreen mode", BrightRed
                        GoTo continue
                    Else
                        SendRequestEditConversation
                    End If
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/editshop"
                If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
                    If GameSetting.Fullscreen = YES Then
                        AddText "You cannot open any editor in fullscreen mode", BrightRed
                        GoTo continue
                    Else
                        SendRequestEditShop
                    End If
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/editquest"
                If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
                    If GameSetting.Fullscreen = YES Then
                        AddText "You cannot open any editor in fullscreen mode", BrightRed
                        GoTo continue
                    Else
                        SendRequestEditQuest
                    End If
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/getitem"
                If Player(MyIndex).Access >= ACCESS_DEVELOPER Then
                    If UBound(Command) < 2 Then
                        AddText "Usage: /getitem itemnum itemval", BrightRed
                        GoTo continue
                    End If
                    
                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /getitem itemnum itemval", BrightRed
                        GoTo continue
                    End If
                    
                    If Not IsNumeric(Command(2)) Then
                        AddText "Usage: /getitem itemnum itemval", BrightRed
                        GoTo continue
                    End If
                    
                    SendGetItem Val(Command(1)), Val(Command(2))
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            '/////////////////////////////
            '/////// Owner Command ///////
            '/////////////////////////////
            Case "/setaccess"
                If Player(MyIndex).Access >= ACCESS_CREATOR Then
                    If UBound(Command) < 2 Then
                        AddText "Usage: /setaccess playername accessnumber", BrightRed
                        GoTo continue
                    End If
                    
                    If Not IsNumeric(Command(2)) Then
                        AddText "Usage: /setaccess playername accessnumber", BrightRed
                        GoTo continue
                    End If
                    
                    If Val(Command(2)) < 0 Or Val(Command(2)) > ACCESS_CREATOR Then
                        AddText "Usage: /setaccess playername accessnumber", BrightRed
                        GoTo continue
                    End If
                    
                    SendSetAccess (Command(1)), Val(Command(2))
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/motd"
                If Player(MyIndex).Access >= ACCESS_CREATOR Then
                    If UBound(Command) < 1 Then
                        AddText "Usage: /motd msg", BrightRed
                        GoTo continue
                    End If
                    
                    If Len(Trim$(Command(1))) <= 0 Then
                        AddText "Usage: /motd msg", BrightRed
                        GoTo continue
                    End If
                    
                    motdText = vbNullString
                    For i = 1 To UBound(Command)
                        motdText = motdText & (Trim$(Command(i))) & " "
                    Next
                    motdText = Trim$(motdText)
                    
                    '//Change MOTD
                    SendMOTD Trim$(motdText)
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/copymap"
                If Player(MyIndex).Access >= ACCESS_CREATOR Then
                    If UBound(Command) < 2 Then
                        AddText "Usage: /copymap destinationmap sourcemap", BrightRed
                        GoTo continue
                    End If
                    
                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /copymap destinationmap sourcemap", BrightRed
                        GoTo continue
                    End If
                    
                    If Not IsNumeric(Command(2)) Then
                        AddText "Usage: /copymap destinationmap sourcemap", BrightRed
                        GoTo continue
                    End If
                    
                    If MsgBox("Are you sure you want to copy map#" & Val(Command(2)) & " to map#" & Val(Command(1)), vbYesNo) = vbYes Then
                        SendCopyMap Val(Command(1)), Val(Command(2))
                    End If
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/giveitemto"
                If Player(MyIndex).Access >= ACCESS_CREATOR Then
                    If UBound(Command) < 3 Then
                        AddText "Usage: /giveitemto playername itemnum itemval", BrightRed
                        GoTo continue
                    End If
                    
                    If Not IsNumeric(Command(2)) Then
                        AddText "Usage: /giveitemto playername itemnum itemval", BrightRed
                        GoTo continue
                    End If
                    
                    If Not IsNumeric(Command(3)) Then
                        AddText "Usage: /giveitemto playername itemnum itemval", BrightRed
                        GoTo continue
                    End If
                    
                    SendGiveItemTo Trim$(Command(1)), (Command(2)), (Command(3))
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/givepokemonto"
                If Player(MyIndex).Access >= ACCESS_CREATOR Then
                    If UBound(Command) < 3 Then
                        AddText "Usage: /givepokemonto playername itemnum itemval", BrightRed
                        GoTo continue
                    End If
                    
                    If Not IsNumeric(Command(2)) Then
                        AddText "Usage: /givepokemonto playername itemnum itemval", BrightRed
                        GoTo continue
                    End If
                    
                    If Not IsNumeric(Command(3)) Then
                        AddText "Usage: /givepokemonto playername itemnum itemval", BrightRed
                        GoTo continue
                    End If
                    
                    SendGivePokemonTo Trim$(Command(1)), (Command(2)), (Command(3))
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/spawnpokemon"
                If Player(MyIndex).Access >= ACCESS_CREATOR Then
                    If UBound(Command) < 2 Then
                        AddText "Usage: /spawnpokemon mappokeslot shiny", BrightRed
                        GoTo continue
                    End If
                    
                    If Not IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                        AddText "Usage: /spawnpokemon mappokeslot shiny", BrightRed
                        GoTo continue
                    End If

                    SendSpawnPokemon (Command(1)), (Command(2))
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/stealthmode"
                If Player(MyIndex).Access >= ACCESS_MODERATOR Then
                    SendStealthMode
                Else
                    AddText "Invalid command!", BrightRed
                    GoTo continue
                End If
            Case "/test"
                Player(MyIndex).Level = Command(1)
            Case Else
                AddText "Invalid command!", BrightRed
                GoTo continue
        End Select
        
continue:
        MyChat = vbNullString
        Exit Sub
    End If
    
    '//Let the msg send through input tab
    cacheChatTab = LCase(ChatTab)
    
    Select Case cacheChatTab
        Case "/map"
            '//Map Msg
            SendMapMsg Trim$(MyChat)
        Case "/all"
            '//Global Msg
            SendGlobalMsg Trim$(MyChat)
        Case Else
            '//Player Msg
            SendPlayerMsg Trim$(ChatTab), Trim$(MyChat)
    End Select
    
    MyChat = vbNullString
End Sub

' **********
' ** Rank **
' **********
Private Sub RankMouseDown(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Dim CanHover As Boolean, MoveNum As Long, MN As Long
Dim x2 As Long

    With GUI(GuiEnum.GUI_RANK)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        '//Set to top most
        UpdateGuiOrder GUI_RANK
        
        '//Loop through all items
        For i = ButtonEnum.Rank_Close To ButtonEnum.Rank_ScrollUp
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateHover Then
                        Button(i).State = ButtonState.StateClick
                    End If
                End If
            End If
        Next
        
        '//Check for dragging
        .OldMouseX = CursorX - .X
        .OldMouseY = CursorY - .Y
        If .OldMouseY >= 0 And .OldMouseY <= 31 Then
            .InDrag = True
        End If
    End With
End Sub

Private Sub RankMouseMove(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmpX As Long, tmpY As Long
Dim i As Long
Dim CanHover As Boolean, MoveNum As Long, MN As Long

    With GUI(GuiEnum.GUI_RANK)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_RANK Then Exit Sub
        
        IsHovering = False
        
        '//Loop through all items
        For i = ButtonEnum.Rank_Close To ButtonEnum.Rank_ScrollUp
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateNormal Then
                        Button(i).State = ButtonState.StateHover
            
                        IsHovering = True
                        MouseIcon = 1 '//Select
                    End If
                End If
            End If
        Next

        '//Check for dragging
        If .InDrag Then
            tmpX = CursorX - .OldMouseX
            tmpY = CursorY - .OldMouseY
            
            '//Check if outbound
            If tmpX <= 0 Then tmpX = 0
            If tmpX >= Screen_Width - .Width Then tmpX = Screen_Width - .Width
            If tmpY <= 0 Then tmpY = 0
            If tmpY >= Screen_Height - .Height Then tmpY = Screen_Height - .Height
            
            .X = tmpX
            .Y = tmpY
        End If
    End With
End Sub

Private Sub RankMouseUp(Buttons As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    With GUI(GuiEnum.GUI_RANK)
        '//Make sure it's visible
        If Not .Visible Then Exit Sub
        
        If GuiVisibleCount <= 0 Then Exit Sub
        If Not GuiZOrder(GuiVisibleCount) = GuiEnum.GUI_RANK Then Exit Sub
        
        '//Loop through all items
        For i = ButtonEnum.Rank_Close To ButtonEnum.Rank_ScrollUp
            If CanShowButton(i) Then
                If CursorX >= .X + Button(i).X And CursorX <= .X + Button(i).X + Button(i).Width And CursorY >= .Y + Button(i).Y And CursorY <= .Y + Button(i).Y + Button(i).Height Then
                    If Button(i).State = ButtonState.StateClick Then
                        Button(i).State = ButtonState.StateNormal
                        Select Case i
                            Case ButtonEnum.Rank_Close
                                If GUI(GuiEnum.GUI_RANK).Visible Then
                                    GuiState GUI_RANK, False
                                End If
                            Case ButtonEnum.Rank_ScrollDown
                                If RankScroll < 13 Then
                                    RankScroll = RankScroll + 1
                                End If
                            Case ButtonEnum.Rank_ScrollUp
                                If RankScroll > 0 Then
                                    RankScroll = RankScroll - 1
                                End If
                        End Select
                    End If
                End If
            End If
        Next

        '//Check for dragging
        .InDrag = False
    End With
End Sub

Public Function FindFrontNPC() As Long
    Dim i As Long
    Dim Y As Long, X As Long
    
    If MyIndex <= 0 Or MyIndex > MAX_PLAYER Then Exit Function
    
    For i = 1 To Npc_HighIndex
        If MapNpc(i).Num > 0 Then
            Select Case Player(MyIndex).Dir
                Case DIR_UP
                    X = Player(MyIndex).X
                    For Y = Player(MyIndex).Y - 2 To Player(MyIndex).Y - 1
                        If Y >= 0 And Y <= Map.MaxY Then
                            If X = MapNpc(i).X And Y = MapNpc(i).Y Then
                                If Npc(MapNpc(i).Num).Convo > 0 Then
                                    FindFrontNPC = i
                                    Exit Function
                                End If
                            End If
                        End If
                    Next
                Case DIR_DOWN
                    X = Player(MyIndex).X
                    For Y = Player(MyIndex).Y + 1 To Player(MyIndex).Y + 2
                        If Y >= 0 And Y <= Map.MaxY Then
                            If X = MapNpc(i).X And Y = MapNpc(i).Y Then
                                If Npc(MapNpc(i).Num).Convo > 0 Then
                                    FindFrontNPC = i
                                    Exit Function
                                End If
                            End If
                        End If
                    Next
                Case DIR_LEFT
                    Y = Player(MyIndex).Y
                    For X = Player(MyIndex).X - 2 To Player(MyIndex).X - 1
                        If X >= 0 And X <= Map.MaxX Then
                            If X = MapNpc(i).X And Y = MapNpc(i).Y Then
                                If Npc(MapNpc(i).Num).Convo > 0 Then
                                    FindFrontNPC = i
                                    Exit Function
                                End If
                            End If
                        End If
                    Next
                Case DIR_RIGHT
                    Y = Player(MyIndex).Y
                    For X = Player(MyIndex).X + 1 To Player(MyIndex).X + 2
                        If X >= 0 And X <= Map.MaxX Then
                            If X = MapNpc(i).X And Y = MapNpc(i).Y Then
                                If Npc(MapNpc(i).Num).Convo > 0 Then
                                    FindFrontNPC = i
                                    Exit Function
                                End If
                            End If
                        End If
                    Next
            End Select
        End If
    Next
    FindFrontNPC = 0
End Function

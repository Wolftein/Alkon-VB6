Attribute VB_Name = "modCrafting"
'***************************************************
'Author: Anagrama
'Last Modification: 05/04/2017
'05/04/2017: Anagrama - Modulo contenedor de todos los metodos y propiedades que forman parte del sistema unificado de crafteo complejo.
'***************************************************
Option Explicit

Public Function PuedeConstruir(ByVal UserIndex As Integer, ByVal CraftingGroup As Integer, ByVal RecipeIndex As Integer) As Boolean
'***************************************************
'Author: Anagrama
'Last Modification: 01/04/2017
'01/04/2017: Anagrama - Valida si el usuario puede construir los items solicitado, dentro de la estacion de crafteo solicitada.
'***************************************************
On Error GoTo ErrHandler
    Dim WeaponIndex As Integer
    Dim ProfessionType As Integer
    
    WeaponIndex = UserList(UserIndex).Invent.WeaponEqpObjIndex
    

    ProfessionType = ObjData(WeaponIndex).ProfessionType
    If ProfessionType <= 0 Or ProfessionType > UBound(Professions) Then
        Call WriteConsoleMsg(UserIndex, "La herramienta utilizada no pertenece a una profesión válida.", FontTypeNames.FONTTYPE_INFO)
        Call WriteStopWorking(UserIndex)
        Exit Function
    End If
    
    With Professions(ProfessionType)
        If Not .Enabled Then
            Call WriteConsoleMsg(UserIndex, "Esta profesión se encuentra deshabilitada. Intenta denuevo más tarde.", FontTypeNames.FONTTYPE_INFO)
            Call WriteStopWorking(UserIndex)
            Exit Function
        End If
                
        If .CraftingRecipeGroupsQty <= 0 Or CraftingGroup <= 0 Or CraftingGroup > .CraftingRecipeGroupsQty Then
            Call WriteConsoleMsg(UserIndex, "Receta inválida.", FontTypeNames.FONTTYPE_INFO)
            Call WriteStopWorking(UserIndex)
            Exit Function
        End If
        
        If .CraftingRecipeGroups(CraftingGroup).RecipesQty <= 0 Or RecipeIndex <= 0 Or RecipeIndex > .CraftingRecipeGroups(CraftingGroup).RecipesQty Then
            Call WriteConsoleMsg(UserIndex, "Receta inválida.", FontTypeNames.FONTTYPE_INFO)
            Call WriteStopWorking(UserIndex)
            Exit Function
        End If
        
        ' Validate skills
        Dim SkillsHerreria As Byte, SkillsCarpinteria As Byte, SkillsTailoring As Byte
    
        SkillsHerreria = GetSkills(UserIndex, eSkill.Herreria)
        SkillsCarpinteria = GetSkills(UserIndex, eSkill.Carpinteria)
        SkillsTailoring = GetSkills(UserIndex, eSkill.Sastreria)
        
        
        If SkillsHerreria < .CraftingRecipeGroups(CraftingGroup).Recipes(RecipeIndex).BlacksmithSkillNeeded Or _
            SkillsCarpinteria < .CraftingRecipeGroups(CraftingGroup).Recipes(RecipeIndex).CarpenterSkillNeeded Or _
            SkillsTailoring < .CraftingRecipeGroups(CraftingGroup).Recipes(RecipeIndex).TailoringSkillNeeded Then
            
            Call WriteConsoleMsg(UserIndex, "No posees los skills necesarios para construir este objeto.", FontTypeNames.FONTTYPE_INFO)
            Call WriteStopWorking(UserIndex)
            Exit Function
        End If
        
        PuedeConstruir = ConstruirObjeto(UserIndex, ProfessionType, CraftingGroup, RecipeIndex)
        
    End With

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function PuedeConstruir de modCrafting.bas")
End Function

Public Function ConstruirObjeto(ByVal UserIndex As Integer, ByVal ProfessionType As Byte, ByVal CraftingGroup As Byte, ByVal RecipeIndex As Integer) As Boolean
'***************************************************
'Author: Anagrama
'Last Modification: 01/04/2017
'01/04/2017: Anagrama - Construye el objeto indicado si es posible.
'***************************************************
On Error GoTo ErrHandler
    Dim CantidadItems As Long
    Dim OtroUserIndex As Integer
    
    With UserList(UserIndex)
        If isTradingWithUser(UserIndex) Then
            OtroUserIndex = getTradingUser(UserIndex)
                
            If (OtroUserIndex > 0) And (OtroUserIndex <= MaxUsers) Then
                Call WriteConsoleMsg(UserIndex, "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(UserIndex)
            End If
        End If
            
        .Construir.Cantidad = .Construir.Cantidad

        If .Construir.Cantidad < 0 Then .Construir.Cantidad = 0
        If .Construir.Cantidad > 10000 Then .Construir.Cantidad = 10000
        
        If .Construir.Cantidad = 0 Then
            Call WriteStopWorking(UserIndex)
            Exit Function
        End If
        
        ' Check user's stamina
        If .Stats.MinSta < ConstantesTrabajo.EsfuerzoExcavarGeneral Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        .Stats.MinSta = .Stats.MinSta - ConstantesTrabajo.EsfuerzoExcavarGeneral
        Call WriteUpdateSta(UserIndex)

        ' Check if it has the required materials
        If Not TieneMateriales(UserIndex, Professions(ProfessionType).CraftingRecipeGroups(CraftingGroup).Recipes(RecipeIndex).Materials, .Construir.Cantidad) Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficientes materiales.", FontTypeNames.FONTTYPE_INFO)
            Call WriteStopWorking(UserIndex)
            Exit Function
        End If
        
        Call QuitarMateriales(UserIndex, Professions(ProfessionType).CraftingRecipeGroups(CraftingGroup).Recipes(RecipeIndex).Materials, .Construir.Cantidad)

                
        ' Multiply the amount of items requested to craft by the amount produced per crafting configured in the recipe
        Dim MiObj As Obj
        Dim QtyToCraft As Long
        CantidadItems = .Construir.Cantidad * Professions(ProfessionType).CraftingRecipeGroups(CraftingGroup).Recipes(RecipeIndex).ProduceAmount
        MiObj.ObjIndex = Professions(ProfessionType).CraftingRecipeGroups(CraftingGroup).Recipes(RecipeIndex).ObjIndex
        
        QtyToCraft = CantidadItems
        
        While QtyToCraft > 0
            MiObj.Amount = IIf(QtyToCraft > MAX_INVENTORY_OBJS, MAX_INVENTORY_OBJS, QtyToCraft)
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)
            End If
            QtyToCraft = QtyToCraft - MiObj.Amount
        Wend

        ' Notify the user
        Select Case ObjData(MiObj.ObjIndex).ObjType
            Case eOBJType.otWeapon
                Call WriteConsoleMsg(UserIndex, "¡Has construido " & IIf(CantidadItems > 1, CantidadItems & " armas!", "el arma!"), FontTypeNames.FONTTYPE_INFO, eMessageType.Trabajo)
            Case eOBJType.otESCUDO
                Call WriteConsoleMsg(UserIndex, "¡Has construido " & IIf(CantidadItems > 1, CantidadItems & " escudos!", "el escudo!"), FontTypeNames.FONTTYPE_INFO, eMessageType.Trabajo)
            Case Is = eOBJType.otCASCO
                Call WriteConsoleMsg(UserIndex, "¡Has construido " & IIf(CantidadItems > 1, CantidadItems & " cascos!", "el casco!"), FontTypeNames.FONTTYPE_INFO, eMessageType.Trabajo)
            Case eOBJType.otArmadura
                Call WriteConsoleMsg(UserIndex, "¡Has construido " & IIf(CantidadItems > 1, CantidadItems & " armaduras!", "la armadura!"), FontTypeNames.FONTTYPE_INFO, eMessageType.Trabajo)
            Case Else
                Call WriteConsoleMsg(UserIndex, "¡Has construido " & IIf(CantidadItems > 1, CantidadItems & " objetos!", "el objeto!"), FontTypeNames.FONTTYPE_INFO, eMessageType.Trabajo)
        End Select
               
        'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
        If ObjData(MiObj.ObjIndex).Log = 1 Then
            Call LogDesarrollo(.Name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name)
        End If
        
        Call SubirSkill(UserIndex, Professions(ProfessionType).SkillNumber, True)
        
        If Professions(ProfessionType).SuccessFx > 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Professions(ProfessionType).SuccessFx, .Pos.X, .Pos.Y, .Char.CharIndex))
        End If
                        
        ConstruirObjeto = True
    End With

  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function ConstruirObjeto de modCrafting.bas")
End Function

Public Function TieneMateriales(ByVal UserIndex As Integer, ByRef CraftingMaterials() As tCraftingItem, ByVal CantidadItems As Integer) As Boolean
On Error GoTo ErrHandler:
    Dim I As Integer
    

    For I = 1 To UBound(CraftingMaterials)
        If Not TieneObjetos(CraftingMaterials(I).ObjIndex, CraftingMaterials(I).Amount * CantidadItems, UserIndex) Then
            Exit Function
        End If
    Next I

    TieneMateriales = True
    
  Exit Function
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Function TieneMateriales de modCrafting.bas")
End Function

Public Sub QuitarMateriales(ByVal UserIndex As Integer, ByRef CraftingMaterials() As tCraftingItem, ByVal CantidadItems As Integer)
On Error GoTo ErrHandler:
    Dim I As Integer
    
    If CantidadItems < 1 Or UBound(CraftingMaterials) <= 0 Then Exit Sub
        
    For I = 1 To UBound(CraftingMaterials)
        Call QuitarObjetos(CraftingMaterials(I).ObjIndex, CraftingMaterials(I).Amount * CantidadItems, UserIndex)
    Next I

  Exit Sub
  
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub QuitarMateriales de modCrafting.bas")
End Sub

Public Sub CloseWorkerStore(ByVal UserIndex)
On Error GoTo ErrHandler:

    With UserList(UserIndex)
        Erase .CraftingStore.Items
        
        Call WriteConsoleMsg(UserIndex, "Cerraste tu tienda de construcción.", FontTypeNames.FONTTYPE_INFO)
        
        If .CraftingStore.CraftedObjectsQty > 0 Then
            Call WriteConsoleMsg(UserIndex, "Mientras estuvo abierta, construíste " & .CraftingStore.CraftedObjectsQty & " objetos y ganaste " & .CraftingStore.MoneyEarned & " monedas de oro.", FontTypeNames.FONTTYPE_INFO)
        End If
        
        .CraftingStore.ItemsQty = 0
        .CraftingStore.IsOpen = False
        .CraftingStore.LastCraftedObjectAt = DateSerial(1900, 1, 1)
        .CraftingStore.MoneyEarned = 0
        .CraftingStore.CraftedObjectsQty = 0
        
        .OverHeadIcon = 0
        
        ' Send the change
        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
    End With
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CloseWorkerStore de modCrafting.bas")
End Sub

Public Sub CreateWorkerStoreTest(ByVal UserIndex As Integer)

    Dim Items() As tCraftingStoreItem
    
    ReDim Items(1 To 2) As tCraftingStoreItem
    
    With Items(1)
        .Recipe = 1
        .ConstructionPrice = 100
    End With
    
    With Items(2)
        .Recipe = 2
        .ConstructionPrice = 200
    End With
    
    Call CreateWorkerStore(UserIndex, Items)
        
End Sub

Public Sub CloseWorkerStoreTest(ByVal UserIndex As Integer)
    Call CloseWorkerStore(UserIndex)
End Sub

Public Sub CreateWorkerStore(ByVal UserIndex As Integer, ByRef CraftingList() As tCraftingStoreItem)
On Error GoTo ErrHandler:

    With UserList(UserIndex)
        
        If .CraftingStore.IsOpen Then
            Call WriteConsoleMsg(UserIndex, "Tu tienda ya se encuentra abierta.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .Invent.WeaponEqpObjIndex <= 0 Then
            Call WriteConsoleMsg(UserIndex, "Necesitas equipar una herramienta para abrir una tienda de trabajo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim ProfessionType As Byte
        ProfessionType = ObjData(.Invent.WeaponEqpObjIndex).ProfessionType
        
        If ProfessionType <= 0 Or ProfessionType > UBound(Professions) Then
            Call WriteConsoleMsg(UserIndex, "Necesitas equipar una herramienta para abrir una tienda de trabajo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim I As Integer
        Dim Cancel As Boolean
        Dim ValidationErrorMessage As String
        ReDim .CraftingStore.Items(1 To UBound(CraftingList))
        Dim Recipe As Integer
        Dim RecipeIndex As Integer
        Dim RecipeGroup As Byte
        
        ' Validate the recipes
        For I = 1 To UBound(CraftingList)
        
            RecipeIndex = CraftingList(I).RecipeIndex
            RecipeGroup = CraftingList(I).RecipeGroup
            
            If CraftingList(I).RecipeIndex > Professions(ProfessionType).CraftingRecipeGroups(RecipeGroup).RecipesQty Or RecipeIndex < 1 Then
                ValidationErrorMessage = "Receta inválida"
                Exit For
            End If
            
            ' Validate if the worker has access to this recipe based on the skills
            If Professions(ProfessionType).CraftingRecipeGroups(RecipeGroup).Recipes(RecipeIndex).BlacksmithSkillNeeded > GetSkills(UserIndex, eSkill.Herreria) Then
                ValidationErrorMessage = "No tienes los suficientes puntos en herrería."
                Exit For
            End If
            If Professions(ProfessionType).CraftingRecipeGroups(RecipeGroup).Recipes(RecipeIndex).CarpenterSkillNeeded > GetSkills(UserIndex, eSkill.Carpinteria) Then
                ValidationErrorMessage = "No tienes los suficientes puntos en carpintería."
                Exit For
            End If
            
            If Professions(ProfessionType).CraftingRecipeGroups(RecipeGroup).Recipes(RecipeIndex).TailoringSkillNeeded > GetSkills(UserIndex, eSkill.Sastreria) Then
                ValidationErrorMessage = "No tienes los suficientes puntos en sastrería."
                Exit For
            End If
            
            .CraftingStore.Items(I).Recipe = RecipeIndex
            .CraftingStore.Items(I).RecipeItem = Professions(ProfessionType).CraftingRecipeGroups(RecipeGroup).Recipes(RecipeIndex).ObjIndex
            .CraftingStore.Items(I).RecipeIndex = RecipeIndex
            .CraftingStore.Items(I).ConstructionPrice = CraftingList(I).ConstructionPrice
            .CraftingStore.Items(I).MaterialsPrice = CraftingList(I).MaterialsPrice
            .CraftingStore.Items(I).RecipeGroup = CraftingList(I).RecipeGroup
            .CraftingStore.ItemsQty = .CraftingStore.ItemsQty + 1
            .CraftingStore.InstanceId = RandomString(10)
        
        Next I
        
        If ValidationErrorMessage <> vbNullString Then
            Call WriteConsoleMsg(UserIndex, ValidationErrorMessage, FontTypeNames.FONTTYPE_INFO)
            ' Close and reset the crafting store because it's invalid at this point
            Call CloseWorkerStore(UserIndex)
            Exit Sub
        End If
                
        ' The store is now open!
        .CraftingStore.IsOpen = True
        .CraftingStore.CraftedObjectsQty = 0
        .CraftingStore.MoneyEarned = 0
        .CraftingStore.ProfessionType = ProfessionType
    
        Call WriteConsoleMsg(UserIndex, "Tu tienda de construcción se encuentra abierta.", FontTypeNames.FONTTYPE_INFO)
        
        ' Write things related to the character icon
        .OverHeadIcon = Constantes.CraftingStoreOverheadIcon
        
        ' Send the change
        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
        
        Call WriteWorkerStore_Open(UserIndex)
        
    End With
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CreateWorkerStore de modCrafting.bas")
End Sub

Public Sub CraftItemOnDemand(ByVal UserIndex As Integer, ByVal WorkerIndex As Integer, ByVal StoreRecipeIndex As Integer, ByRef InstanceId As String)
On Error GoTo ErrHandler:
    
    Dim SelectedRecipe As Integer
    Dim CraftingStoreItem As tCraftingStoreItem

    WorkerIndex = UserList(UserIndex).flags.TargetUser
    
    If WorkerIndex <= 0 Then
        Exit Sub
    End If
    
    If WorkerIndex = UserIndex Then
        Call WriteErrorMsg(UserIndex, "No puedes utilizar tu propia tienda.")
        Exit Sub
    End If
    
    If FreeInventorySlots(UserIndex) = 0 Then
        Call WriteErrorMsg(UserIndex, "No tienes suficiente espacio en el inventario.")
        Exit Sub
    End If
            
    ' Check if workerindex is online
    With UserList(WorkerIndex)
    
        If Not .CraftingStore.IsOpen Then
            Call WriteConsoleMsg(UserIndex, "El usuario no tiene una tienda abierta.", FontTypeNames.FONTTYPE_INFO)
            Call WriteCloseForm(UserIndex, "frmCraftingStore")
            Exit Sub
        End If
        
        If .CraftingStore.InstanceId <> InstanceId Then
            Call WriteConsoleMsg(UserIndex, "Esta tienda ya no se encuentra disponible porque sus objetos o precios cambiaron.", FontTypeNames.FONTTYPE_INFO)
            Call WriteCloseForm(UserIndex, "frmCraftingStore")
            Exit Sub
        End If
        
        If .CraftingStore.ItemsQty = 0 Then
            Call WriteErrorMsg(UserIndex, "Esta tienda no tiene objetos disponibles para construir.")
            Exit Sub
        End If
        
        If StoreRecipeIndex < 1 Or StoreRecipeIndex > .CraftingStore.ItemsQty Then
            Call WriteErrorMsg(UserIndex, "El objeto seleccionado para construir es inválido.")
            Exit Sub
        End If
        
        CraftingStoreItem = .CraftingStore.Items(StoreRecipeIndex)
        Dim ProfessionType As Byte
        Dim CraftingGroup As Integer
        ProfessionType = .CraftingStore.ProfessionType
        CraftingGroup = CraftingStoreItem.RecipeGroup
        
        ' Check if the worker has the required skills to craft this recipe.
        If Professions(ProfessionType).CraftingRecipeGroups(CraftingGroup).Recipes(CraftingStoreItem.RecipeIndex).BlacksmithSkillNeeded > GetSkills(WorkerIndex, eSkill.Herreria) Or _
            Professions(ProfessionType).CraftingRecipeGroups(CraftingGroup).Recipes(CraftingStoreItem.RecipeIndex).CarpenterSkillNeeded > GetSkills(WorkerIndex, eSkill.Carpinteria) Or _
            Professions(ProfessionType).CraftingRecipeGroups(CraftingGroup).Recipes(CraftingStoreItem.RecipeIndex).TailoringSkillNeeded > GetSkills(WorkerIndex, eSkill.Sastreria) Then
        
            Call WriteErrorMsg(UserIndex, UserList(WorkerIndex).Name & " no tiene los skills suficientes para construir este objeto.")
            Exit Sub
        End If

        Dim I As Integer
        Dim CraftObj As tCraftingItem
        Dim FullConstructionPrice As Long
        
        FullConstructionPrice = CraftingStoreItem.ConstructionPrice + CraftingStoreItem.MaterialsPrice
        
        ' Validate if the user has the required gold
        If UserList(UserIndex).Stats.GLD < FullConstructionPrice Then
            Call WriteErrorMsg(UserIndex, "No tienes oro suficiente para pagar la contrucción de este objeto.")
            Exit Sub
        End If
        
        ' Validate if it has all the objects required by the recipe
        Dim UserIndexMaterials As Integer
        UserIndexMaterials = IIf(.CraftingStore.StoreType = CustomerMaterials, UserIndex, WorkerIndex)
        
        For I = 1 To Professions(ProfessionType).CraftingRecipeGroups(CraftingGroup).Recipes(CraftingStoreItem.RecipeIndex).MaterialsQty
            CraftObj = Professions(ProfessionType).CraftingRecipeGroups(CraftingGroup).Recipes(CraftingStoreItem.RecipeIndex).Materials(I)
            If Not TieneObjetos(CraftObj.ObjIndex, CraftObj.Amount, UserIndexMaterials) Then
                Call WriteErrorMsg(UserIndex, UserList(WorkerIndex).Name & " no tiene materiales suficientes para la construcción de este objeto.")
                Exit Sub
            End If
        Next I
        
        ' Remove the objects from the inventory
        For I = 1 To Professions(ProfessionType).CraftingRecipeGroups(CraftingGroup).Recipes(CraftingStoreItem.RecipeIndex).MaterialsQty
            CraftObj = Professions(ProfessionType).CraftingRecipeGroups(CraftingGroup).Recipes(CraftingStoreItem.RecipeIndex).Materials(I)
            
            Call QuitarObjetos(CraftObj.ObjIndex, CraftObj.Amount, UserIndexMaterials)
        Next I
        
        ' Create the item and give it to the player
        Dim CraftedObj As Obj
        CraftedObj.ObjIndex = CraftingStoreItem.RecipeItem
        CraftedObj.Amount = 1
                    
        ' Substract the money required and give it to the worker
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - FullConstructionPrice
        Call WriteUpdateGold(UserIndex)
        
        UserList(WorkerIndex).Stats.GLD = UserList(WorkerIndex).Stats.GLD + FullConstructionPrice
        Call WriteUpdateGold(WorkerIndex)
        
        ' Increment historic values for this store instance
        UserList(WorkerIndex).CraftingStore.MoneyEarned = UserList(WorkerIndex).CraftingStore.MoneyEarned + FullConstructionPrice
        UserList(WorkerIndex).CraftingStore.CraftedObjectsQty = UserList(WorkerIndex).CraftingStore.CraftedObjectsQty + 1
        UserList(WorkerIndex).CraftingStore.LastCraftedObjectAt = GetTickCount()
        
        Call MeterItemEnInventario(UserIndex, CraftedObj)
        
        Call WriteErrorMsg(UserIndex, "Compraste 1 " & ObjData(CraftedObj.ObjIndex).Name & " en la tienda de construcción de " & UserList(WorkerIndex).Name)

        Call WriteWorkerStore_ItemCraftedNotification(WorkerIndex, UserList(UserIndex).Name, CraftedObj.ObjIndex, CraftedObj.Amount, FullConstructionPrice)
    End With
    
    Exit Sub
ErrHandler:
  Call LogError("Error" & Err.Number & "(" & Err.Description & ") en Sub CraftItemOnDemand de modCrafting.bas")
End Sub


Public Function IsStoreOpenNearby(ByVal UserIndex As Integer) As Boolean
    ' Find players close to the user with an oepn store
    
    Dim query() As Collision.UUID
    Dim TargetIndex As Integer
    Dim I As Integer
    
    For I = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
        TargetIndex = query(I).Name
        If UserList(TargetIndex).CraftingStore.IsOpen And _
            Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(TargetIndex).Pos.X, UserList(TargetIndex).Pos.Y) <= 2 Then
                                                                    
            IsStoreOpenNearby = True
            Exit Function
        End If
    Next I
    
    IsStoreOpenNearby = False

End Function

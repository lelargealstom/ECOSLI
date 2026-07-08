'--------------------------------------------------------------------------------------------------------
'
'                                       Generateur Livrables SLI
'                                       Module F_EMMST_Remplissage
'                                       Developpement : LE LARGE Thomas (LGM/Alstom)
'                                       Module contenant toutes les fonctions de remplissage EMMST
'                                       Creation : Fevrier 2026
'
'-----------------------------------------------------------------------------------------------------------
Option Compare Text
Option Explicit
' =========================================================================
' PROCEDURE: REMPLIR INFOS MINIMALES EMMST
' =========================================================================
Public Sub RemplirInfosMinimalesEMMST(ByRef dest As Worksheet, ByRef index As Worksheet, ByVal indexDTR As Integer)
' Remplit uniquement DTR et Designation quand EMMST non trouvee en BDD
    
    Dim r As New EMMST_Ranges
    Call r.Init(dest)
    
    r.Designation.Value = index.Cells(indexDTR + 5, 3)
    r.dtr.Value = tDTR(indexDTR, 1)
    
End Sub

' =========================================================================
' PROCEDURE: REMPLIR CHAMPS GENERAUX
' =========================================================================
Public Sub RemplirChampsGeneraux(ByRef r As EMMST_Ranges, ByRef index As Worksheet, ByVal indexDTR As Integer, ByVal ligne As Integer)
' Remplit les champs generaux : Spare/Tool, Designation, DTR
    
    ' Spare or Tool
    If IsNull(EMMST_tab(ligne, EMMST_type.TypeOutil)) Then
        r.spare.Value = ""
    ElseIf EMMST_tab(ligne, EMMST_type.TypeOutil) = "Rechange" Then
        r.spare.Value = "X"
    Else
        r.Tool.Value = "X"
    End If
    
    ' Product name
    r.Designation.Value = index.Cells(indexDTR + 5, 3)
    
    ' DTR
    r.dtr.Value = tDTR(indexDTR, 1)
    
End Sub

' =========================================================================
' PROCEDURE: REMPLIR CHAMPS DIMENSIONS
' =========================================================================
Public Sub RemplirChampsDimensions(ByRef r As EMMST_Ranges, ByVal ligne As Integer)
' Remplit dimensions nues et conditionnees
    
    ' Bare Dimensions
    r.blength.Value = EMMST_tab(ligne, EMMST_type.DimNuLongueur)
    r.bwidth.Value = EMMST_tab(ligne, EMMST_type.DimNuLargeur)
    r.bheight.Value = EMMST_tab(ligne, EMMST_type.DimNuHauteur)
    r.bmass.Value = EMMST_tab(ligne, EMMST_type.DimNuPoids)
    
    ' Packaging Dimensions
    r.plength.Value = EMMST_tab(ligne, EMMST_type.DimCondLongueur)
    r.pwidth.Value = EMMST_tab(ligne, EMMST_type.DimCondLargeur)
    r.pheight.Value = EMMST_tab(ligne, EMMST_type.DimCondHauteur)
    r.pmass.Value = EMMST_tab(ligne, EMMST_type.DimCondPoids)
    
End Sub

' =========================================================================
' PROCEDURE: REMPLIR CHAMPS EMBALLAGE
' =========================================================================
Public Sub RemplirChampsEmballage(ByRef r As EMMST_Ranges, ByVal ligne As Integer)
' Remplit champs emballage : peremption, conditionnement, reutilisable
    
    ' Lapsing (Date de peremption)
    If EMMST_tab(ligne, EMMST_type.DatePeremption) = 0 Then
        r.lapsing_no.Value = "X"
    ElseIf EMMST_tab(ligne, EMMST_type.DatePeremption) = -1 Then
        r.lapsing_yes.Value = "X"
    Else
        r.lapsing_no.Value = ""
    End If
    
    ' Expiry date
    r.expiry_date.Value = EMMST_tab(ligne, EMMST_type.DatePeremption)
    
    ' Multiple packaging
    r.multiple_packaging.Value = EMMST_tab(ligne, EMMST_type.UniteCond)
    
    ' Reusable
    If EMMST_tab(ligne, EMMST_type.Reutilisable) = 0 Then
        r.non_reusable_box.Value = "X"
    ElseIf EMMST_tab(ligne, EMMST_type.Reutilisable) = -1 Then
        r.reusable_box.Value = "X"
    Else
        r.reusable_box.Value = ""
    End If
    
End Sub

' =========================================================================
' PROCEDURE: REMPLIR CHAMPS MARQUAGE
' =========================================================================
Public Sub RemplirChampsMarquage(ByRef r As EMMST_Ranges, ByVal ligne As Integer)
' Remplit champs marquage/etiquetage
    
    ' Labelling on packaging
    If EMMST_tab(ligne, EMMST_type.EtiquetteEmballage) = 0 Then
        r.pack_label_no.Value = "X"
    ElseIf EMMST_tab(ligne, EMMST_type.EtiquetteEmballage) = -1 Then
        r.pack_label_yes.Value = "X"
    Else
        r.pack_label_no.Value = "X"
    End If
    
    ' Labelling on equipment
    If EMMST_tab(ligne, EMMST_type.MarquageEquipement) = 0 Then
        r.equip_label_no.Value = "X"
    ElseIf EMMST_tab(ligne, EMMST_type.MarquageEquipement) = -1 Then
        r.equip_label_yes.Value = "X"
    Else
        r.equip_label_no.Value = ""
    End If
    
End Sub

' =========================================================================
' PROCEDURE: REMPLIR CHAMPS MANUTENTION
' =========================================================================
Public Sub RemplirChampsManutention(ByRef r As EMMST_Ranges, ByVal ligne As Integer)
' Remplit champs manutention/handling
    
    ' Handling mean
    If EMMST_tab(ligne, EMMST_type.ManutentionEquipement) = 0 Then
        r.handling_mean_no.Value = "X"
    ElseIf EMMST_tab(ligne, EMMST_type.ManutentionEquipement) = -1 Then
        r.handling_mean_yes.Value = "X"
    Else
        r.handling_mean_no.Value = ""
    End If
    
    ' Specific tool
    If EMMST_tab(ligne, EMMST_type.OutillageSpeTransport) = 0 Then
        r.spe_tool_no.Value = "X"
    ElseIf EMMST_tab(ligne, EMMST_type.OutillageSpeTransport) = -1 Then
        r.spe_tool_yes.Value = "X"
    Else
        r.spe_tool_no.Value = ""
    End If
    
    ' Precautions
    If EMMST_tab(ligne, EMMST_type.PrecautionManutention) = 0 Then
        r.man_prec_no.Value = "X"
    ElseIf EMMST_tab(ligne, EMMST_type.PrecautionManutention) = -1 Then
        r.man_prec_yes.Value = "X"
    Else
        r.man_prec_no.Value = ""
    End If
    
End Sub

' =========================================================================
' PROCEDURE: REMPLIR CHAMPS STOCKAGE
' =========================================================================
Public Sub RemplirChampsStockage(ByRef r As EMMST_Ranges, ByVal ligne As Integer)
' Remplit champs stockage/storage
    
    ' Storage position
    Select Case EMMST_tab(ligne, EMMST_type.PositionStockage)
        Case "Horizontal"
            r.pos_hor.Value = "X"
        Case "Vertical"
            r.pos_ver.Value = "X"
        Case "Indifferent"
            r.pos_indif.Value = "X"
        Case Else
            r.pos_hor.Value = ""
    End Select
    
    ' Outdoor storage
    If EMMST_tab(ligne, EMMST_type.StockageExterieur) <> 0 And Not IsNull(EMMST_tab(ligne, EMMST_type.StockageExterieur)) Then
        
        r.out.Value = "X"
        
        If EMMST_tab(ligne, EMMST_type.StockageSousAbri) = -1 Then
            r.out_shelter.Value = "X"
        Else
            r.out_air.Value = "X"
        End If
        
        If EMMST_tab(ligne, EMMST_type.StockageHorsPoussiere) = -1 Then
            r.out_dust.Value = "X"
        End If
        
        If EMMST_tab(ligne, EMMST_type.StockageHorsEau) = -1 Then
            r.out_water.Value = "X"
        End If
    End If
    
    ' Indoor storage
    If EMMST_tab(ligne, EMMST_type.StockageInterieur) <> 0 And Not IsNull(EMMST_tab(ligne, EMMST_type.StockageInterieur)) Then
        
        r.ind.Value = "X"
        
        If EMMST_tab(ligne, EMMST_type.StockageChaud) = -1 Then
            r.ind_hot.Value = "X"
        End If
        
        If EMMST_tab(ligne, EMMST_type.StockageFroid) = -1 Then
            r.ind_cold.Value = "X"
        End If
        
        If EMMST_tab(ligne, EMMST_type.StockageHorsPoussiere) = -1 Then
            r.ind_dust.Value = "X"
        End If
        
        If EMMST_tab(ligne, EMMST_type.StockageHorsEau) = -1 Then
            r.ind_water.Value = "X"
        End If
        
        If EMMST_tab(ligne, EMMST_type.StockageHorsLumiere) = -1 Then
            r.ind_light.Value = "X"
        End If
    End If
    
    ' Stackable
    If EMMST_tab(ligne, EMMST_type.Gerbable) = 0 Or IsNull(EMMST_tab(ligne, EMMST_type.Gerbable)) Then
        r.stackable_no.Value = "X"
    ElseIf EMMST_tab(ligne, EMMST_type.Gerbable) = -1 Then
        r.stackable_yes.Value = "X"
    End If
    
    ' Storage mean
    If EMMST_tab(ligne, EMMST_type.MoyenStockage) = 0 Then
        r.stor_mean_no.Value = "X"
    ElseIf EMMST_tab(ligne, EMMST_type.MoyenStockage) = -1 Then
        r.stor_mean_yes.Value = "X"
    Else
        r.stor_mean_no.Value = ""
    End If
    
    ' Storage precautions
    If EMMST_tab(ligne, EMMST_type.PrecautionStockage) = 0 Then
        r.stor_prec_no.Value = "X"
    ElseIf EMMST_tab(ligne, EMMST_type.PrecautionStockage) = -1 Then
        r.stor_prec_yes.Value = "X"
    Else
        r.stor_prec_no.Value = ""
    End If
    
    ' Maintenance during storage
    If EMMST_tab(ligne, EMMST_type.MaintenanceStockage) = 0 Then
        r.stor_maint_no.Value = "X"
    ElseIf EMMST_tab(ligne, EMMST_type.MaintenanceStockage) = -1 Then
        r.stor_maint_yes.Value = "X"
    Else
        r.stor_maint_no.Value = ""
    End If
    
End Sub

' =========================================================================
' PROCEDURE: REMPLIR CHAMPS TRANSPORT
' =========================================================================
Public Sub RemplirChampsTransport(ByRef r As EMMST_Ranges, ByVal ligne As Integer)
' Remplit champs transport
    
    ' Transport mean
    If EMMST_tab(ligne, EMMST_type.MoyenTransport) = 0 Then
        r.transp_mean_no.Value = "X"
    ElseIf EMMST_tab(ligne, EMMST_type.MoyenTransport) = -1 Then
        r.transp_mean_yes.Value = "X"
    Else
        r.transp_mean_no.Value = ""
    End If
    
    ' Specific tool
    If EMMST_tab(ligne, EMMST_type.OutillageSpeTransport) = 0 Then
        r.transp_tool_no.Value = "X"
    ElseIf EMMST_tab(ligne, EMMST_type.OutillageSpeTransport) = -1 Then
        r.transp_tool_yes.Value = "X"
    Else
        r.transp_tool_no.Value = ""
    End If
    
    ' Specific requirement
    If EMMST_tab(ligne, EMMST_type.PrescriptionTransport) = 0 Then
        r.transp_req_no.Value = "X"
    ElseIf EMMST_tab(ligne, EMMST_type.PrescriptionTransport) = -1 Then
        r.transp_req_yes.Value = "X"
    Else
        r.transp_req_no.Value = ""
    End If
    
    ' Precaution
    If EMMST_tab(ligne, EMMST_type.PrecautionTransport) = 0 Then
        r.transp_prec_no.Value = "X"
    ElseIf EMMST_tab(ligne, EMMST_type.PrecautionTransport) = -1 Then
        r.transp_prec_yes.Value = "X"
    Else
        r.transp_prec_no.Value = ""
    End If
    
End Sub

' =========================================================================
' PROCEDURE: REMPLIR COMMENTAIRES
' =========================================================================
Public Sub RemplirCommentaires(ByRef r As EMMST_Ranges, ByVal ligne As Integer)
' Remplit les commentaires selon la langue (FR ou EN)
    
    If langue = "EN" Then
        ' Commentaires anglais
        r.pack_com.Value = EMMST_tab(ligne, EMMST_type.CommEmballage_EN)
        r.lab_com.Value = EMMST_tab(ligne, EMMST_type.CommMarquage_EN)
        r.man_com.Value = EMMST_tab(ligne, EMMST_type.CommManutention_EN)
        r.stor_com.Value = EMMST_tab(ligne, EMMST_type.CommStockage_EN)
        r.transp_com.Value = EMMST_tab(ligne, EMMST_type.CommTransport_EN)
    Else
        ' Commentaires francais
        r.pack_com.Value = EMMST_tab(ligne, EMMST_type.CommEmballage)
        r.lab_com.Value = EMMST_tab(ligne, EMMST_type.CommMarquage)
        r.man_com.Value = EMMST_tab(ligne, EMMST_type.CommManutention)
        r.stor_com.Value = EMMST_tab(ligne, EMMST_type.CommStockage)
        r.transp_com.Value = EMMST_tab(ligne, EMMST_type.CommTransport)
    End If
    
End Sub

' =========================================================================
' PROCEDURE: MARQUER STATUT INDEX EMMST
' =========================================================================
Public Sub MarquerStatutIndexEMMST(ByRef index As Worksheet, ByVal indexDTR As Integer, ByVal statut As String)
' Marque le statut dans l'Index (A creer, A remplir)
    
    If langue = "FR" Then
        index.Cells(indexDTR + 5, 6) = statut
    Else
        ' Traduction anglaise
        Select Case statut
            Case "A creer": index.Cells(indexDTR + 5, 6) = "To create"
            Case "A remplir": index.Cells(indexDTR + 5, 6) = "To fill"
            Case Else: index.Cells(indexDTR + 5, 6) = statut
        End Select
    End If
    
End Sub

' =========================================================================
' PROCEDURE: CREER LIEN HYPERTEXTE EMMST
' =========================================================================
Public Sub CreerLienHypertexteEMMST(ByRef index As Worksheet, ByVal indexDTR As Integer, ByRef dest As Worksheet)
' Cree le lien hypertexte depuis l'Index vers la feuille
    
    Dim colonneNom As String
    
    If langue = "FR" Then
        colonneNom = "Lien Hypertexte"
    Else
        colonneNom = "Hyperlink"
    End If
    
    With index
        .Hyperlinks.Add Anchor:=.Range("Tableau1").ListObject.ListColumns(colonneNom).DataBodyRange(indexDTR), Address:="", SubAddress:="'" & dest.Name & "'!A1", TextToDisplay:=dest.Name
    End With
    
End Sub

' =========================================================================
' PROCEDURE: REMPLIR COUVERTURE EMMST
' =========================================================================
Public Sub RemplirCouvertureEMMST(ByRef CoverPage As Worksheet)
' Remplit la page de couverture avec les infos projet
    
    CoverPage.Range("C13") = Code_Projet
    CoverPage.Range("B17") = Revision
    CoverPage.Range("C17") = Auteur
    CoverPage.Range("F11") = Nv_Maint_Projet
    
End Sub

' =========================================================================
' PROCEDURE: FINALISER FICHIER EMMST
' =========================================================================
Public Sub FinaliserFichierEMMST(ByRef Template As Workbook)
' Finalise le fichier EMMST : Index, couverture, suppression template
    
    Dim index As Worksheet
    Dim TabIndex As ListObject
    Dim Template_EMMST As Worksheet
    
    ' Identifier feuilles selon langue
    If langue = "FR" Then
        Set index = Template.Worksheets("Sommaire")
        Set Template_EMMST = Template.Worksheets("EMMST_VBN_FR")
    Else
        Set index = Template.Worksheets("Index")
        Set Template_EMMST = Template.Worksheets("PHLST")
    End If
    
    ' Renommer colonne Index
    Set TabIndex = index.Range("Tableau1").ListObject
    If langue = "FR" Then
        TabIndex.ListColumns(5).Name = "A creer/remplir?"
        index.Range("B3") = "EMMST Projet " & Projet_Abrev & " Rev " & Revision
    Else
        TabIndex.ListColumns(5).Name = "To create/fill?"
        index.Range("B3") = "PHLST Project " & Projet_Abrev & " Rev " & Revision
    End If
    
    ' Supprimer feuille template
    Application.DisplayAlerts = False
    Template_EMMST.Delete
    Application.DisplayAlerts = True
    
    ' Fermer connexion BDD
    Close_CONN (Replace(CONNEXION_EXCEL_BDD_TRAVAIL, "UserProfile", userpath))
    
End Sub


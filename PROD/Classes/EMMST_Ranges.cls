VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EMMST_Ranges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =========================================================================
' Classe : EMMST_Ranges
' Regroupe les ranges de destination pour la fiche EMMST
' =========================================================================
Option Explicit
' --- General ---
Public spare As Range
Public Tool As Range
Public designation As Range
Public dtr As Range
Public supplier_id As Range
Public supplier_name As Range
' --- Bare Dimensions ---
Public blength As Range
Public bwidth As Range
Public bheight As Range
Public bmass As Range
' --- Packaging Dimensions ---
Public plength As Range
Public pwidth As Range
Public pheight As Range
Public pmass As Range
' --- Lapsing / Packaging type ---
Public lapsing_no As Range
Public lapsing_yes As Range
Public expiry_date As Range
Public multiple_packaging As Range
' --- Reusable / Labelling ---
Public non_reusable_box As Range
Public reusable_box As Range
Public pack_label_no As Range
Public pack_label_yes As Range
Public equip_label_no As Range
Public equip_label_yes As Range
' --- Handling ---
Public handling_mean_no As Range
Public handling_mean_yes As Range
Public spe_tool_no As Range
Public spe_tool_yes As Range
Public man_prec_no As Range
Public man_prec_yes As Range
' --- Storage Position ---
Public pos_hor As Range
Public pos_ver As Range
Public pos_indif As Range
' --- Storage Outdoor ---
Public out As Range
Public out_shelter As Range
Public out_air As Range
Public out_dust As Range
Public out_water As Range
' --- Storage Indoor ---
Public ind As Range
Public ind_hot As Range
Public ind_cold As Range
Public ind_dust As Range
Public ind_water As Range
Public ind_light As Range
' --- Stackable / Storage means ---
Public stackable_no As Range
Public stackable_yes As Range
Public stor_mean_no As Range
Public stor_mean_yes As Range
Public stor_prec_no As Range
Public stor_prec_yes As Range
Public stor_maint_no As Range
Public stor_maint_yes As Range
' --- Transport ---
Public transp_mean_no As Range
Public transp_mean_yes As Range
Public transp_tool_no As Range
Public transp_tool_yes As Range
Public transp_req_no As Range
Public transp_req_yes As Range
Public transp_prec_no As Range
Public transp_prec_yes As Range
' --- Comments ---
Public pack_com As Range
Public lab_com As Range
Public man_com As Range
Public stor_com As Range
Public transp_com As Range
' =========================================================================
' Méthode d'initialisation : remplace ton InitRangeEMMST
' =========================================================================
Public Sub Init(ByRef dEst As Worksheet)
   With dEst
       Set spare = .Range("E4")
       Set Tool = .Range("G4")
       Set designation = .Range("D6")
       Set dtr = .Range("H6")
       Set supplier_id = .Range("D8")
       Set supplier_name = .Range("H8")
       Set blength = .Range("F10")
       Set bwidth = .Range("H10")
       Set bheight = .Range("J10")
       Set bmass = .Range("L10")
       Set plength = .Range("F12")
       Set pwidth = .Range("H12")
       Set pheight = .Range("J12")
       Set pmass = .Range("L12")
       Set lapsing_no = .Range("F14")
       Set lapsing_yes = .Range("H14")
       Set expiry_date = .Range("J14")
       Set multiple_packaging = .Range("E16")
       Set non_reusable_box = .Range("F18")
       Set reusable_box = .Range("H18")
       Set pack_label_no = .Range("E22")
       Set pack_label_yes = .Range("F22")
       Set equip_label_no = .Range("E24")
       Set equip_label_yes = .Range("F24")
       Set handling_mean_no = .Range("E28")
       Set handling_mean_yes = .Range("F28")
       Set spe_tool_no = .Range("E30")
       Set spe_tool_yes = .Range("F30")
       Set man_prec_no = .Range("E32")
       Set man_prec_yes = .Range("F32")
       Set pos_hor = .Range("F34")
       Set pos_ver = .Range("H34")
       Set pos_indif = .Range("J34")
       Set out = .Range("E36")
       Set out_shelter = .Range("G36")
       Set out_air = .Range("I36")
       Set out_dust = .Range("K36")
       Set out_water = .Range("M36")
       Set ind = .Range("E38")
       Set ind_hot = .Range("G38")
       Set ind_cold = .Range("I38")
       Set ind_dust = .Range("K38")
       Set ind_water = .Range("M38")
       Set ind_light = .Range("O38")
       Set stackable_no = .Range("E42")
       Set stackable_yes = .Range("F42")
       Set stor_mean_no = .Range("E44")
       Set stor_mean_yes = .Range("F44")
       Set stor_prec_no = .Range("E46")
       Set stor_prec_yes = .Range("F46")
       Set stor_maint_no = .Range("E48")
       Set stor_maint_yes = .Range("F48")
       Set transp_mean_no = .Range("E52")
       Set transp_mean_yes = .Range("F52")
       Set transp_tool_no = .Range("E54")
       Set transp_tool_yes = .Range("F54")
       Set transp_req_no = .Range("E56")
       Set transp_req_yes = .Range("F56")
       Set transp_prec_no = .Range("E58")
       Set transp_prec_yes = .Range("F58")
       Set pack_com = .Range("I15")
       Set lab_com = .Range("G20")
       Set man_com = .Range("G26")
       Set stor_com = .Range("G40")
       Set transp_com = .Range("G50")
   End With
End Sub

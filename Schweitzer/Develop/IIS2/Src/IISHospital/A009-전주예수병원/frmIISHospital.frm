VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIISHospital 
   Caption         =   "IISHospital"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin MSComctlLib.ImageList imlHospital 
      Left            =   60
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   84
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":0000
            Key             =   "Uriscan Pro-1"
            Object.Tag             =   "Uriscan-1,Uriscan Pro-1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":0E52
            Key             =   "Uriscan Pro-2"
            Object.Tag             =   "Uriscan-2,Uriscan Pro-2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":1CA4
            Key             =   "Dimension RXL"
            Object.Tag             =   "D-RXL,Dimension RXL"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":2876
            Key             =   "RapidLab 865"
            Object.Tag             =   "R-865,RapidLab 865"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":3AF8
            Key             =   "Stks-1"
            Object.Tag             =   "Stks-1,Stks-1"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":46CA
            Key             =   "Stks-2"
            Object.Tag             =   "Stks-2,Stks-2"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":529C
            Key             =   "Hitachi 7600"
            Object.Tag             =   "H7600,Hitachi 7600"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":62EE
            Key             =   "Axsym"
            Object.Tag             =   "Axsym,Axsym"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":7140
            Key             =   "SE-9000"
            Object.Tag             =   "SE-9000,SE-9000"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":7D12
            Key             =   "Variant II"
            Object.Tag             =   "Variant II,Variant II"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":85EC
            Key             =   "CX3 Delta"
            Object.Tag             =   "CX3 Delta, CX3 Delta"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":943E
            Key             =   "LPIA-NV7"
            Object.Tag             =   "LPIA,LPIA-NV7"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":A490
            Key             =   "Thrombolyzer compact"
            Object.Tag             =   "Compact,Thrombolyzer compact"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":B4E2
            Key             =   "Vitek"
            Object.Tag             =   "Vitek,Vitek"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":C334
            Key             =   "RapidLab 850"
            Object.Tag             =   "R-850,RapidLab 850"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":CF06
            Key             =   "RapidLab 860"
            Object.Tag             =   "R-860,RapidLab 860"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":DAD8
            Key             =   "Vitek II"
            Object.Tag             =   "Vitek II,Vitek II"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":E92A
            Key             =   "Thrombolyzer RackRotor"
            Object.Tag             =   "RackRotor,Thrombolyzer RackRotor"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":F239
            Key             =   "XT-1800i"
            Object.Tag             =   "1800i,XT-1800i"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":FB13
            Key             =   "CA-500"
            Object.Tag             =   "CA500,CA-500"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":107ED
            Key             =   "CA-1500"
            Object.Tag             =   "CA1500,CA-1500"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":1163F
            Key             =   "CA-1500ER"
            Object.Tag             =   "CA1500ER,CA-1500ER"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":11F19
            Key             =   "D-RXLM"
            Object.Tag             =   "D-RXLM,D-RXLM"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":12BF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":1404D
            Key             =   "Architect"
            Object.Tag             =   "Architect,Architect"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":14927
            Key             =   "ESR"
            Object.Tag             =   "ESR,ESR"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":15201
            Key             =   "PCX1"
            Object.Tag             =   "PCX1,PCX1"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":15ADB
            Key             =   "Gem3000"
            Object.Tag             =   "Gem3000,Gem3000"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":16F35
            Key             =   "D-RXLM2"
            Object.Tag             =   "D-RXLM2,D-RXLM2"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":1838F
            Key             =   "XE2100"
            Object.Tag             =   "XE2100,XE2100"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":197E9
            Key             =   "ChorusTRIO"
            Object.Tag             =   "ChorusTRIO,ChorusTRIO"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":1A4C3
            Key             =   "NS1000"
            Object.Tag             =   "NS1000,NS1000"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":1B19D
            Key             =   "RapidLab348-1"
            Object.Tag             =   "RapidLab348-1,RapidLab348-1"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":1BE77
            Key             =   "RapidLab348-2"
            Object.Tag             =   "RapidLab348-2,RapidLab348-2"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":1CB51
            Key             =   "RapidLab348-3"
            Object.Tag             =   "RapidLab348-3,RapidLab348-3"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":1D82B
            Key             =   "RapidLab348-4"
            Object.Tag             =   "RapidLab348-4,RapidLab348-4"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":1E505
            Key             =   "Hitachi7600-1"
            Object.Tag             =   "Hitachi7600-1,Hitachi7600-1"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":1F1DF
            Key             =   "Hitachi7180"
            Object.Tag             =   "Hitachi7180,Hitachi7180"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":1FEB9
            Key             =   "Hitachi7180-1"
            Object.Tag             =   "Hitachi7180-1,Hitachi7180-1"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":20B93
            Key             =   "MultiReader"
            Object.Tag             =   "MultiReader,MultiReader"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":2186D
            Key             =   "IT1000"
            Object.Tag             =   "IT1000,IT1000"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":22547
            Key             =   "Gemini"
            Object.Tag             =   "Gemini,Gemini"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":23221
            Key             =   "Vitek 3"
            Object.Tag             =   "Vitek 3,Vitek 3"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":23673
            Key             =   "Vitek 4"
            Object.Tag             =   "Vitek 4,Vitek 4"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":23AC5
            Key             =   "XT-2000i"
            Object.Tag             =   "XT-2000i,XT-2000i"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":23F17
            Key             =   "Coa1"
            Object.Tag             =   "Coa1"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":25371
            Key             =   "CoaER"
            Object.Tag             =   "CoaER"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":267CB
            Key             =   "UF1000i"
            Object.Tag             =   "UF1000i"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":27C25
            Key             =   "CobasE411"
            Object.Tag             =   "CobasE411,CobasE411"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":2907F
            Key             =   "CobasE601"
            Object.Tag             =   "CobasE601,CobasE601"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":2A4D9
            Key             =   "GEM3500_1"
            Object.Tag             =   "GEM3500_1,GEM3500_1"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":2B933
            Key             =   "GEM3500_2"
            Object.Tag             =   "GEM3500_2,GEM3500_2"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":2CD8D
            Key             =   "Taqman"
            Object.Tag             =   "Taqman,Taqman"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":2E1E7
            Key             =   "Triage"
            Object.Tag             =   "Triage,Triage"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":2F641
            Key             =   "Sofia"
            Object.Tag             =   "Sofia,Sofia"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":30A9B
            Key             =   "Sofia2"
            Object.Tag             =   "Sofia2,Sofia2"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":31EF5
            Key             =   "GeneXpert"
            Object.Tag             =   "GeneXpert,GeneXpert"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":3334F
            Key             =   "B4C"
            Object.Tag             =   "B4C,B4C"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":347A9
            Key             =   "GreenCross"
            Object.Tag             =   "GreenCross,GreenCross"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":355FB
            Key             =   "NSPRIME"
            Object.Tag             =   "NSPRIME,NSPRIME"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":36A55
            Key             =   "Hitachi7600-2"
            Object.Tag             =   "Hitachi7600-2,Hitachi7600-2"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":37EAF
            Key             =   "CobasC513"
            Object.Tag             =   "CobasC513,CobasC513"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":39309
            Key             =   "Sofia3"
            Object.Tag             =   "Sofia3,Sofia3"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":3A763
            Key             =   "MEDIWISS"
            Object.Tag             =   "MEDIWISS,MEDIWISS"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":3BBBD
            Key             =   "CFX96"
            Object.Tag             =   "CFX96,CFX96"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":3D017
            Key             =   "IH500"
            Object.Tag             =   "IH500,IH500"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":3E471
            Key             =   "Uriscan Pro-3"
            Object.Tag             =   "Uriscan Pro-3,Uriscan Pro-3"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":3F8CB
            Key             =   "AlloStation"
            Object.Tag             =   "AlloStation,AlloStation"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":40D25
            Key             =   "CFX96_RV16"
            Object.Tag             =   "CFX96_RV16,CFX96_RV16"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":4217F
            Key             =   "EPOC_1"
            Object.Tag             =   "EPOC_1,EPOC_1"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":435D9
            Key             =   "EPOC_2"
            Object.Tag             =   "EPOC_2,EPOC_2"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":44A33
            Key             =   "LIAISON"
            Object.Tag             =   "LIAISON,LIAISON"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":45E8D
            Key             =   "CobasE411_Disk"
            Object.Tag             =   "CobasE411_Disk,CobasE411_Disk"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":472E7
            Key             =   "MGIT960"
            Object.Tag             =   "MGIT960,MGIT960"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":48741
            Key             =   "UF5000"
            Object.Tag             =   "UF5000,UF5000"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":49B9B
            Key             =   "GEM5000"
            Object.Tag             =   "GEM5000,GEM5000"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":4AFF5
            Key             =   "GEM5000_2"
            Object.Tag             =   "GEM5000_2,GEM5000_2"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":4C44F
            Key             =   "OCSENSORM"
            Object.Tag             =   "OCSENSORM,OCSENSORM"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":4D8A9
            Key             =   "ALINITY"
            Object.Tag             =   "ALINITY,ALINITY"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":4ED03
            Key             =   "GEM5000_OP"
            Object.Tag             =   "GEM5000_OP,GEM5000_OP"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":5015D
            Key             =   "LABOSPECT_1"
            Object.Tag             =   "LABOSPECT_1,LABOSPECT_1"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":50A6C
            Key             =   "LABOSPECT_2"
            Object.Tag             =   "LABOSPECT_2,LABOSPECT_2"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":5137B
            Key             =   "LABOSPECT_ER"
            Object.Tag             =   "LABOSPECT_ER,LABOSPECT_ER"
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIISHospital.frx":51C8A
            Key             =   "PATHFAST"
            Object.Tag             =   "PATHFAST,PATHFAST"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmIISHospital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
'   파일명  : frmIISHospital.frm (전주예수병원)
'   작성자  : 이상대
'   내  용  : 병원별로 사용장비의 아이콘을 관리하는 폼
'   작성일  : 2004-10-05
'   메  모  :
'       1.imlHospital에 이미지 추가시에
'         Key : 해당 장비키 (되도록 전체이름 입력)
'         Tag : 툴바에 표시되는 캡션,메뉴바(툴팁)에 표시되는 캡션
'         예) Key:Hitachi 7600
'             Tag:H7600,Hitachi 7600
'-----------------------------------------------------------------------------'

Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    Set frmIISHospital = Nothing
End Sub



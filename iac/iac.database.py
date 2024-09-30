import os
import subprocess
import pandas as pd

root_measurment = "/Users/piratejet/Documents/IAC.SERVER/Quality/07 Uvolnenie produkcie/2024/"
root_formular = "/Users/piratejet/DOcuments/IAC.SERVER/01 MS/06 Dokumentacia/05 Q (QUA)/Uvolnenie 1ks/"

details_paths = {
"AU736":"123",
"WIP IM ARM VL folie":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger AAL Folie VL,VR,HL,HR"),
   "formular_table":("Audi Q7/F-Q052-08","28"),
},
"WIP IM ARM VR folie":{
   "measurmnet_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger AAL Folie VL,VR,HL,HR"),
   "formular_table":("Audi Q7/F-Q052-08","28"),
},
"WIP IM ARM HL folie":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger AAL Folie VL,VR,HL,HR"),
   "formular_table":("Audi Q7/F-Q052-08","29"),
},
"WIP IM ARM HR folie":{
   "measurment_table": ("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger AAL Folie VL,VR,HL,HR"),
   "formular_table":("Audi Q7/F-Q052-08","29"),
},
"WIP IM BRS VL folie":{
    "measurment_table": ("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Folia VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","16"),
},
"WIP IM BRS VR folie":{
   "measurment_table": ("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Folia VL,VR,HL,HR"),
   "formular_table":("Audi Q7/F-Q052-08","16"),
},
"WIP IM BRS HL folie":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Folia VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","17"),
},
"WIP IM BRS HR folie":{
   "measurment_table": ("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Folia VL,VR,HL,HR"),
   "formular_table":("Audi Q7/F-Q052-08","17"),
},
"WIP IM BRS HL m ROLLO folie":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Folia VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","18"),
},
"WIP IM BRS HR m ROLLO folie":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Folia VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","18"),
},
"WIP IM BRS VL LL leder":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Leder VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","19"),
},
"WIP IM BRS VR LL leder":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Leder VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","19"),
},
"WIP IM BRS VL RL leder {inzert}":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Leder VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","20"),
},
"WIP IM BRS VR RL leder  {inzert}":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Leder VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","20"),
},
"WIP IM BRS HL leder":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Leder VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","21"),
},
"WIP IM BRS HR leder":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Leder VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","21"),
},
"WIP IM ARM VL leder":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger AAL Leder VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","30"),
},
"WIP IM ARM VR leder":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger AAL Leder VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","30"),
},
"WIP IM ARM HL leder":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger AAL Leder VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","31"),
},
"WIP IM ARM HR leder":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger AAL Leder VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","31"),
},
"WIP IM Bruestung L soul  PHEV":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
    "formular_table":("Audi Q7/F-Q052-08","47"),
},
"WIP IM Bruestung R soul  PHEV":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
    "formular_table":("Audi Q7/F-Q052-08","47"),
},
"Bruestung L soul":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
    "formular_table":("Audi Q7/F-Q052-08"),
},
"Bruestung R soul":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
    "formular_table":("Audi Q7/F-Q052-08"),
},
#Farebne
"WIP IM Grundausstattung VL Soul PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","43"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung VL Metropolgrau PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","39"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung VL Titangrau PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","41"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung VR Soul PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","43"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung VR Metropolgrau PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","39"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung VR Titangrau PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","41"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung HL Soul PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","44"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung HL Metropolgrau PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","40"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung HL Titangrau PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","42"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung HR Soul PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","44"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung HR Metropolgrau PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","40"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung HR Titangrau PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","42"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM TUT VL folie":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger TUT Folie VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","22"),
},
"WIP IM TUT VR folie":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger TUT Folie VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","22"),
},
"WIP IM TUT HL folie":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger TUT Folie VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","23"),
},
"WIP IM TUT HR folie":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger TUT Folie VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","23"),
},
"WIP IM ETO HL":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger ETO Folie VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","27"),
},
"WIP IM ETO HR":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger ETO Folie VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","27"),
},
"WIP IM Servicedeckel L":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
    "formular_table":("Audi Q7/F-Q052-08","36"),
},
"WIP IM Armlehne L Soul":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
    "formular_table":("Audi Q7/F-Q052-08","10"),
},
"WIP IM Armlehne R Soul":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
    "formular_table":("Audi Q7/F-Q052-08","10"),
},
"WIP IM Rahmen Fluter L PHEV":{
    "formular_table":("Audi Q7/F-Q052-08","44"),
},
"WIP IM Rahmen Blende L PHEV":{
    "formular_table":("Audi Q7/F-Q052-08","45"),
},
"WIP IM Servicedeckel L PHEV":{
    "formular_table":("Audi Q7/F-Q052-08","46"),
},
"WIP IM Rahmen Fluter L":{
    "formular_table":("Audi Q7/F-Q052-08","32"),
},
"WIP IM Rahmen Fluter R":{
    "formular_table":("Audi Q7/F-Q052-08","33"),
},
"WIP IM Retainer Blende L":{
    "formular_table":("Audi Q7/F-Q052-08","34"),
},
"WIP IM Retainer R":{
    "formular_table":("Audi Q7/F-Q052-08","35"),
},
"WIP IM Retainer Blende PHEV":{
    "formular_table":("Audi Q7/F-Q052-08","37"),
},
"WIP IM Rahmen Serviced.PHEV":{
    "formular_table":("Audi Q7/F-Q052-08","38"),
},
}

'''
details_paths = {

#Touareg
"Touareg3":"123",
"WIP IM AAL TVL Basis 106":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","AAl Folie VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","26")
},
"WIP IM AAL TVR Basis 106":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","AAl Folie VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","26"),
},
"WIP IM AAL THL Basis 108":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","AAl Folie VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","28"),
    },
"WIP IM AAL THR Basis 108":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","AAl Folie VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","28"),
    },
"WIP IM AAL TVL Leder 107":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","AAl Leder VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","27"),
    },
"WIP IM AAL TVR Leder 107":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","AAl Leder VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","27"),
    },
"WIP IM AAL THL Leder 109": {
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","AAl Leder VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","29"),
    },
"WIP IM AAL THR Leder 109":  {
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","AAl Leder VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","29"),
    },
"WIP IM Bruestung THL":   {
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Brustung HL,HR,SSR HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","22"),
    },
"WIP IM Bruestung THR":   {
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Brustung HL,HR,SSR HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","22"),
    },
"WIP IM Bruestung THL SSR":   {
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Brustung HL,HR,SSR HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","23"),
    },
"WIP IM Bruestung THR SSR":   {
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Brustung HL,HR,SSR HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","23"),
    },
#farebne
"Abl. Rueckwand THL Soul":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Ablage Rückwand HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","17"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Ablage Rückwand HL,HR Soul"),
},
"Abl. Rueckwand THR Soul":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Ablage Rückwand HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","17"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Ablage Rückwand HL,HR Soul"),
},
"Abl. Rueckwand THL Raven":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Ablage Rückwand HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","18"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Ablage Rückwand HL,HR Raven"),
},
"Abl. Rueckwand THR Raven":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Ablage Rückwand HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","18"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Ablage Rückwand HL,HR Raven"),
},
"WIP IM Abl.Ruck THL Grigio":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Ablage Rückwand HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","21"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Ablage Rückwand HL,HR Grigio"),
},
"WIP IM Abl.Ruck THR Grigio":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Ablage Rückwand HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","21"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Ablage Rückwand HL,HR Grigio"),
},
"WIP IM Abl.Rueckwand TVL Soul":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Ablage Rückwand VL,VR"),
    "formular_table":("Touareg 3/F-Q003-14","12"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Ablage Rückwand VL,VR Soul"),
},
"WIP IM Abl.Rueckwand TVR Soul":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Ablage Rückwand VL,VR"),
    "formular_table":("Touareg 3/F-Q003-14","12"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Ablage Rückwand VL,VR Soul"),
},
"WIP IM Abl.Rueckwand TVL Raven":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Ablage Rückwand VL,VR"),
    "formular_table":("Touareg 3/F-Q003-14","13"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Ablage Rückwand VL,VR Raven"),
},
"WIP IM Abl.Rueckwand TVR Raven":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Ablage Rückwand VL,VR"),
    "formular_table":("Touareg 3/F-Q003-14","13"),    
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Ablage Rückwand VL,VR Raven"),
},
"WIP IM Abl.Rueckwand TVL Grigi":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Ablage Rückwand VL,VR"),
    "formular_table":("Touareg 3/F-Q003-14","16"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Ablage Rückwand VL,VR Grigio"),
},
"WIP IM Abl.Rueckwand TVR Grigi":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Ablage Rückwand VL,VR"),
    "formular_table":("Touareg 3/F-Q003-14","16"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Ablage Rückwand VL,VR Grigio"),
},
"WIP IM Rahmenteil TVL LL Soul":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Träger Rahmenteil VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","2"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Träger Rahmenteil VL,VR Soul"),
},
"WIP IM Rahmenteil TVR LL Soul":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Träger Rahmenteil VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","2"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Träger Rahmenteil VL,VR Soul"),
},
"WIP IM Rahmenteil TVL LL Raven":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Träger Rahmenteil VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","4"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Träger Rahmenteil VL,VR Raven"),
},
"WIP IM Rahmenteil TVR LL Raven":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Träger Rahmenteil VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","4"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Träger Rahmenteil VL,VR Raven"),
},
"WIP IM Rahmentail TVl LL Grigi":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Träger Rahmenteil VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","10"),
        "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Träger Rahmenteil VL,VR Grigio"),
},
"WIP IM Rahmentail TVR LL Grigi":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Träger Rahmenteil VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","10"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Träger Rahmenteil VL,VR Grigio"),
},

"WIP IM Rahmenteil TVL RL Soul":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Träger Rahmenteil VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","2"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Träger Rahmenteil VL,VR Grigio"),
},
"WIP IM Rahmenteil TVR RL Soul":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Träger Rahmenteil VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","2"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Träger Rahmenteil VL,VR Grigio"),
},
"WIP IM Rahmenteil TVL RL Raven":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Träger Rahmenteil VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","4"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Träger Rahmenteil VL,VR Grigio"),
},
"WIP IM Rahmenteil TVR RL Raven":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Träger Rahmenteil VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","4"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Träger Rahmenteil VL,VR Grigio"),
},
"WIP IM Rahmentail TVR RL Grigi":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Träger Rahmenteil VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","10"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Träger Rahmenteil VL,VR Grigio"),
},
"WIP IM Rahmentail TVR RL Grigi":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Träger Rahmenteil VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","10"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Träger Rahmenteil VL,VR Grigio"),
},

"WIP IM Rahmenteil THL Soul":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Träger Rahmenteil VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","6"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Träger Rahmenteil HL,HR Soul"),
},
"WIP IM Rahmenteil THR Soul":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Träger Rahmenteil VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","6"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Träger Rahmenteil HL,HR Soul"),
},
"WIP IM Rahmenteil THL Raven":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Träger Rahmenteil VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","8"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Träger Rahmenteil HL,HR Raven"),
},
"WIP IM Rahmenteil THR Raven":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Träger Rahmenteil VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","8"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Träger Rahmenteil HL,HR Raven"),
},
"WIP IM Rahmentail THL Grigi":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Träger Rahmenteil VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","11"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Träger Rahmenteil HL,HR Grigio"),
},
"WIP IM Rahmentail THR Grigi":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Träger Rahmenteil VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","11"),
    "color_table":("Touareg 3/Touareg 3 Meranie farebnosti.xlsx", "Träger Rahmenteil HL,HR Grigio"),
},
"WIP IM Crashcliphalt TVL schwa":"none",
"WIP IM Crashcliphalt TVR schwa":"none",
"WIP IM Crashcliphalt THL schwa":"none",
"WIP IM Crashcliphalt THR schwa":"none",
"WIP IM Aufna Ring TVL schwarz":"none",
"WIP IM Aufna Ring TVR schwarz":"none",

"WIP IM Einsteckleiste TVL":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Einsteckleiste VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","24"),
},
"WIP IM Einsteckleiste TVR":{
    "measurmen_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Einsteckleiste VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","24"),
},
"WIP IM Einsteckleiste THL":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Einsteckleiste VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","25"),
},
"WIP IM Einsteckleiste THR":{
    "measurment_table":("Touareg 3/Vysledky meranie Uvolnenie prveho dielu na IM.xlsx","Einsteckleiste VL,VR,HL,HR"),
    "formular_table":("Touareg 3/F-Q003-14","25"),
},


"AU736":"123",
"WIP IM ARM VL folie":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger AAL Folie VL,VR,HL,HR"),
   "formular_table":("Audi Q7/F-Q052-08","28"),
},
"WIP IM ARM VR folie":{
   "measurmnet_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger AAL Folie VL,VR,HL,HR"),
   "formular_table":("Audi Q7/F-Q052-08","28"),
},
"WIP IM ARM HL folie":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger AAL Folie VL,VR,HL,HR"),
   "formular_table":("Audi Q7/F-Q052-08","29"),
},
"WIP IM ARM HR folie":{
   "measurment_table": ("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger AAL Folie VL,VR,HL,HR"),
   "formular_table":("Audi Q7/F-Q052-08","29"),
},
"WIP IM BRS VL folie":{
    "measurment_table": ("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Folia VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","16"),
},
"WIP IM BRS VR folie":{
   "measurment_table": ("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Folia VL,VR,HL,HR"),
   "formular_table":("Audi Q7/F-Q052-08","16"),
},
"WIP IM BRS HL folie":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Folia VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","17"),
},
"WIP IM BRS HR folie":{
   "measurment_table": ("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Folia VL,VR,HL,HR"),
   "formular_table":("Audi Q7/F-Q052-08","17"),
},
"WIP IM BRS HL m ROLLO folie":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Folia VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","18"),
},
"WIP IM BRS HR m ROLLO folie":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Folia VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","18"),
},
"WIP IM BRS VL LL leder":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Leder VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","19"),
},
"WIP IM BRS VR LL leder":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Leder VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","19"),
},
"WIP IM BRS VL RL leder {inzert}":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Leder VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","20"),
},
"WIP IM BRS VR RL leder  {inzert}":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Leder VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","20"),
},
"WIP IM BRS HL leder":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Leder VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","21"),
},
"WIP IM BRS HR leder":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Brustung Leder VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","21"),
},
"WIP IM ARM VL leder":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger AAL Leder VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","30"),
},
"WIP IM ARM VR leder":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger AAL Leder VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","30"),
},
"WIP IM ARM HL leder":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger AAL Leder VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","31"),
},
"WIP IM ARM HR leder":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger AAL Leder VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","31"),
},
"WIP IM Bruestung L soul  PHEV":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
    "formular_table":("Audi Q7/F-Q052-08","47"),
},
"WIP IM Bruestung R soul  PHEV":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
    "formular_table":("Audi Q7/F-Q052-08","47"),
},
"Bruestung L soul":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
    "formular_table":("Audi Q7/F-Q052-08"),
},
"Bruestung R soul":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
    "formular_table":("Audi Q7/F-Q052-08"),
},
#Farebne
"WIP IM Grundausstattung VL Soul PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","43"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung VL Metropolgrau PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","39"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung VL Titangrau PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","41"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung VR Soul PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","43"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung VR Metropolgrau PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","39"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung VR Titangrau PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","41"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung HL Soul PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","44"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung HL Metropolgrau PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","40"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung HL Titangrau PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","42"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung HR Soul PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","44"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung HR Metropolgrau PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","40"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM Grundausstattung HR Titangrau PA":{
   "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
   "formular_table":("Audi Q7/F-Q052-08","42"),
   "color_table":("Audi Q7/AUDI Q7 IM Meranie farebnosti 1.xlsx",)
},
"WIP IM TUT VL folie":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger TUT Folie VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","22"),
},
"WIP IM TUT VR folie":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger TUT Folie VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","22"),
},
"WIP IM TUT HL folie":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger TUT Folie VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","23"),
},
"WIP IM TUT HR folie":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger TUT Folie VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","23"),
},
"WIP IM ETO HL":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger ETO Folie VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","27"),
},
"WIP IM ETO HR":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx","Träger ETO Folie VL,VR,HL,HR"),
    "formular_table":("Audi Q7/F-Q052-08","27"),
},
"WIP IM Servicedeckel L":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
    "formular_table":("Audi Q7/F-Q052-08","36"),
},
"WIP IM Armlehne L Soul":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
    "formular_table":("Audi Q7/F-Q052-08","10"),
},
"WIP IM Armlehne R Soul":{
    "measurment_table":("Audi Q7/Vysledky meranie Uvolnenie prveho dielu na IM.rev1.xlsx",),
    "formular_table":("Audi Q7/F-Q052-08","10"),
},
"WIP IM Rahmen Fluter L PHEV":{
    "formular_table":("Audi Q7/F-Q052-08","44"),
},
"WIP IM Rahmen Blende L PHEV":{
    "formular_table":("Audi Q7/F-Q052-08","45"),
},
"WIP IM Servicedeckel L PHEV":{
    "formular_table":("Audi Q7/F-Q052-08","46"),
},
"WIP IM Rahmen Fluter L":{
    "formular_table":("Audi Q7/F-Q052-08","32"),
},
"WIP IM Rahmen Fluter R":{
    "formular_table":("Audi Q7/F-Q052-08","33"),
},
"WIP IM Retainer Blende L":{
    "formular_table":("Audi Q7/F-Q052-08","34"),
},
"WIP IM Retainer R":{
    "formular_table":("Audi Q7/F-Q052-08","35"),
},
"WIP IM Retainer Blende PHEV":{
    "formular_table":("Audi Q7/F-Q052-08","37"),
},
"WIP IM Rahmen Serviced.PHEV":{
    "formular_table":("Audi Q7/F-Q052-08","38"),
},

"X540 MY21":"123",
"WIP IM Centre Stack Carrier":{
    "measurment_table":("JLR X540 MY 21/IM/IP/JLRx540_MY21_Centre Stack Carrier.xlsx"),
    "formular_table":("JLR X540 MY21/F-Q298-10","7"),
},
"WIP IM Cluster surround CSL LHD":{
    "measurment_table":("JLR X540 MY 21/IM/IP/JLRx540_MY21_Cluster Surroun Lower LHD.xlsx"),
    "formular_table":("JLR X540 MY21"),
},
"WIP IM Cluster surround CSL RHD":{
    "measurment_table":("JLR X540 MY 21/IM/IP/JLRx540_MY21_Cluster Surroun Lower RHD.xlsx"),
    "formular_table":("JLR X540 MY21"),
},
"WIP IM Duct Manifold Flange":{
   "measurment_table":("JLR X540 MY 21/IM/IP/JLRx540_MY21_Duct manifold Flange.xlsx"),
   "formular_table":("JLR X540 MY21/F-Q298-10","3"),
},
"WIP IM Duct Manifold Outer":{
   "measurment_table":("JLR X540 MY 21/IM/IP/JLRx540_MY21_Duct manifold Outer.xlsx"),
   "formular_table":("JLR X540 MY21/F-Q298-10","2"),
},
"WIP IM DSL Manual LHD Jet":{
   "measurment_table":("JLR X540 MY 21/IM/IP"),
   "formular_table":("JLR X540 MY21/F-Q298-10","9"),
},
"WIP IM DSL Manual RHD Jet":{
   "measurment_table":("JLR X540 MY 21/IM/IP"),
   "formular_table":("JLR X540 MY21/F-Q298-10","9"),
},
"WIP IM DSL Manual L Oyster LHD":{
   "measurment_table":("JLR X540 MY 21/IM/IP"),
   "formular_table":("JLR X540 MY21/F-Q298-10","9"),
},
"WIP IM DSL Manual L Oyster RHD":{
   "measurment_table":("JLR X540 MY 21/IM/IP"),
   "formular_table":("JLR X540 MY21/F-Q298-10","9"),
},
"WIP IM DSL Auto LHD Jet":{
   "measurment_table":("JLR X540 MY 21/IM/IP"),
   "formular_table":("JLR X540 MY21/F-Q298-10","9"),
},
"WIP IM DSL Auto RHD Jet":{
   "measurment_table":("JLR X540 MY 21/IM/IP"),
   "formular_table":("JLR X540 MY21/F-Q298-10","9"),
},
"WIP IM DSL Auto L Oyster LHD":{
   "measurment_table":("JLR X540 MY 21/IM/IP"),
   "formular_table":("JLR X540 MY21/F-Q298-10","9"),
},
"WIP IM DSL Auto L Oyster RHD":{
   "measurment_table":("JLR X540 MY 21/IM/IP"),
   "formular_table":("JLR X540 MY21/F-Q298-10","9"),
},
"WIP IM Auto shifter bracket":{
   "measurment_table":("JLR X540 MY 21/IM/IP"),
   "formular_table":("JLR X540 MY21/F-Q298-10","9"),
},
"WIP IM Tray Mat for WCD":{
   "measurment_table":("JLR X540 MY 21/IM/Konzola/JLRx540 MY21_TRY mat WDC, Non wdc.xlsx",),
   "formular_table":("JLR X540 MY21/F-Q297-03","2"),
},
"WIP IM Tray Mat for NON WDC":{
   "measurment_table":("JLR X540 MY 21/IM/Konzola/JLRx540 MY21_TRY mat WDC, Non wdc.xlsx",),
   "formular_table":("JLR X540 MY21/F-Q297-03","3"),
},
"WIP IM ARM CONSOLE SUBSTRAT":{
    "measurment_table":("JLR X540 MY 21/IM/Konzola/JLRx540 MY21_Console Armrest Substrate.xlsx",),
    "formular_table":("JLR X540/F-Q222-02"),
},
"DSL Reinforcement LHD":{
   "measurment_table":("JLRx540_MY21_Drive Side Lower Reinforce LHD.xlsx",),
   "formular_table":("JLR X540 MY21/F-Q298-10","10"),
},
"DSL Reinforcement RHD":{
   "measurment_table":("JLRx540_MY21_Cluster Surroun Lower RHD.xlsx",),
   "formular_table":("JLR X540 MY21/F-Q298-10","11"),
},
"WIP IM Inner LHD":"none",
"WIP IM Inner RHD":"none",
"WIP IM Ambi Light Cradle LHD":{
   "measurment_table":("JLR X540 MY 21/IM/IP/JLRx540_MY21_Ambient Light Cradle.xlsx",),
   "formular_table":("JLR X540 MY21/F-Q298-10","8"),
},
"WIP IM Ambi Light Cradle RHD":{
   "measurment_table":("JLR X540 MY 21/IM/IP/JLRx540_MY21_Ambient Light Cradle.xlsx",),
   "formular_table":("JLR X540 MY21/F-Q298-10","8"),
},
"WIP IM Duct Manifold Baffle":{
   "measurment_table":("JLR X540 MY 21/IM/IP/JLRx540_MY21_Duct manifold Baffle.xlsx",),
   "formular_table":("JLR X540 MY21/F-Q298-10","4"),
},
"REAR STOWEAGE END CAP MAT X540 _MAGNA":"none",
"REAR STOWEAGE END CAP MAT L550 _ HALEWOOD":"none",
"WIP IM ARM IN SUBSTR":{
   "measurment_table":("JLR X540 MY 21/IM/Konzola/JLRx540 MY21_Armrest Inner Substrate.xlsx",),
   "formular_table":("JLR X540/F-Q222-02"),
},

"NEW FORD BX726 MCA":"123",
"WIP IM Top Roll Rei.FR LHS Ebo":{
    "measurment_table":("Ford BX 726 MCA/Uvolnenie 1ks IM.xlsx","Top Roll LH"),
    "formular_table":("Ford BX726 MCA/F-Q384-01","6"),
},
"WIP IM Top Roll Rei.FR RHS Ebo":{
    "measurment_table":("Ford BX 726 MCA/Uvolnenie 1ks IM.xlsx","Top Roll RH"),
    "formular_table":("Ford BX726 MCA/F-Q384-01","6"),
},
"WIP IM Insert FR.LHS Ebo MCA":{
    "measurment_table":("Ford BX 726 MCA/Uvolnenie 1ks IM.xlsx","Insert Panel LH"),
    "formular_table":("Ford BX726 MCA/F-Q384-01","3"),
},
"WIP IM Insert FR.RHS Ebo MCA":{
    "measurment_table":("Ford BX 726 MCA/Uvolnenie 1ks IM.xlsx","Insert Panel RH"),
    "formular_table":("Ford BX726 MCA/F-Q384-01","3"),
},
"WIP IM Main.Car.FR LHS Ebo MCA":{
    "measurment_table":("Ford BX 726 MCA/Uvolnenie 1ks IM.xlsx","Main Carrier LH"),
    "formular_table":("Ford BX726 MCA/F-Q384-01","4"),
},
"WIP IM Main.Car.FR RHS Ebo MCA":{
    "measurment_table":("Ford BX 726 MCA/Uvolnenie 1ks IM.xlsx","Main Carrier RH"),
    "formular_table":("Ford BX726 MCA/F-Q384-01","4"),
},
"WIP IM Grab Han. FR LHS str.MCA":{
    "measurment_table":("Ford BX 726 MCA/Uvolnenie 1ks IM.xlsx","Grab Handle Support LH"),
    "formular_table":("Ford BX726 MCA/F-Q384-01","2"),
},
"WIP IM Grab Han. FR RHS str.MCA":{
    "measurment_table":("Ford BX 726 MCA/Uvolnenie 1ks IM.xlsx","Grab Handle Support RH"),
    "formular_table":("Ford BX726 MCA/F-Q384-01","2"),
},
"WIP IM Leg Pusher.FR.LHS MCA":{
    "measurment_table":("Ford BX 726 MCA/Uvolnenie 1ks IM.xlsx","Leg Pusher LH"),
    "formular_table":("Ford BX726 MCA/F-Q384-01","5"),
},
"WIP IM Leg Pusher.FR.RHS MCA":{
    "measurment_table":("Ford BX 726 MCA/Uvolnenie 1ks IM.xlsx","Leg Pusher RH"),
    "formular_table":("Ford BX726 MCA/F-Q384-01","5"),
},
"Switch Bezel Bolt Cover":{
    "measurment_table":("Ford BX 726 MCA/Uvolnenie 1ks IM.xlsx","Bolt cover New"),
    "formular_table":("Ford BX726 MCA/F-Q384-01","9"),
},

"L462":"123",
"WIP IM LHF Arm rest":{
    "measurment_table":("JLR L462/JLR L462 Meranie na Lehre IM.xlsx","Armrest Front"),
    "formular_table":("JLR 462/F-Q255-03","10"),
},
"WIP IM RHF Arm rest":{
    "measurment_table":("JLR L462/JLR L462 Meranie na Lehre IM.xlsx","Armrest Front"),
    "formular_table":("JLR 462/F-Q255-03","10"),
},
"WIP IM LHR Arm rest":{
    "measurment_table":("JLR L462/JLR L462 Meranie na Lehre IM.xlsx","Armrest Rear"),
    "formular_table":("JLR 462/F-Q255-03","11"),
},
"WIP IM RHR Arm rest":{
    "measurment_table":("JLR L462/JLR L462 Meranie na Lehre IM.xlsx","Armrest Rear"),
    "formular_table":("JLR 462/F-Q255-03","11"),
},
"WIP IM LHR Map pocket Ebony":{
    "measurment_table":("JLR L462/JLR L462 Meranie na Lehre IM.xlsx","Map Pocket Rear"),
    "formular_table":("JLR 462/F-Q255-03","4"),
},
"WIP IM RHR Map pocket Ebony":{
    "measurment_table":("JLR L462/JLR L462 Meranie na Lehre IM.xlsx","Map Pocket Rear"),
    "formular_table":("JLR 462/F-Q255-03","4"),
},
"WIP IM LHF Map pocket Ebony":{
    "measurment_table":("JLR L462/JLR L462 Meranie na Lehre IM.xlsx","Map Pocket Front"),
    "formular_table":("JLR 462/F-Q255-03","2"),
},
"WIP IM RHF Map pocket Ebony":{
    "measurment_table":("JLR L462/JLR L462 Meranie na Lehre IM.xlsx","Map Pocket Front"),
    "formular_table":("JLR 462/F-Q255-03","2"),
},
"WIP IM LHR Main casing Ebony":{
    "measurment_table":("JLR L462/JLR L462 Meranie na Lehre IM.xlsx","Main Casing Rear"),
    "formular_table":("JLR 462/F-Q255-03","8"),
},
"WIP IM RHR Main casing Ebony":{
    "measurment_table":("JLR L462/JLR L462 Meranie na Lehre IM.xlsx","Main Casing Rear"),
    "formular_table":("JLR 462/F-Q255-03","8"),
},
"WIP IM LHF Main casing Ebony":{
    "measurment_table":("JLR L462/JLR L462 Meranie na Lehre IM.xlsx","Main Casing Front"),
    "formular_table":("JLR 462/F-Q255-03","6"),
},
"WIP IM RHF Main casing Ebony":{
    "measurment_table":("JLR L462/JLR L462 Meranie na Lehre IM.xlsx","Main Casing Front"),
    "formular_table":("JLR 462/F-Q255-03","6"),
},

"L663":"33",
"WIP IM LHF Hp.Dr.Ebony":{
    "measurment_table":("JLR L663/L663 Meranie na Lehre IM.xlsx","Hush Panel Driver,Passenger"),
    "formular_table":("JLR L663/F-Q314-01","4"),    
},
"WIP IM RHF Hp.Pas.Ebony":{
    "measurment_table":("JLR L663/L663 Meranie na Lehre IM.xlsx","Hush Panel Driver,Passenger"),
    "formular_table":("JLR L663/F-Q314-01","5"),    
},
"WIP IM LHF Hp.Pas.Ebony":{
    "measurment_table":("JLR L663/L663 Meranie na Lehre IM.xlsx","Hush Panel Driver,Passenger"),
    "formular_table":("JLR L663/F-Q314-01","4"),    
},
"WIP IM RHF Hp.Dr.Ebony":{
    "measurment_table":("JLR L663/L663 Meranie na Lehre IM.xlsx","Hush Panel Driver,Passenger"),
    "formular_table":("JLR L663/F-Q314-01","5"),    
},
"LHF DM Passenger Ebony":"none",
"LHF DM Passenger Lunar":"none",
"LHF DM Driver Ebony":"none",
"LHF DM Driver Lunar":"none",
"LHR DM Ebony":"none",
"LHR DM Lunar":"none",
"RHF DM Passenger Ebony":"none",
"RHF DM Passenger Lunar":"none",
"RHF DM Driver Ebony":"none",
"RHF DM Driver Lunar":"none",
"RHR DM Ebony":"none",
"RHR DM Lunar":"none",
"WIP LH BRB Base Ebony":{
    "measurment_table":("JLR L663/L663 Meranie na Lehre IM.xlsx","Blocker Blonnet Base"),
    "formular_table":("JLR L663/F-Q314-01","3"),    
},
"LH BRB Cover Ebony":{
    "measurment_table":("JLR L663/L663 Meranie na Lehre IM.xlsx","Blocker Blonnet Cap"),
    "formular_table":("JLR L663/F-Q314-01","2"),    
},
"WIP LH BRB Base Lunar":{
    "measurment_table":("JLR L663/L663 Meranie na Lehre IM.xlsx","Blocker Blonnet Base"),
    "formular_table":("JLR L663/F-Q314-01","3"),    
},
"LH BRB Cover Lunar":{
    "measurment_table":("JLR L663/L663 Meranie na Lehre IM.xlsx","Blocker Blonnet Cap"),
    "formular_table":("JLR L663/F-Q314-01","2"),    
},
"WIP RH BRB Base Ebony":{
    "measurment_table":("JLR L663/L663 Meranie na Lehre IM.xlsx","Blocker Blonnet Base"),
    "formular_table":("JLR L663/F-Q314-01","3"),    
},
"RH BRB Cover Ebony":{
    "measurment_table":("JLR L663/L663 Meranie na Lehre IM.xlsx","Blocker Blonnet Cap"),
    "formular_table":("JLR L663/F-Q314-01","2"),    
},
"WIP RH BRB Base Lunar":{
    "measurment_table":("JLR L663/L663 Meranie na Lehre IM.xlsx","Blocker Blonnet Base"),
    "formular_table":("JLR L663/F-Q314-01","3"),    
},
"RH BRB Cover Lunar":{
    "measurment_table":("JLR L663/L663 Meranie na Lehre IM.xlsx","Blocker Blonnet Cap"),
    "formular_table":("JLR L663/F-Q314-01","2"),    
},
"WIP IM CEDS Bracket LH":{
    "measurment_table":("JLR L663/L663 Meranie na Lehre IM.xlsx",),
    "formular_table":("JLR L663/F-Q314-01","3"),    
},
"WIP IM CEDS Bracket RH":{
    "measurment_table":("JLR L663/L663 Meranie na Lehre IM.xlsx",),
    "formular_table":("JLR L663/F-Q314-01","2"),    
},

"L663 Reinf 130":"123",
"LS 130 REINF MLDG LH-REAR":{
    "formular_table":("JLR L663 MY21-22 Kasai/130/F-Q354-04","5"),    
},
"LS 130 REINF MLDG RH-REAR":{
    "formular_table":("JLR L663 MY21-22 Kasai/130/F-Q354-04","5"),    
},

"L663 130":"123",
"TURRET COVER LH EBONY":{
    "formular_table":("JLR L663 MY21-22 Kasai/130/F-Q354-04","9"),    
},
"TURRET COVER RH EBONY":{
    "formular_table":("JLR L663 MY21-22 Kasai/130/F-Q354-04","9"),    
},
"TURRET COVER LH LUNAR":{
    "formular_table":("JLR L663 MY21-22 Kasai/130/F-Q354-04","9"),    
},
"TURRET COVER RH LUNAR":{
    "formular_table":("JLR L663 MY21-22 Kasai/130/F-Q354-04","9"),    
},
"heated switch bezel LHS EBONY":{
    "formular_table":("JLR L663 MY21-22 Kasai/130/F-Q354-04","6"),    
},
"heated switch bezel RHS EBONY":{
    "formular_table":("JLR L663 MY21-22 Kasai/130/F-Q354-04","6"),    
},
"heated switch bezel LHS LUNAR":{
    "formular_table":("JLR L663 MY21-22 Kasai/130/F-Q354-04","6"),    
},
"heated switch bezel RHS LUNAR":{
    "formular_table":("JLR L663 MY21-22 Kasai/130/F-Q354-04","6"),    
},

"L663 Retainer 130":"123",
"SEAT BELT RETAINER EBONY":{
    "formular_table":("JLR L663 MY21-22 Kasai/130/F-Q354-04","8"),    
},
"SEAT BELT RETAINER LUNAR":{
    "formular_table":("JLR L663 MY21-22 Kasai/130/F-Q354-04","8"),    
},

"L663 Armrest":"123",
"WIP IN LS L ARMREST 130":{
    "formular_table":("JLR L663 MY21-22 Kasai/130/F-Q354-04","7"),    
},
"WIP IN LS R ARMREST 130":{
    "formular_table":("JLR L663 MY21-22 Kasai/130/F-Q354-04","7"),    
},

"SK370":"123",
"WIP IM SK370 BRU VORNE LINKS":{
    "measurment_table":("SK 370/Škoda SK 370 - meranie na Lehre.xlsx","Bruestung VL,VR"),
    "formular_table":("Škoda SK370/F-Q254-06","2"), 
},
"WIP IM SK370 BRU VORNE RECHTS":{
    "measurment_table":("SK 370/Škoda SK 370 - meranie na Lehre.xlsx","Bruestung VL,VR"),
    "formular_table":("Škoda SK370/F-Q254-06","2"), 
},
"ZB Regenschirmfach vorne links":{
    "measurment_table":("SK 370/Škoda SK 370 - meranie na Lehre.xlsx",),
    "formular_table":("Škoda SK370/F-Q254-06","3"), 
},
"ZB Regenschirmfach vorne rechts":{
    "measurment_table":("SK 370/Škoda SK 370 - meranie na Lehre.xlsx",),
    "formular_table":("Škoda SK370/F-Q254-06","4"), 
},

"FORD PANDA":"123",
"WIP IM Map Pocket FR LHS Ebony":{
    "measurment_table":("Ford BX 726/DP FRONT components BX726_PC 2024.xlsx","Map Pocket"),
    "formular_table":("Ford BX726/DP/F-Q274-06","2"), 
},
"WIP IM Map Pocket FR RHS Ebony":{
    "measurment_table":("Ford BX 726/DP FRONT components BX726_PC 2024.xlsx","Map Pocket"),
    "formular_table":("Ford BX726/DP/F-Q274-06","2"), 
},
"WIP IM Main Carrier RR. LHS Ebony":{
    "measurment_table":("Ford BX 726/DP REAR components BX726_PC 2024.xlsx","Main carrier"),
    "formular_table":("Ford BX726/DP/F-Q274-06","11"), 
},
"WIP IM Main Carrier RR. RHS Ebony":{
    "measurment_table":("Ford BX 726/DP REAR components BX726_PC 2024.xlsx","Main carrier"),
    "formular_table":("Ford BX726/DP/F-Q274-06","11"), 
},
"WIP IM Pull Cup Close RR.LHS":{
    "measurment_table":("Ford BX 726/DP REAR components BX726_PC 2024.xlsx","Pull Cup Closer"),
    "formular_table":("Ford BX726/DP/F-Q274-06","18"), 
},
"WIP IM Pull Cup Close RR.RHS":{
    "measurment_table":("Ford BX 726/DP REAR components BX726_PC 2024.xlsx","Pull Cup Closer"),
    "formular_table":("Ford BX726/DP/F-Q274-06","18"), 
},
"WIP IM Window Garn.RR.LHS Ebo":{
    "measurment_table":("Ford BX 726/DP REAR components BX726_PC 2024.xlsx","Window garnish"),
    "formular_table":("Ford BX726/DP/F-Q274-06","19"), 
},
"WIP IM Window Garn.RR.RHS Ebo":{
    "measurment_table":("Ford BX 726/DP REAR components BX726_PC 2024.xlsx","Window garnish"),
    "formular_table":("Ford BX726/DP/F-Q274-06","19"), 
},
"WIP IM IRH/TR Reinf.RR.LHS Eb":{
    "measurment_table":("Ford BX 726/DP REAR components BX726_PC 2024.xlsx","IRH Reinforcement"),
    "formular_table":("Ford BX726/DP/F-Q274-06","16"), 
},
"WIP IM IRH/TR Reinf.RR.RHS Eb":{
    "measurment_table":("Ford BX 726/DP REAR components BX726_PC 2024.xlsx","IRH Reinforcement"),
    "formular_table":("Ford BX726/DP/F-Q274-06","16"), 
},
"IP IM DR Doghouse RR.LHS Ebo":{
    "measurment_table":("Ford BX 726/DP REAR components BX726_PC 2024.xlsx","DOGHOUSE"),
    "formular_table":("Ford BX726/DP/F-Q274-06","15"), 
},
"IP IM DR Doghouse RR.RHS Ebo":{
    "measurment_table":("Ford BX 726/DP REAR components BX726_PC 2024.xlsx","DOGHOUSE"),
    "formular_table":("Ford BX726/DP/F-Q274-06","15"), 
},
"WIP IM Armrest. MIC Rear LHS Ebony":{
    "measurment_table":("Ford BX 726/DP REAR components BX726_PC 2024.xlsx","Armrest Substrate MIC"),
    "formular_table":("Ford BX726/DP/F-Q274-06","13"), 
},
"WIP IM Armrest. MIC Rear RHS Ebony":{
    "measurment_table":("Ford BX 726/DP REAR components BX726_PC 2024.xlsx","Armrest Substrate MIC"),
    "formular_table":("Ford BX726/DP/F-Q274-06","13"), 
},

"SK316":"123",
"IP LEISTE ABDECKUNG m.HUD LHD":{
    "measurment_table":("SK 316 Leiste/1ks IM Leiste.xlsx","Leiste LHD"),
    "formular_table":("Škoda SK316/F-Q381-01","1"),
},
"IP LEISTE ABDECKUNG m.HUD RHD":{
    "measurment_table":("SK 316 Leiste/1ks IM Leiste.xlsx","Leiste RHD"),
    "formular_table":("Škoda SK316/F-Q381-01","1"),
},
}
'''
measurment_list = []
formular_list = []
#Функция которая достает путь и лист в виде массива кортежей для каждой детали(по необходимой таблице)
def tables(i,list_name):
    for part, data in details_paths.items():
        # Проверяем, что это словарь с ключом "formular_table"
        if isinstance(data, dict) and i in data:
            # Извлекаем путь к файлу и лист
            table_value = data[i]
            if len(table_value) == 2:
                file_path, sheet = table_value
            else:
                file_path = table_value[0]
                sheet = None
            # Добавляем это в массив
            list_name.append((file_path, sheet))


tables("measurment_table",measurment_list)
tables("formular_table",formular_list)

#full_measurment_list = os.path.join(root_measurment,measurment_list)
#full_formular_list = os.path.join(root_formular,formular_list)

# Функция, которая преобразует деталь в путь до ее папки
for key, value in formular_list[:8]:
    sheet_name = value
    if not key.endswith('.xlsx'):
        # Извлекаем папку и имя файла
        directory = os.path.join(*key.split('/')[:-1])
        key = key.split('/')[-1]
        
        # Формируем текущий путь до директории
        current_formular_path = os.path.join(root_formular, directory)
        
        # Перебираем файлы в директории
        for file_name in os.listdir(current_formular_path):
            if file_name.startswith(key):
                current_formular_path = os.path.join(current_formular_path, file_name)
                break  # Нашли нужный файл, выходим из цикла
        
        # Загружаем данные из Excel
        data = pd.read_excel(current_formular_path, sheet_name=sheet_name)
    else:
        current_formular_path = os.path.join(root_formular, key)
        data = pd.read_excel(current_formular_path, sheet_name=sheet_name)

print(data)

#for i,x in formular_list:
    #print(i)

'''a = "Ford BX726/F-Q274-06"
b = a.split('/')[-1]
print(b)'''

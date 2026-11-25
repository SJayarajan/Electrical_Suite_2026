##**************************************************************************
# ELECTRICAL SUITE VER 1.0 DESIGNED AND PROGRAMMED BY :
# SANTHOSH JAYARAJAN .
# PROGRAMMING START ON 7/6/2023.
# NO EARLIER REVISIONS.
# NOTES:
# THE APP HAS FAULT CALCULATION,CABLE DETAILS QUERY,PROTECTION RELAY QUERY
# AND LIGHTNING PROTECTION RISK ANALYSIS.
# EXCEL FILE GENERATION FOR REPORTS IN FAULT ANALYSIS AND RISK ASSESEMENT
# COMPLETED ON 20/9/2023
# DID ADDITIONS TO THE ELEVATOR TAB: ADDED NUMBER OF STOPS "S",CAR CAPACITY FACTOR AND ACCELERATION.
# **************************************************************************#

#ELEVATOR CALCULATOR DESIGNED AND PROGRAMMED BY :
#SANTHOSH JAYARAJAN .
#PROGRAMMING START ON 16/11/2021.
#NO EARLIER REVISIONS.
#NOTES:
#Checked on 7/12/2021
#USED 3 NESTED LOOPS TO CALCULATE FOR VARIOUS SPEEDS
#WAITING INTERVALS AND PASSENGER CAPACITY
#ADDED A COMBO BOX FOR SELECTION OF PROJECT TYPE

# ADDED A 4.0 M/S FOR LIFT SPEED ON 1/7/2023
# CORRECTED A BUG IN MESSAGE DISPLAY AFTER CALCULATIONS DUE TO WRONG INDENTATION ON 1/7/2023
# ADDED THE NBC CLAUSES FOR RESIDENTIAL TO WRITE IN WORKSHEET ON 3/7/2023
# ON 28/10/2024 DID ADDITIONS TO THE ELEVATOR TAB: ADDED NUMBER OF STOPS "S",CAR CAPACITY FACTOR AND ACCELERATION.
# CORRECTED A BUG IN THE RTT CALCULATION FOR ELEVATOR TAB.
# ON 13/6/2025 - INCREASED THE WAITING INTERVAL MAX VALUE TO 200 FROM 60 SECS
# ON 8/1/2026 CORRECTED A BUG IN THE ELEVATOR PROGRAM THAT CALCULATED RTT WRONGLY.THE LAST VALUE IS 2*tp NOT 2*P*tp
#**************************************************************************#

# IMPORT LIBRARIES
import tkinter
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from math import *
import sys
import xlwt
import math
import numpy
import pandas as pd
import datetime
from datetime import date
from PIL import ImageTk, Image
from io import BytesIO
import os

###VARIABLE ASSIGNMENT#

##########################################################################################################
####################################START OF PROGRAM######################################################
##########################################################################################################
root = Tk()
root.wm_title("Electrical Engineering Suite Ver 1.01 November 2026 SEED Engineering")
root.resizable(0, 0)
root.config(bg="LIGHTGREEN")

# SETTING OF FONT0 FOR EXCEL
Font0 = xlwt.Font()
Font0.name = "Calibri"
Font0.bold = True
Font0.height = 500

# SETTING FONT1
Font1 = xlwt.Font()
Font1.name = "Calibri"
Font1.bold = True
Font1.height = 200

# SETTING OF FONT 2,3 FOR EXCEL)
Font2 = xlwt.Font()
Font2.name = "Calibri"
Font2.height = 180

Font3 = xlwt.Font()
Font3.name = "Calibri"
Font3.height = 400

Font4 = xlwt.Font()
Font4.name = "Calibri"
Font4.bold = True
Font4.height = 200

# EXCEL STYLE SETTINGS
Style0 = xlwt.XFStyle()
pattern = xlwt.Pattern()
pattern.pattern = xlwt.Pattern.SOLID_PATTERN
pattern.pattern_fore_colour = xlwt.Style.colour_map['light_green']
Style0.pattern = pattern
Style0.font = Font0

Style3 = xlwt.XFStyle()
pattern = xlwt.Pattern()
pattern.pattern = xlwt.Pattern.SOLID_PATTERN
pattern.pattern_fore_colour = xlwt.Style.colour_map['orange']
Style3.pattern = pattern
Style3.font = Font2

Style4 = xlwt.XFStyle()
pattern = xlwt.Pattern()
pattern.pattern = xlwt.Pattern.SOLID_PATTERN
pattern.pattern_fore_colour = xlwt.Style.colour_map['white']
Style4.pattern = pattern
Style4.font = Font4
Style4.alignment.wrap = 1
Style4.alignment.horz=xlwt.Alignment.HORZ_CENTER
borders=xlwt.Borders()
Style4.borders.left=1
Style4.borders.right=1
Style4.borders.top=1
Style4.borders.bottom=1

Style5 = xlwt.XFStyle()
pattern = xlwt.Pattern()
pattern.pattern = xlwt.Pattern.SOLID_PATTERN
pattern.pattern_fore_colour = xlwt.Style.colour_map['green']
Style5.pattern = pattern
Style5.font = Font4
Style5.alignment.wrap = 1
Style5.alignment.horz=xlwt.Alignment.HORZ_CENTER

Style6 = xlwt.XFStyle()
pattern = xlwt.Pattern()
pattern.pattern = xlwt.Pattern.SOLID_PATTERN
pattern.pattern_fore_colour = xlwt.Style.colour_map['red']
Style6.pattern = pattern
Style6.font = Font4
Style6.alignment.wrap = 1
Style6.alignment.horz=xlwt.Alignment.HORZ_CENTER

Style1 = xlwt.XFStyle()
Style1.font = Font1

Style2 = xlwt.XFStyle()
Style2.font = Font2

###########ABOUT SUB###################
def About1(Tab1):
    Abot = messagebox.showinfo("About this App..", "Developed in July 2023 by Santhosh Jayarajan (Python Programming) and Anuya Lade for Tech Support...Version -1.0..")
    return
################QUIT SUB################
def Quitt(Tab1):
    sys.exit()

#################################################################################################
##############################SUBS FOR FAULT CALCULATION#########################################
#################################################################################################
def Calc_System_Fault(Tab1):

    ##### CHECK IF ALL DATA AVAILABLE #####
    if(Name_Proj_Entry.get() == ""):
        Err1 = messagebox.showinfo("Data Not Entered..", "Please enter name of Project in the Project Name Box..")
        return

    if (HT_Voltage_Entry.get() == ""):
        Err2 = messagebox.showinfo("Data Not Entered..", "Please enter HT Voltage in the Entry Box in Kilo Volts..")
        return

    if (Fault_Level_HT_Entry.get() == ""):
        Err3 = messagebox.showinfo("Data Not Entered..", "Please enter HT Fault Level in Entry Box in MVA..")
        return

    if (Transformer_Rating_Entry.get() == ""):
        Err4 = messagebox.showinfo("Data Not Entered..", "Please enter Transformer rating in MVA..")
        return

    if (Transformer_Impedence_Entry.get() == ""):
        Err5 = messagebox.showinfo("Data Not Entered..", "Please enter the Transformer impedance in entry box in %..")
        return

    if (HT_BASEKVA_Estimated.get() == ""):
        Err6 = messagebox.showinfo("Data Not Entered..", "Please enter the HT Base MVA -Estimated in MVA")
        return

    if (HT_BASEKV_Estimated.get() == ""):
        Err7 = messagebox.showinfo("Data Not Entered..", "Please Enter the Base HT Voltage in the Entry Box in Kilo Volts..")
        return

    if (LT_BASEKVA_Estimated.get() == ""):
        Err8 = messagebox.showinfo("Data Not Entered..", "Please enter LT Base MVA in the Entry Box in MVA..")
        return

    if (LT_BASEKV_Entry.get() == ""):
        Err9 = messagebox.showinfo("Data Not Entered..", "Please enter LT Base Volts in the Entry Box in Volts..")
        return

    if (Search_HT_Cable_Size == ""):
        Err10 = messagebox.showinfo("Data Not Entered..", "Please Select the HT cable size from the Combo Box For Level -A..")
        return

    if (A_HT_Cable_Runs_Entry.get() == ""):
        Err11 = messagebox.showinfo("Data Not Entered..", "Please enter HT Cable Runs for Section A in Numbers..")
        return

    if (A_HT_Cable_Length_Entry.get() == ""):
        Err12 = messagebox.showinfo("Data Not Entered..", "Please enter HT Cable Length for Section A in Meters..")
        return

##### LEVEL -B #####

    if (Search_HT_Cable_Size_B == ""):
        Err13 = messagebox.showinfo("Data Not Entered..", "Please Select the HT cable size from the Combo Box For Level -B..")
        return

    if (B_Cable_Length_Entry.get() == ""):
        Err14 = messagebox.showinfo("Data Not Entered..", "Please Select the HT cable Length For Level - B..")
        return

    if (B_Cable_Runs_Entry.get() == ""):
        Err15 = messagebox.showinfo("Data Not Entered..", "Please Select the HT cable Runs For Level - B..")
        return

    if (Search_HT_Cable_Size_D == ""):
        Err16 = messagebox.showinfo("Data Not Entered..", "Please Select the LT cable size from the Combo Box For Level -D..")
        return

    if (D_Cable_Runs_Entry.get() == ""):
        Err17 = messagebox.showinfo("Data Not Entered..", "Please Select the LT cable Runs For Level - D..")
        return

    if (D_Cable_Length_Entry.get() == ""):
        Err18 = messagebox.showinfo("Data Not Entered..", "Please Select the LT cable Length For Level - D..")
        return

##### LEVEL E #####
    if ( Search_HT_Cable_Size_E== ""):
        Err19 = messagebox.showinfo("Data Not Entered..", "Please Select the LT cable size from the Combo Box For Level -E..")
        return

    if (E_Cable_Runs_Entry.get() == ""):
        Err20 = messagebox.showinfo("Data Not Entered..", "Please Select the LT cable Runs For Level - E..")
        return

    if (E_Cable_Length_Entry.get() == ""):
        Err21 = messagebox.showinfo("Data Not Entered..", "Please Select the LT cable Length For Level - E..")
        return

##### LEVEL F #####
    if (Search_HT_Cable_Size_F == ""):
        Err22 = messagebox.showinfo("Data Not Entered..",
                                    "Please Select the LT cable size from the Combo Box For Level -F..")
        return

    if (F_Cable_Runs_Entry.get() == ""):
        Err23 = messagebox.showinfo("Data Not Entered..", "Please Select the LT cable Runs For Level - F..")
        return

    if (F_Cable_Length_Entry.get() == ""):
        Err24 = messagebox.showinfo("Data Not Entered..", "Please Select the LT cable Length For Level - F..")
        return


    ###################CALCULATION OF HT FAULT MVA AND CURRENT FOR ALL LEVELS###################################
    Total_Resistance_A = (float(A_HT_Cable_Length_Entry.get()) * float(Cable_Resistance_A)) / (
            1000 * float(A_HT_Cable_Runs_Entry.get()))
    Total_Reactance_A = (float(A_HT_Cable_Length_Entry.get()) * Cable_Reactance_A) / (
            1000 * float(A_HT_Cable_Runs_Entry.get()))
    Z1 = math.sqrt(Total_Resistance_A * Total_Resistance_A + Total_Reactance_A * Total_Reactance_A)
    A_HT_Cable_Impedence_Calculation.config(text=Z1)

    ############ LEVEL- A:CALCULATION OF FAULT CURRENT##########

    Z2 = float(HT_BASEKVA_Estimated.get()) / float(Fault_Level_HT_Entry.get())
    A_U_Reactance_Lbl_Calculation.config(text=Z2)

    #### LEVEL -A TOTAL IMPEDENCE #####
    Z3 = Z1 + Z2
    A_Total_Impedence_Calculation.config(text=Z3)

    Fault_MVA_A = float(HT_BASEKVA_Estimated.get()) / Z3
    A_Fault_MVA_HTCB_Calculation.config(text=Fault_MVA_A)
    Fault_Current_A = Fault_MVA_A / float(HT_BASEKV_Estimated.get())
    A_Fault_Calculation.config(text=Fault_Current_A)
    A_Fault_Level.config(text="{:.2f}".format(Fault_Current_A))

    ########## LEVEL- B:CALCULATION OF FAULT CURRENT ##########
    Total_Resistance_B = (float(B_Cable_Length_Entry.get()) * float(Cable_Resistance_B)) / (
            1000 * float(B_Cable_Runs_Entry.get()))
    Total_Reactance_B = (float(B_Cable_Length_Entry.get()) * Cable_Reactance_B) / (
            1000 * float(B_Cable_Runs_Entry.get()))
    Z4 = math.sqrt(Total_Resistance_B * Total_Resistance_B + Total_Reactance_B * Total_Reactance_B)
    Z5 = (Z4 * float(HT_BASEKVA_Estimated.get())) / (
            float(HT_BASEKV_Estimated.get()) * float(HT_BASEKV_Estimated.get()) * 1000)
    Z6 = Z3 + Z4 + Z5
    B_Total_Impedence_Calculation.config(text=Z6)

    Fault_MVA_B = float(HT_BASEKVA_Estimated.get()) / Z6
    Fault_Current_B = Fault_MVA_B / float(HT_BASEKV_Estimated.get())

    B_Fault_MVA_Calculation.config(text=Fault_MVA_B)
    B_Fault_Current_Calculation.config(text=Fault_Current_B)
    B_Fault_Level.config(text="{:.2f}".format(Fault_Current_B))

    ########## LEVEL- C:CALCULATION OF FAULT CURRENT ##########
    Per_Reactance_BASEKVA = (float(LT_BASEKVA_Estimated.get()) * float(Transformer_Impedence_Entry.get())) / float(
        LT_BASEKVA_Estimated.get())
    Z8 = Z6 + (Per_Reactance_BASEKVA / 100)
    Fault_MVA_C = float(LT_BASEKVA_Estimated.get()) / Z8
    Fault_Current_C = Fault_MVA_C / (1.732 * float(LT_BASEKV_Entry.get()) / 1000)

    C_Total_Impedence_Calculation.config(text=Z8)
    C_Fault_MVA_Calculation.config(text=Fault_MVA_C)
    C_Fault_Current_Calculation.config(text=Fault_Current_C)
    C_Fault_Level.config(text="{:.2f}".format(Fault_Current_C))

    ########### LEVEL -D CALCULATION OF FAULT CURRENT ##########
    Total_Resistance_D = (float(D_Cable_Length_Entry.get()) * float(Cable_Resistance_D)) / (
            1000 * float(D_Cable_Runs_Entry.get()))
    Total_Reactance_D = (float(D_Cable_Length_Entry.get()) * Cable_Reactance_D) / (
            1000 * float(D_Cable_Runs_Entry.get()))

    Z9 = math.sqrt(Total_Resistance_D * Total_Resistance_D + Total_Reactance_D * Total_Reactance_D)

    Z10 = Z9 * float(LT_BASEKVA_Estimated.get()) * 1000 / (
            float(LT_BASEKV_Entry.get()) * float(LT_BASEKV_Entry.get()) / 1000)

    Z11 = Z8 + Z9 + Z10

    D_Total_Impedence_Calculation.config(text=Z11)

    Fault_MVA_D = float(LT_BASEKVA_Estimated.get()) / Z11
    Fault_Current_D = Fault_MVA_D * 1000 / (float(LT_BASEKV_Entry.get()) * 1.732)

    D_Fault_MVA_Calculation.config(text=Fault_MVA_D)
    D_Fault_Current_Calculation.config(text=Fault_Current_D)
    D_Fault_Level.config(text="{:.2f}".format(Fault_Current_D))

    ################LEVEL - E CALCULATIONS ####################
    Total_Resistance_E = (float(E_Cable_Length_Entry.get()) * float(Cable_Resistance_E)) / (
            1000 * float(E_Cable_Runs_Entry.get()))
    Total_Reactance_E = (float(E_Cable_Length_Entry.get()) * Cable_Reactance_E) / (
            1000 * float(E_Cable_Runs_Entry.get()))

    Z12 = math.sqrt(Total_Resistance_E * Total_Resistance_E + Total_Reactance_E * Total_Reactance_E)

    Z13 = Z12 * float(LT_BASEKVA_Estimated.get()) * 1000 / (
            float(LT_BASEKV_Entry.get()) * float(LT_BASEKV_Entry.get()) / 1000)

    Z14 = Z11 + Z12 + Z13

    E_Total_Impedence_Calculation.config(text=Z14)

    Fault_MVA_E = float(LT_BASEKVA_Estimated.get()) / Z14
    Fault_Current_E = Fault_MVA_E * 1000 / (float(LT_BASEKV_Entry.get()) * 1.732)

    E_Fault_MVA_Calculation.config(text=Fault_MVA_E)
    E_Fault_Current_Calculation.config(text=Fault_Current_E)
    E_Fault_Level.config(text="{:.2f}".format(Fault_Current_E))

    ########################### LEVEL -F #################################

    Total_Resistance_F = (float(F_Cable_Length_Entry.get()) * float(Cable_Resistance_F)) / (
            1000 * float(F_Cable_Runs_Entry.get()))
    Total_Reactance_F = (float(F_Cable_Length_Entry.get()) * Cable_Reactance_F) / (
            1000 * float(F_Cable_Runs_Entry.get()))

    Z15 = math.sqrt(Total_Resistance_F * Total_Resistance_F + Total_Reactance_F * Total_Reactance_F)

    Z16 = Z15 * float(LT_BASEKVA_Estimated.get()) * 1000 / (
            float(LT_BASEKV_Entry.get()) * float(LT_BASEKV_Entry.get()) / 1000)

    Z17 = Z14 + Z15 + Z16
    F_Total_Impedence_Calculation.config(text=Z17)

    Fault_MVA_F = float(LT_BASEKVA_Estimated.get()) / Z17
    Fault_Current_F = Fault_MVA_F * 1000 / (float(LT_BASEKV_Entry.get()) * 1.732)

    F_Fault_MVA_Calculation.config(text=Fault_MVA_F)
    F_Fault_Current_Calculation.config(text=Fault_Current_F)
    F_Fault_Level.config(text="{:.2f}".format(Fault_Current_F))

    ############### EXCEL REPORTS #########################
    if (Excel_Report.get() == 1):
        ##### GETTING THE PRESENT DATE AND TIME FOR EXCEL FILE NAME #####
        Current_Time_Excel = datetime.datetime.now()
        Hr = Current_Time.hour
        Min = Current_Time.minute
        Sec = Current_Time.second
        date_today = date.today()

        File_Name_Excel = "Fault_Analysis_" + str(Hr) + "_" + str(Min) + "_" + str(Sec) + ".xls"
        Excel_Workbook = xlwt.Workbook()
        Excel_WorkSheet = Excel_Workbook.add_sheet('Project_' + str(Name_Proj_Entry.get()))
        Excel_WorkSheet.write_merge(0, 1, 0, 6, "Tata Realty and Infrastructure - Fault Level Analysis", Style0)
        Excel_WorkSheet.write(2, 0, "Project:" + str(Name_Proj_Entry.get()), Style1)
        Excel_WorkSheet.write(3, 0, "Date of Report:" + str(date_today), Style1)
        Excel_WorkSheet.write(4, 0, "Time of Report:" + str(Hr) + ":" + str(Min) + ":" + str(Sec), Style1)

        Excel_WorkSheet.write(6, 0, "Sl.No", Style4)
        Excel_WorkSheet.write(6, 1, "Fault Level", Style4)
        Excel_WorkSheet.write(6, 2, "Fault MVA", Style4)
        Excel_WorkSheet.write(6, 3, "Total Impedence at Levels", Style4)
        Excel_WorkSheet.write(6, 4, "Fault Current in KA", Style4)
        Excel_WorkSheet.write(6, 5, "Switchgear Capacity", Style4)

        # WRITE EXCEL DATA
        Excel_WorkSheet.write(7, 0, "1", Style2)
        Excel_WorkSheet.write(8, 0, "2", Style2)
        Excel_WorkSheet.write(9, 0, "3", Style2)
        Excel_WorkSheet.write(10, 0, "4", Style2)
        Excel_WorkSheet.write(11, 0, "5", Style2)
        Excel_WorkSheet.write(12, 0, "6", Style2)

        # LEVEL
        Excel_WorkSheet.write(7, 1, "Level-A", Style2)
        Excel_WorkSheet.write(8, 1, "Level-B", Style2)
        Excel_WorkSheet.write(9, 1, "Level-C", Style2)
        Excel_WorkSheet.write(10, 1, "Level-D", Style2)
        Excel_WorkSheet.write(11, 1, "Level-E", Style2)
        Excel_WorkSheet.write(12, 1, "Level-F", Style2)

        # FAULT MVA
        Excel_WorkSheet.write(7, 2, Fault_MVA_A, Style2)
        Excel_WorkSheet.write(8, 2, Fault_MVA_B, Style2)
        Excel_WorkSheet.write(9, 2, Fault_MVA_C, Style2)
        Excel_WorkSheet.write(10, 2, Fault_MVA_D, Style2)
        Excel_WorkSheet.write(11, 2, Fault_MVA_E, Style2)
        Excel_WorkSheet.write(12, 2, Fault_MVA_F, Style2)

        # TOTAL IMPEDENCE

        Excel_WorkSheet.write(7, 3, Z1, Style2)
        Excel_WorkSheet.write(8, 3, Z6, Style2)
        Excel_WorkSheet.write(9, 3, Z8, Style2)
        Excel_WorkSheet.write(10, 3, Z11, Style2)
        Excel_WorkSheet.write(11, 3, Z14, Style2)
        Excel_WorkSheet.write(12, 3, Z17, Style2)

        # FAULT CURRENT
        Excel_WorkSheet.write(7, 4, Fault_Current_A, Style2)
        Excel_WorkSheet.write(8, 4, Fault_Current_B, Style2)
        Excel_WorkSheet.write(9, 4, Fault_Current_C, Style2)
        Excel_WorkSheet.write(10, 4, Fault_Current_D, Style2)
        Excel_WorkSheet.write(11, 4, Fault_Current_E, Style2)
        Excel_WorkSheet.write(12, 4, Fault_Current_F, Style2)

        ###### INSERT IMAGE ######
        Power_System_Image = Image.open("Images/Power_System_1.png")
        r, g, b, a = Power_System_Image.split()
        Power_System_Image = Image.merge("RGB", (r, g, b))
        Pow = BytesIO()
        Power_System_Image.save(Pow,format='bmp')

        Excel_WorkSheet.insert_bitmap_data(Pow.getvalue(),1,7)

        # SAVE IN EXCEL FILE
        Excel_Workbook.save("Results/" + File_Name_Excel)
        done = messagebox.showinfo("Results Data..", "Results Stored in Excel File :" + str(File_Name_Excel))

#########################################################################################################
###################################CAPACITOR SUBS########################################################
#########################################################################################################
def Calc_KVAR(Tab4):
    ##### ERROR CHECKS #####

    if (Capacitor_KW_Entry.get() == ""):
        Cap_Err1 = messagebox.showinfo("Data Not Entered..", "Please Enter the Total KW for Site")
        return

    if (Capacitor_Present_PF_Entry == ""):
        Cap_Err2 = messagebox.showinfo("Data Not Entered..", "Please Enter the Present Estimated PF in range 0 to 1.0..")
        return

    if (Capacitor_Target_PF_Entry.get() == ""):
        Cap_Err3 = messagebox.showinfo("Data Not Entered..", "Please Enter the Target PF in Range 0 to 1.0..")
        return

    if (Capacitor_Voltage_Entry.get() == ""):
        Cap_Err4 = messagebox.showinfo("Data Not Entered..", "Please Enter the System Voltage in Volts..")
        return

    if float(Capacitor_Present_PF_Entry.get())>1.0 or float(Capacitor_Present_PF_Entry.get())<0.0:
        Cap_Err5 = messagebox.showinfo("Data Not Entered..", "Please Enter the Present PF in Range 0.0 to 1.0..")
        return

    if float(Capacitor_Target_PF_Entry.get())>1.0 or float(Capacitor_Target_PF_Entry.get())<0.0:
        Cap_Err6 = messagebox.showinfo("Data Not Entered..", "Please Enter the Target PF in Range 0.0 to 1.0..")
        return

    if float(Capacitor_Target_PF_Entry.get()) <= float(Capacitor_Present_PF_Entry.get()):
        Cap_Err7 = messagebox.showinfo("Data Not Entered..","Target Power Factor cannot be less than Present Power Factor..")
        return


    Calculated_KVAR=float(Capacitor_KW_Entry.get())*(math.tan(math.acos(float(Capacitor_Present_PF_Entry.get())))-math.tan(math.acos(float(Capacitor_Target_PF_Entry.get()))))
    Capacitor_Result.config(text="{:.2f}".format(Calculated_KVAR))
    Calculated_MFarad=Calculated_KVAR*10**9/(2*math.pi*50*float(Capacitor_Voltage_Entry.get())**2)
    Capacitor_Result_MFarad.config(text="{:.2f}".format(Calculated_MFarad))

    #ORIGINAL PF POINT ALLOCATION
    x1_Ori = 0
    y1_Ori = 400
    x2_Ori = 400
    y2_Ori = 400
    x3_Ori = 400
    y3_Ori = 0

    #POINTS FOR ORIGINAL PF
    Original_PF_points = [x1_Ori, y1_Ori, x2_Ori, y2_Ori, x3_Ori, y3_Ori]

    # TRIANGLE FOR ORIGINAL PF
    Scratch_Board.create_polygon(Original_PF_points, outline="blue",fill="orange", width=2)

   #POINTS FOR TARGET PF
    x1_Tar = 0
    y1_Tar = 400
    x2_Tar = 400
    y2_Tar = 400
    x3_Tar = 400
    y3_Tar = 400*math.tan(math.acos(float(Capacitor_Target_PF_Entry.get())))/math.tan(math.acos(float(Capacitor_Present_PF_Entry.get())))

    # POINTS FOR TARGET PF
    Target_PF_points = [x1_Tar, y1_Tar, x2_Tar, y2_Tar, x3_Tar, y3_Tar]

    # TRIANGLE FOR TARGET PF
    Scratch_Board.create_polygon(Target_PF_points, outline="red", fill="green", width=2)
    Scratch_Board.create_text(200, 175, fill="darkblue", font="Calbri 12 bold", text=" KVA ")
    Scratch_Board.create_text(375, 200, fill="darkblue", font="Calbri 12 bold", text=" KVAr")
    Scratch_Board.create_text(200, 390, fill="darkblue", font="Calbri 12 bold", text=" Kilo Watt")

    #CALCULATED DATA DISPALYAED AS RESULTS

    Scratch_Board.itemconfig(Required_KVAR,text=" Required KVAr: " + str("{:.2f}".format(Calculated_KVAR)))
    Scratch_Board.itemconfig(Present_Phase_Angle, text=" Present Phase Angle(Rad) : " + str("{:.2f}".format(math.acos(float(Capacitor_Present_PF_Entry.get())))))
    Scratch_Board.itemconfig(Target_Phase_Angle,text=" Target Phase Angle(Rad): " + str("{:.2f}".format(math.acos(float(Capacitor_Target_PF_Entry.get())))))

#########################################################################################################
##########################################CABLE QUERY SUBS###############################################
#########################################################################################################
def HT_Cable_Selected_A(Tab1):
    global Cable_Resistance_A
    global Cable_Reactance_A
    Cable_Sel_A = Cable_Main[Cable_Main['Cable Size'] == Search_HT_Cable_Size.get()]
    Cable_Resistance_A = Cable_Sel_A.iat[0, 1]
    Cable_Reactance_A = Cable_Sel_A.iat[0, 2]
    Res_Reac_A.config(text=str(Cable_Resistance_A) + "," + str(Cable_Reactance_A))


##### LEVEL-B HT CABLE SELECTION #####
def HT_Cable_Selected_B(Tab1):
    global Cable_Resistance_B
    global Cable_Reactance_B
    Cable_Sel_B = Cable_Main[Cable_Main['Cable Size'] == Search_HT_Cable_Size_B.get()]
    Cable_Resistance_B = Cable_Sel_B.iat[0, 1]
    Cable_Reactance_B = Cable_Sel_B.iat[0, 2]
    Res_Reac_B.config(text=str(Cable_Resistance_B) + "," + str(Cable_Reactance_B))


def HT_Cable_Selected_D(Tab1):
    global Cable_Resistance_D
    global Cable_Reactance_D
    Cable_Sel_D = Cable_Main[Cable_Main['Cable Size'] == Search_HT_Cable_Size_D.get()]
    Cable_Resistance_D = Cable_Sel_D.iat[0, 1]
    Cable_Reactance_D = Cable_Sel_D.iat[0, 2]
    Res_Reac_D.config(text=str(Cable_Resistance_D) + "," + str(Cable_Reactance_D))


def HT_Cable_Selected_E(Tab1):
    global Cable_Resistance_E
    global Cable_Reactance_E
    Cable_Sel_E = Cable_Main[Cable_Main['Cable Size'] == Search_HT_Cable_Size_E.get()]
    Cable_Resistance_E = Cable_Sel_E.iat[0, 1]
    Cable_Reactance_E = Cable_Sel_E.iat[0, 2]
    Res_Reac_E.config(text=str(Cable_Resistance_E) + "," + str(Cable_Reactance_E))


def HT_Cable_Selected_F(Tab1):
    global Cable_Resistance_F
    global Cable_Reactance_F
    Cable_Sel_F = Cable_Main[Cable_Main['Cable Size'] == Search_HT_Cable_Size_F.get()]
    Cable_Resistance_F = Cable_Sel_F.iat[0, 1]
    Cable_Reactance_F = Cable_Sel_F.iat[0, 2]
    Res_Reac_F.config(text=str(Cable_Resistance_F) + "," + str(Cable_Reactance_F))

##################################################################################################
############################PROTECTION RELAYS QUERY SUBS##########################################
##################################################################################################
def Protect_Selected_Query(Tab5):
    # SELECTING THE CRITERIA
    if Prot_Type_Selection.get()=="Relay Number":
        Protect_Sel = Protect_Main_Query[Protect_Main_Query['Relay Number'] == Search_Protect_Number.get()]

        # POPULATING THE TEXT BOXES
        Protect_Field_1_Data.config(text=Protect_Sel.iat[0, 1])
        Protect_Field_2_Data.config(text=Protect_Sel.iat[0, 2])
        Protect_Field_3_Data.config(text=Protect_Sel.iat[0, 3])

        Protect_Panda = pd.read_csv("Data\\Protection_Relays_31_8_2023.csv")
        Protect_Field_1_Lbl.config(text=Protect_Panda.columns[1])
        Protect_Field_2_Lbl.config(text=Protect_Panda.columns[2])
        Protect_Field_3_Lbl.config(text=Protect_Panda.columns[3])

    # IMAGE FOR PROTECTION RELAY
        Protect_Lbl.config(image=Protect_Relay_Image)

    if Prot_Type_Selection.get()=="Relay Description":
        Protect_Sel = Protect_Main_Query[Protect_Main_Query['Relay Description'] == Search_Protect_Number.get()]

        # POPULATING THE TEXT BOXES
        Protect_Field_1_Data.config(text=Protect_Sel.iat[0, 1])
        Protect_Field_2_Data.config(text=Protect_Sel.iat[0, 2])
        Protect_Field_3_Data.config(text=Protect_Sel.iat[0, 3])

        Protect_Panda = pd.read_csv("Data\\Protection_Relays_31_8_2023.csv")
        Protect_Field_1_Lbl.config(text=Protect_Panda.columns[1])
        Protect_Field_2_Lbl.config(text=Protect_Panda.columns[2])
        Protect_Field_3_Lbl.config(text=Protect_Panda.columns[3])

    # IMAGE FOR PROTECTION RELAY
        Protect_Lbl.config(image=Protect_Relay_Image)

def Prot_Type_Sel():
    # EXTRACTING THE HEADER NAMES WITH PANDAS
    Prot_Panda = pd.read_csv("Data\\Protection_Relays_31_8_2023.csv")
    Protect_Field_1_Lbl.config(text=Prot_Panda.columns[1])
    Protect_Field_2_Lbl.config(text=Prot_Panda.columns[2])
    Protect_Field_3_Lbl.config(text=Prot_Panda.columns[3])


    if Prot_Type_Selection.get()=="Relay Number":
        Protect_Main_Query = pd.read_csv("Data\\Protection_Relays_31_8_2023.csv")
        Prot_Sizes_Query = Protect_Main_Query["Relay Number"]
        Protect_Number_Query_Combo.config(values=Prot_Sizes_Query.tolist())
        Protect_Lbl_Text.config(text='Select the Relay Number')


    if Prot_Type_Selection.get()=="Relay Description":
        Protect_Main_Query = pd.read_csv("Data\\Protection_Relays_31_8_2023.csv")
        Prot_Sizes_Query = Protect_Main_Query["Relay Description"]
        Protect_Number_Query_Combo.config(values=Prot_Sizes_Query.tolist())
        Protect_Lbl_Text.config(text='Select the Protection Relay Type')

#############################################################################################
#########################################CABLE SELECTION QUERY###############################
#############################################################################################

def Cable_Selected_Query(Tab3):

    if Cable_Type_Selection.get() == "HT Cable":
        #SELECTING THE CRITERIA
        Cable_Sel = Cable_Main_Query[Cable_Main_Query['Size'] == Search_Cable_Size.get()]

        #POPULATING THE TEXT BOXES
        Cable_Feild_1_Data.config(text=Cable_Sel.iat[0, 2])
        Cable_Feild_2_Data.config(text=Cable_Sel.iat[0, 3])
        Cable_Feild_3_Data.config(text=Cable_Sel.iat[0, 4])
        Cable_Feild_4_Data.config(text=Cable_Sel.iat[0, 5])
        Cable_Feild_5_Data.config(text=Cable_Sel.iat[0, 6])
        Cable_Feild_6_Data.config(text=Cable_Sel.iat[0, 7])
        Cable_Feild_7_Data.config(text=Cable_Sel.iat[0, 8])
        Cable_Feild_8_Data.config(text=Cable_Sel.iat[0, 9])
        Cable_Feild_9_Data.config(text=Cable_Sel.iat[0, 10])
        #IMAGE FOR HT CABLE
        Cable_HT_Lbl.config(image=Cable_HT_Cable_Image)

    if Cable_Type_Selection.get() == "LT Cable":
        # SELECTING THE CRITERIA
        Cable_Sel_LT = Cable_Main_Query_LT[Cable_Main_Query_LT['Size'] == Search_Cable_Size.get()]

        # POPULATING THE TEXT BOXES
        Cable_Feild_1_Data.config(text=Cable_Sel_LT.iat[0, 2])
        Cable_Feild_2_Data.config(text=Cable_Sel_LT.iat[0, 3])
        Cable_Feild_3_Data.config(text=Cable_Sel_LT.iat[0, 4])
        Cable_Feild_4_Data.config(text=Cable_Sel_LT.iat[0, 5])
        Cable_Feild_5_Data.config(text=Cable_Sel_LT.iat[0, 6])
        Cable_Feild_6_Data.config(text=Cable_Sel_LT.iat[0, 7])
        Cable_Feild_7_Data.config(text=Cable_Sel_LT.iat[0, 8])
        Cable_Feild_8_Data.config(text=Cable_Sel_LT.iat[0, 9])
        Cable_Feild_9_Data.config(text=Cable_Sel_LT.iat[0, 10])
        Cable_Feild_10_Data.config(text=Cable_Sel_LT.iat[0, 11])
        Cable_Feild_11_Data.config(text=Cable_Sel_LT.iat[0, 12])

        # IMAGE FOR HT CABLE
        Cable_HT_Lbl.config(image=Cable_LT_Cable_Image)

def Cable_Type_Sel():
    if Cable_Type_Selection.get()=="HT Cable":
        Cable_Main_Query = pd.read_csv("Data\\HT_Cables_24_8_2023.csv")
        Cable_Sizes_Query = Cable_Main_Query["Size"]
        Cable_Size_Query_Combo.config(values=Cable_Sizes_Query.tolist())

        #EXTRACTING THE HEADER NAMES WITH PANDAS
        HT_Cable_Panda = pd.read_csv("Data\\HT_Cables_24_8_2023.csv")
        Cable_Feild_1_Lbl.config(text=HT_Cable_Panda.columns[2])
        Cable_Feild_2_Lbl.config(text=HT_Cable_Panda.columns[3])
        Cable_Feild_3_Lbl.config(text=HT_Cable_Panda.columns[4])
        Cable_Feild_4_Lbl.config(text=HT_Cable_Panda.columns[5])
        Cable_Feild_5_Lbl.config(text=HT_Cable_Panda.columns[6])
        Cable_Feild_6_Lbl.config(text=HT_Cable_Panda.columns[7])
        Cable_Feild_7_Lbl.config(text=HT_Cable_Panda.columns[8])
        Cable_Feild_8_Lbl.config(text=HT_Cable_Panda.columns[9])
        Cable_Feild_9_Lbl.config(text=HT_Cable_Panda.columns[10])
        Cable_Feild_10_Lbl.config(text="...")
        Cable_Feild_11_Lbl.config(text="...")


    if Cable_Type_Selection.get()=="LT Cable":
        Cable_Main_Query_LT = pd.read_csv("Data\\LT_Cables_30_8_2023.csv")
        Cable_Sizes_Query_LT = Cable_Main_Query_LT["Size"]
        Cable_Size_Query_Combo.config(values=Cable_Sizes_Query_LT.tolist())

        # EXTRACTING THE HEADER NAMES WITH PANDAS
        LT_Cable_Panda = pd.read_csv("Data\\LT_Cables_30_8_2023.csv")
        Cable_Feild_1_Lbl.config(text=LT_Cable_Panda.columns[2])
        Cable_Feild_2_Lbl.config(text=LT_Cable_Panda.columns[3])
        Cable_Feild_3_Lbl.config(text=LT_Cable_Panda.columns[4])
        Cable_Feild_4_Lbl.config(text=LT_Cable_Panda.columns[5])
        Cable_Feild_5_Lbl.config(text=LT_Cable_Panda.columns[6])
        Cable_Feild_6_Lbl.config(text=LT_Cable_Panda.columns[7])
        Cable_Feild_7_Lbl.config(text=LT_Cable_Panda.columns[8])
        Cable_Feild_8_Lbl.config(text=LT_Cable_Panda.columns[9])
        Cable_Feild_9_Lbl.config(text=LT_Cable_Panda.columns[10])
        Cable_Feild_10_Lbl.config(text=LT_Cable_Panda.columns[11])
        Cable_Feild_11_Lbl.config(text=LT_Cable_Panda.columns[12])

###############################################################################
##################TAB-2 LIGHTNING RISK ANALYSIS################################
###############DANGER TO PEOPLE SUB############################################

def Lightning_Danger_People_Sub(Tab2):
    global h
    Lightning_Help_Box.config(text="This parameter measures the presence of people in the building,depending on the occupancy")
    h_index=Lightning_Danger_People_Combo.current()

    if h_index==0:
        h=1
    if h_index==1:
        h=2
    if h_index==2:
        h=5
    if h_index==3:
        h=5
    if h_index==4:
        h=10
    if h_index==5:
        h=20
    if h_index==6:
        h=50

################OCCUPANCY SUB################################
def Lightning_Occupancy_Sub(Tab2):
    global Lf1
    Lightning_Help_Box.config(text="This parameter measures the risk reduction factor with respect to Occupancy")
    h_index = Lightning_Danger_People_Combo.current()
    Lf1_index = Lightning_Occupancy_Combo.current()
    if Lf1_index == 0:
        Lf1 = 0.1
    if Lf1_index == 1:
        Lf1 = 0.01

#################RELATIVE LOCATION SUB Cd#####################
def Lightning_Rel_Loc_Sub(Tab2):
    global Cd
    Lightning_Help_Box.config(text="This parameter measures the risk reduction factor with respect to Location of building.")
    Cd_index = Lightning_Rel_Loc_Combo.current()
    if Cd_index == 0:
        Cd = 0.25
    if Cd_index == 1:
        Cd = 0.5
    if Cd_index == 2:
        Cd = 1
    if Cd_index == 3:
        Cd = 2

#################LOCATION FIRE RISK rf#####################
def Lightning_Fire_Risk_Sub(Tab2):
    global rf
    Lightning_Help_Box.config(text="This parameter measures the risk reduction factor with respect to flammable materials present.")
    rf_index = Lightning_Fire_Risk_Combo.current()
    if rf_index == 0:
        rf = 1
    if rf_index == 1:
        rf = 0.1
    if rf_index == 2:
        rf = 0.01
    if rf_index == 3:
        rf = 0.001

####################ASSSOCIATED SERVICES####################
def Lightning_Ass_Services_Sub(Tab2):
    global Lf2
    Lightning_Help_Box.config(text="This parameter measures the risk reduction factor with respect to Associated Services in Building.")
    Lf2_index = Lightning_Ass_Services_Combo.current()
    if Lf2_index == 0:
        Lf2 = 0
    if Lf2_index == 1:
        Lf2 = 0.1
    if Lf2_index == 2:
        Lf2 = 0.01

###################ELECTRICAL LINES Ai#####################
def Lightning_Electrical_Lines_Sub(Tab2):
    global Ai
    Lightning_Help_Box.config(text="This parameter measures the risk reduction factor with respect to electrical lines.")
    Ai_index = Lightning_Electrical_Lines_Combo.current()
    if Ai_index == 0:
        Ai = 14400
    if Ai_index == 1:
        Ai = 6600
    if Ai_index == 2:
        Ai = 0

###################LIGHTNING PROTECTION LEVELS Pd#####################
def Lightning_Protection_Level_Sub(Tab2):
    global Pd
    Lightning_Help_Box.config(text="This parameter defines the KA rating of the lightning protection system.")
    Pd_index = Lightning_Protection_Level_Combo.current()
    if Pd_index == 0:
        Pd = 1
    if Pd_index == 1:
        Pd = 0.2
    if Pd_index == 2:
        Pd= 0.1
    if Pd_index == 3:
        Pd = 0.05
    if Pd_index == 4:
        Pd = 0.02

###################SURGE PROTECTION DEVICES Pi#####################
def Lightning_Surge_Arrestor_Sub(Tab2):
    global Pi
    Lightning_Help_Box.config(text="This parameter takes into consideration SPD's installed at incomer supply.")
    Pi_index = Lightning_Surge_Arrestor_Combo.current()
    if Pi_index == 0:
        Pi = 1
    if Pi_index == 1:
        Pi = 0.03
    if Pi_index == 2:
        Pi = 0.02
    if Pi_index == 3:
        Pi = 0.01

###################CALCULATE LIGHTNING RISK BUTTON###################
def Calc_Lightning_Risk(Tab2):

    if Lightning_Building_Length_Entry.get() == "":
        Lightning_Err1 = messagebox.showinfo("Data Not Entered..", "Please Enter the Length of the Building in m..")
        return
    if Lightning_Building_Bredth_Entry.get() == "":
        Lightning_Err2 = messagebox.showinfo("Data Not Entered..", "Please Enter the Breadth of the Building in m..")
        return

    if Lightning_Building_Height_Entry.get() == "":
        Lightning_Err3 = messagebox.showinfo("Data Not Entered..", "Please Enter the Height of the Building in m..")
        return

    if Lightning_Strike_Density_Entry.get() == "":
        Lightning_Err4 = messagebox.showinfo("Data Not Entered..","Please Enter the Lightning Density..")
        return

    if Lightning_Danger_People_Combo.get() == "":
        Lightning_Err5 = messagebox.showinfo("Data Not Entered..", "Please Select Danger to People from the ComboBox..")
        return

    if Lightning_Occupancy_Combo.get() == "":
        Lightning_Err6 = messagebox.showinfo("Data Not Entered..", "Please Select Occupancy Coeffecient from the ComboBox..")
        return

    if Lightning_Rel_Loc_Combo.get() == "":
        Lightning_Err7 = messagebox.showinfo("Data Not Entered..", "Please Select Relative Location of the Site from the ComboBox..")
        return

    if Lightning_Fire_Risk_Combo.get() == "":
        Lightning_Err8 = messagebox.showinfo("Data Not Entered..", "Please Select Fire Risk Reduction Factor from the ComboBox..")
        return

    if Lightning_Ass_Services_Combo.get() == "":
        Lightning_Err9 = messagebox.showinfo("Data Not Entered..", "Please Select Associated Services in the Building  from the ComboBox..")
        return

    if Lightning_Electrical_Lines_Combo.get() == "":
        Lightning_Err10 = messagebox.showinfo("Data Not Entered..", "Please Select Presence of Electrical Lines in the site from the ComboBox..")
        return

    if Lightning_Protection_Level_Combo.get() == "":
        Lightning_Err11 = messagebox.showinfo("Data Not Entered..", "Please Select Lightning Protection Level None to IV from the ComboBox..")
        return

    if Lightning_Surge_Arrestor_Combo.get() == "":
        Lightning_Err12 = messagebox.showinfo("Data Not Entered..", "Please Select Details of SPD's in the Incoming Circuit from the ComboBox..")
        return


    #CALCULATING RT1
    #CALCULATING THE LENGTH \,BREADTH AND HEIGHT
    L=float(Lightning_Building_Length_Entry.get())
    W=float(Lightning_Building_Bredth_Entry.get())
    H=float(Lightning_Building_Height_Entry.get())

    #CALCULATING Ad
    Ad=L*W+6*H*(L+W)+math.pi*9*H*H
    Lightning_Building_Ad_Calculated.config(text=Ad)
    Ng=float(Lightning_Strike_Density_Entry.get())
    Rd1=Ng*Ad*Cd*0.000001*Pd*h*rf*Lf2
    Ri2=Ng*Cd*Ai*0.000001*Pi*h*rf*Lf1
    RT1=Rd1+Ri2

    #######CHANGE COLOR OF RESULTS##########
    if RT1>=0.00001:
        Lightning_RT1_Result.config(text=RT1,bg='red',fg='yellow',font="Calbri 12 bold")
        Lightning_Help_Box.config(text='Unsafe ..Lightning Protection Required ',font="Calbri 12 bold")
    else:
        Lightning_RT1_Result.config(text=RT1, bg='green', fg='yellow',font="Calbri 12 bold")
        Lightning_Help_Box.config(text='Safe ..Lightning Protection Not Required ',font="Calbri 12 bold")

    #######CALCULATING RT2###########
    Rd2=Ng*Ad*Cd*0.000001*Pd*rf*Lf2
    Ri2=Ng*Cd*Ai*0.000001*Pi*rf*Lf2
    RT2=Rd2+Ri2

    #######CHANGE COLOR OF RESULTS RT2####
    if RT2 >= 0.001:
        Lightning_RT2_Result.config(text=RT2, bg='red', fg='yellow',font="Calbri 12 bold")

    else:
        Lightning_RT2_Result.config(text=RT2, bg='green', fg='yellow',font="Calbri 12 bold")


     ######CALCULATING RT3#########
    Rd3 = Ng * Ad * Cd * 0.000001 * Pd * rf * 0.1
    Ri3 = Ng * Cd * Ai * 0.000001 * Pi * rf * 0.1
    RT3 = Rd3 + Ri3

    #### CHANGE COLOR OF RESULTS RT3#####
    if RT3 >= 0.001:
        Lightning_RT3_Result.config(text=RT3, bg='red', fg='yellow',font="Calbri 12 bold")

    else:
         Lightning_RT3_Result.config(text=RT3, bg='green', fg='yellow',font="Calbri 12 bold")

#######SAVING IN THE EXCEL FILE ########################

######GETTING THE PRESENT DATE AND TIME FOR EXCEL FILE NAME#####
    Current_Time_Excel = datetime.datetime.now()
    Hr = Current_Time.hour
    Min = Current_Time.minute
    Sec = Current_Time.second
    date_today = date.today()

    File_Name_Excel_Lightning = "Lightning_Protection_" + str(Hr) + "_" + str(Min) + "_" + str(Sec) + ".xls"
    Excel_Workbook = xlwt.Workbook()
    Excel_WorkSheet = Excel_Workbook.add_sheet('Project_' + str(Lightning_Proj_Data_Entry.get()))
    Excel_WorkSheet.write_merge(0, 1, 0, 6, "Tata Realty and Infrastructure - Lightning Risk Level Analysis", Style0)
    Excel_WorkSheet.write(2, 0, "Project:" + str(Lightning_Proj_Data_Entry.get()), Style1)
    Excel_WorkSheet.write(3, 0, "Date of Report:" + str(date_today), Style1)
    Excel_WorkSheet.write(4, 0, "Time of Report:" + str(Hr) + ":" + str(Min) + ":" + str(Sec), Style1)
    Excel_WorkSheet.write(5, 0, "Parameter", Style5)
    Excel_WorkSheet.write(5, 1, "Set Value", Style5)
    Excel_WorkSheet.col(0).width=12000
    Excel_WorkSheet.col(1).width = 5000
    Excel_WorkSheet.write(6, 0, "Lightning Strike Density in Strikes/Km^2/Year", Style4)
    Excel_WorkSheet.write(6, 1, Ng, Style4)
    Excel_WorkSheet.write(7, 0, "Danger to People Factor", Style4)
    Excel_WorkSheet.write(7, 1, h, Style4)
    Excel_WorkSheet.write(8, 0, "Occupancy Coeffecient", Style4)
    Excel_WorkSheet.write(8, 1, Lf1, Style4)
    Excel_WorkSheet.write(9, 0, "Occupancy Coeffecient (Lf1)", Style4)
    Excel_WorkSheet.write(9, 1, Lf1, Style4)
    Excel_WorkSheet.write(10, 0, "Relative Location of Site (Cd)", Style4)
    Excel_WorkSheet.write(10, 1, Cd, Style4)
    Excel_WorkSheet.write(11, 0, "Fire Risk to Site(rf)", Style4)
    Excel_WorkSheet.write(11, 1, rf, Style4)
    Excel_WorkSheet.write(12, 0, "Details of Associated Services", Style4)
    Excel_WorkSheet.write(12, 1, Lf2, Style4)
    Excel_WorkSheet.write(13, 0, "Details of Electrical Lines", Style4)
    Excel_WorkSheet.write(13, 1, Ai, Style4)
    Excel_WorkSheet.write(14, 0, "Lighting Protection Level Factor(Pd)", Style4)
    Excel_WorkSheet.write(14, 1, Pd, Style4)
    Excel_WorkSheet.write(15, 0, "Details of Surge Arrestors", Style4)
    Excel_WorkSheet.write(15, 1, Pi, Style4)
    Excel_WorkSheet.write(17, 0, "Final Results", Style4)
    Excel_WorkSheet.write(18, 0, "Loss to Human Life( RT1 )", Style4)
    Excel_WorkSheet.write(18, 1, RT1, Style4)

    Excel_WorkSheet.write(20, 0, "Loss of Service to Public ( RT2 )", Style4)
    Excel_WorkSheet.write(20, 1, RT2, Style4)

    Excel_WorkSheet.write(22, 0, "Loss to Cultural Heritage ( RT3 )", Style4)
    Excel_WorkSheet.write(22, 1, RT3, Style4)

    if RT1>=0.00001:
        Excel_WorkSheet.write(24, 0, "RT1 - Need Lightning Protection", Style6)
    else:
        Excel_WorkSheet.write(24, 0, "RT1 - Safe No need for  Lightning Protection", Style5)

    if RT2>=0.001:
        Excel_WorkSheet.write(26, 0, "RT2 - Need Lightning Protection", Style6)
    else:
        Excel_WorkSheet.write(26, 0, "RT2 - Safe No need for  Lightning Protection", Style5)

    if RT3>=0.001:
        Excel_WorkSheet.write(28, 0, "RT3 - Need Lightning Protection", Style6)
    else:
        Excel_WorkSheet.write(28, 0, "RT3 - Safe No need for  Lightning Protection", Style5)

    ##########SAVE IN EXCEL FILE#################
    Excel_Workbook.save("Results/" + File_Name_Excel_Lightning)
    Done_Lightning = messagebox.showinfo("Results Data..", "Results Stored in Excel File :" + str(File_Name_Excel_Lightning))

##############################ELEVATOR CALCULATIONS##################
def Calc_Elevator(Tab6):
    # READING ALL THE VARAIABLES
    if Entry.get(Ent1) == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                      "Please Enter the Name of Project...  ")
        return
    if Entry.get(Ent2) == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                       "Please Enter the Total Site Population...  ")
        return

    if Entry.get(Ent3) == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                       "Please Enter the Number of Floors...  ")
        return

    if Entry.get(Ent4) == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                       "Please Enter the Floor Height in Meters...  ")
        return

    if Entry.get(Ent5) == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                       "Please Enter the Door Open Time in Secs...  ")
        return

    if Entry.get(Ent6) == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                       "Please Enter the Door Close Time in Secs...  ")
        return

    if Entry.get(Ent7) == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                       "Please Enter the Passenger Transfer Time in Secs...  ")
        return

    if Entry.get(Ent8) == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                       "Please Enter the Extra Time or 0.0 for zero time..  ")
        return

    if building_type.get() == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                       "Please Select the Building type from the Building Type Combobox..  ")
        return

    if Min_Elev_Speed.get() == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                       "Please Select the Minimum Elevator Speed from The Minimum Elevator Speed Combobox  ")
        return

    if Wait_Int_Cat.get() == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                       "Please Select the Waiting Interval Category from the Waiting Interval Category ComboBox..  ")
        return


    if Entry.get(Ent16) == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                       "Please Enter the Number of Stops..  ")
        return

    if Entry.get(Ent17) == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                       "Please Enter the Elevator Capacity factor(Typical 0.8)..  ")
        return

    if Entry.get(Ent18) == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                       "Please Enter the Elevator Acceleration in m/s2(Typical 1 m/s2)...")
        return

    # ADD THE PROJECT NAME WORKSHEET
    # EXCEL FILE NAME GENERATION
    Current_Time = datetime.datetime.now()
    Hr = Current_Time.hour
    Min = Current_Time.minute
    Sec = Current_Time.second
    date_today = date.today()
    File_Name_Excel = 'Results/' + "Elevator_Design_" + str(Hr) + "_" + str(Min) + "_" + str(Sec) + ".xls"

    # WORSHEET SAVING DATA
    Excel_WorkSheet = Excel_Workbook.add_sheet('Project_' + str(Entry.get(Ent1)))
    Excel_Summary = Excel_Workbook.add_sheet('Project_' + str(Entry.get(Ent1) + "_Summary"))
    Excel_WorkSheet.write_merge(0, 1, 0, 6, "Tata Realty and Infrastructure - Elevator Design Data(UP PEAK)", Style0)
    Excel_WorkSheet.write(2, 0, "Project:" + str(Entry.get(Ent1)), Style1)
    Excel_WorkSheet.write(3, 0, "Date of Report:" + str(date_today), Style1)
    Excel_WorkSheet.write(4, 0, "Time of Report:" + str(Hr) + ":" + str(Min) + ":" + str(Sec), Style1)
    Excel_WorkSheet.write(6, 0, "Sl.No", Style4)
    Excel_WorkSheet.write(6, 1, "Speed-M/s", Style4)
    Excel_WorkSheet.write(6, 2, "Number of Elevators", Style4)
    Excel_WorkSheet.write(6, 3, "Max Passenger Capacity", Style4)
    Excel_WorkSheet.write(6, 4, "Up Peak RTT(Secs)", Style4)
    Excel_WorkSheet.write(6, 5, "Waiting Interval (Secs)", Style4)
    Excel_WorkSheet.write(6, 6, "5 Min Handling Capacity-%", Style4)
    Excel_WorkSheet.write(6, 7, "H-Highest Reversal Floor", Style4)
    Excel_WorkSheet.write(6, 8, "Time for Traversal of full Height", Style4)
    Excel_WorkSheet.write(6, 9, "RTT Down Peak", Style4)
    Excel_WorkSheet.write(6, 10,"RTT Mid Day Peak", Style4)

    # SUMMARY SHEET FILLING DATA
    Excel_Summary.write_merge(0, 1, 0, 6, "Results Summary", Style0)
    Excel_Summary.write(2, 0, "Project:" + str(Entry.get(Ent1)), Style1)
    Excel_Summary.write(3, 0, "Date of Report:" + str(date_today), Style1)
    Excel_Summary.write(4, 0, "Time of Report:" + str(Hr) + ":" + str(Min) + ":" + str(Sec), Style1)
    Excel_Summary.write(6, 0, "Sl.No", Style4)
    Excel_Summary.write(6, 1, "Speed-M/s", Style4)
    Excel_Summary.write(6, 2, "Waiting Interval (Secs)", Style4)
    Excel_Summary.write(6, 3, "5 Min Handling Capacity(%)", Style4)
    Excel_Summary.write(6, 4, "RTT(Secs)", Style4)
    Excel_Summary.write(6, 5, "Number Of Elevators)", Style4)
    Excel_Summary.write(6, 6, "Traverse Time(Secs)", Style4)
    Excel_Summary.write(6, 7, "Average Number of Stops(S)", Style4)

    # DATA AREA
    Speed_Range = [0.75, 1.0, 1.5, 1.75, 2.0, 2.5, 3.0, 4.0]
    Elev_Capacity = [6, 8, 12, 16, 18, 20, 24]
    Waiting_Interval = [item for item in range(1, 200, 1)]
    Population = int(Entry.get(Ent2))
    Excel_WorkSheet.write(5, 0, "Site Population:" + str(Population), Style1)
    inc = 1
    inc_sum = 1

    # GETTING THE BUILDING TYPE
    build_type = building_type.get()
    if build_type == ' Multi Tenant Commercial':
        FMHCMin = 10
    if build_type == ' Single Tenant Commercial':
        FMHCMin = 15
    if build_type == ' Residential':
        FMHCMin = 7.5

        # GETTING THE MINIMUM ELEVATOR SPEED
    min_elev_speed = Min_Elev_Speed.get()
    if min_elev_speed == '0.75':
        MINELEVSPD = 0.75
    if min_elev_speed == '1.0':
        MINELEVSPD = 1.0
    if min_elev_speed == '1.5':
        MINELEVSPD = 1.5
    if min_elev_speed == '1.75':
        MINELEVSPD = 1.75
    if min_elev_speed == '2.0':
        MINELEVSPD = 2.0
    if min_elev_speed == '2.5':
        MINELEVSPD = 2.5
    if min_elev_speed == '3.0':
        MINELEVSPD = 3.0
    if min_elev_speed == '4.0':
        MINELEVSPD = 4.0
        # GETTING THE WAITING INTERVAL CATEGORY

    wait_int_cat = Wait_Int_Cat.get()
    if wait_int_cat == 'Excellent':
        Waiting_Int_Min = 0.0
        Waiting_Int_Max = 25
    if wait_int_cat == 'Very Good':
        Waiting_Int_Min = 0.0
        Waiting_Int_Max = 30.0
    if wait_int_cat == 'Good':
        Waiting_Int_Min = 0.0
        Waiting_Int_Max = 35.0
    if wait_int_cat == 'Fair':
        Waiting_Int_Min = 0.0
        Waiting_Int_Max = 40.0
    if wait_int_cat == 'Poor':
        Waiting_Int_Min = 0.0
        Waiting_Int_Max = 45.0
    if wait_int_cat == 'Unsatisfactory':
        Waiting_Int_Min = 45.0
        Waiting_Int_Max = 100.0

    #DATA COLLECTION
    # RETRIVING ALL DATA FROM FORM AND CALCULATING
    tex = float(Entry.get(Ent8))

    # N IS NUMBER OF FLOORS
    N = int(Entry.get(Ent3))

    # ELEVATOR DATA
    tp = float(Entry.get(Ent7))
    to = float(Entry.get(Ent5))
    tc =  float(Entry.get(Ent6))
    Flr_Ht = float(Entry.get(Ent4))
    Elv_Cap_Factor = float(Entry.get(Ent17))
    Acc = float(Entry.get(Ent18))
    S=float(Entry.get(Ent16))
    Population = int(Entry.get(Ent2))



    # START OF NESTED LOOPS
    for k in Waiting_Interval:
        for j in Elev_Capacity:
            for i in Speed_Range:

                # RETRIVING ALL DATA FROM FORM AND CALCULATING

                Speed = i
                Pmax = j
                P = Elv_Cap_Factor * Pmax
                tv = float(Entry.get(Ent4)) / Speed
                tf1 = Speed / Acc
                T = to + tc + tf1 + tex
                INT = k
                Elevator_Speed_Min = MINELEVSPD

                # TOTAL HEIGHT OF TRAVEL
                TOTHEIGHT = int(Entry.get(Ent3)) * float(Entry.get(Ent4))

                # CALCULATION OF H - AVEARGE HIGHEST REVERSAL FLOOR
                HSum = 0
                for i in range(1, N - 1):
                    HSum = HSum + (i / N) ** P
                H = N - HSum


                # CALCULATION OF RTT
                RTT = 2 * H * tv + (S + 1) * (T - tv) + 2 * tp
                RTTDOWNPEAK=N*tv+(0.5*S+1)*(T-tv)+2*tp
                RTTMIDDAY=2*H*tv+2*S*(T-tv)+4*tp

                # CALCULATION OF NUMBER OF ELEVATORS

                L = RTT / INT

                # CALCULATION OF 5 MIN HANDLING CAPACITY

                FIVEMINHC = (300 * P * L)* 100/ (RTT*Population)

                TRAV_TIME = (int(Entry.get(Ent3)) * float(Entry.get(Ent4))) / Speed

                # SEND DATA TO EXCEL

                Excel_WorkSheet.write(inc + 6, 0, inc, Style2)
                Excel_WorkSheet.write(inc + 6, 1, Speed, Style2)
                Excel_WorkSheet.write(inc + 6, 2, math.ceil(L), Style2)
                Excel_WorkSheet.write(inc + 6, 3, round(Pmax, 1), Style2)
                Excel_WorkSheet.write(inc + 6, 4, round(RTT, 2), Style2)


                #5 MIN HANDLING CAPACITY COLOR CODING
                if (FIVEMINHC >= FMHCMin) and (INT <= Waiting_Int_Max):
                    Excel_WorkSheet.write(inc + 6, 6, round(FIVEMINHC, 2), Style3)

                else:
                    Excel_WorkSheet.write(inc + 6, 6, round(FIVEMINHC, 2), Style2)

                #WAITING INTERVAL COLOUR CODING
                if INT <= Waiting_Int_Max:
                    Excel_WorkSheet.write(inc + 6, 5, round(INT, 2), Style3)
                else:
                    Excel_WorkSheet.write(inc + 6, 5, round(INT, 2), Style2)

                Excel_WorkSheet.write(inc + 6, 7, round(H, 2), Style2)
                Excel_WorkSheet.write(inc + 6, 8, round(TRAV_TIME, 2), Style2)
                Excel_WorkSheet.write(inc + 6, 9, round(RTTDOWNPEAK, 2), Style2)
                Excel_WorkSheet.write(inc + 6, 10, round(RTTMIDDAY, 2), Style2)


                inc = inc + 1
                # SUMMARY DATA
                if (Speed >= Elevator_Speed_Min and FIVEMINHC >= FMHCMin and (
                        INT >= Waiting_Int_Min and INT <= Waiting_Int_Max) and TRAV_TIME <= 60.0):
                    Excel_Summary.write(inc_sum + 6, 0, inc_sum, Style2)
                    Excel_Summary.write(inc_sum + 6, 1, Speed, Style2)
                    Excel_Summary.write(inc_sum + 6, 2, INT, Style3)
                    Excel_Summary.write(inc_sum + 6, 3, round(FIVEMINHC, 3), Style3)
                    Excel_Summary.write(inc_sum + 6, 4, round(RTT, 3), Style2)
                    Excel_Summary.write(inc_sum + 6, 5, math.ceil(L), Style2)
                    Excel_Summary.write(inc_sum + 6, 6, round(TRAV_TIME, 3), Style2)
                    Excel_Summary.write(inc_sum + 6, 7, round(S, 3), Style2)
                    inc_sum = inc_sum + 1

    # BASIC RESULTS DISPLAY
    Ent9.config(text="{:.2f}".format(H))
    Ent10.config(text="{:.2f}".format(S))
    Ent11.config(text="{:.2f}".format(TOTHEIGHT))
    Ent12.config(text="{:.2f}".format(tv))
    Ent13.config(text="{:.2f}".format(T))

    Excel_WorkSheet.write(inc + 8, 0, " Note: As per NBC 2016", Style2)
    Excel_WorkSheet.write(inc + 9, 0, " <Commercial>", Style2)
    Excel_WorkSheet.write(inc + 10, 0, " 5 Min Handling Capacity:", Style2)
    Excel_WorkSheet.write(inc + 11, 0, " ---Multi Tenant = 10 -15 %", Style2)
    Excel_WorkSheet.write(inc + 12, 0, " ---Single Tenant = 15-25 %", Style2)
    Excel_WorkSheet.write(inc + 13, 0, " Waiting Interval:", Style2)
    Excel_WorkSheet.write(inc + 14, 0, " ---Excellent = Less Than 25 Secs", Style2)
    Excel_WorkSheet.write(inc + 15, 0, " ---Very Good = 25 to less than 30 Secs", Style2)
    Excel_WorkSheet.write(inc + 16, 0, " ---Good = 30 to less than 35 Secs", Style2)
    Excel_WorkSheet.write(inc + 17, 0, " ---Fair = 35 to less than 40 Secs", Style2)
    Excel_WorkSheet.write(inc + 18, 0, " ---Poor = 40 to less than 45 Secs", Style2)
    Excel_WorkSheet.write(inc + 19, 0, " ---Unsatisfactory = Greater than 45", Style2)
    # FOR RESIDENTIAL
    Excel_WorkSheet.write(inc + 20, 0, " <Residential>", Style2)
    Excel_WorkSheet.write(inc + 21, 0, " 5 Min Handling Capacity:", Style2)
    Excel_WorkSheet.write(inc + 22, 0, " ---HIGH END = >8%", Style2)
    Excel_WorkSheet.write(inc + 23, 0, " ---MID END = 6-8 %", Style2)
    Excel_WorkSheet.write(inc + 24, 0, " ---LOW END = 5-7 %", Style2)
    Excel_WorkSheet.write(inc + 25, 0, " Waiting Interval:", Style2)
    Excel_WorkSheet.write(inc + 26, 0, " ---HIGH END = <30 Secs", Style2)
    Excel_WorkSheet.write(inc + 27, 0, " ---MID END = 31 to less than 45 Secs", Style2)
    Excel_WorkSheet.write(inc + 28, 0, " ---LOW END = 46 to less than 60 Secs", Style2)

    # SUMMARY SHEET FILLING
    Excel_Summary.write(inc_sum + 8, 0, " Note: As per NBC 2016", Style2)
    Excel_Summary.write(inc_sum + 9, 0, " <Commercial> ", Style2)
    Excel_Summary.write(inc_sum + 10, 0, " 5 Min Handling Capacity:", Style2)
    Excel_Summary.write(inc_sum + 11, 0, " ---Multi Tenant = 10 -15 %", Style2)
    Excel_Summary.write(inc_sum + 12, 0, " ---Single Tenant = 15-25 %", Style2)
    Excel_Summary.write(inc_sum + 13, 0, " Waiting Interval:", Style2)
    Excel_Summary.write(inc_sum + 14, 0, " ---Excellent = Less Than 25 Secs", Style2)
    Excel_Summary.write(inc_sum + 15, 0, " ---Very Good = 25 to less than 30 Secs", Style2)
    Excel_Summary.write(inc_sum + 16, 0, " ---Good = 30 to less than 35 Secs", Style2)
    Excel_Summary.write(inc_sum + 17, 0, " ---Fair = 35 to less than 40 Secs", Style2)
    Excel_Summary.write(inc_sum + 18, 0, " ---Poor = 40 to less than 45 Secs", Style2)
    Excel_Summary.write(inc_sum + 19, 0, " ---Unsatisfactory = Greater than 45", Style2)

    Excel_Summary.write(inc + 20, 0, " <Residential>", Style2)
    Excel_Summary.write(inc + 21, 0, " 5 Min Handling Capacity:", Style2)
    Excel_Summary.write(inc + 22, 0, " ---HIGH END = >8%", Style2)
    Excel_Summary.write(inc + 23, 0, " ---MID END = 6-8 %", Style2)
    Excel_Summary.write(inc + 24, 0, " ---LOW END = 5-7 %", Style2)
    Excel_Summary.write(inc + 25, 0, " Waiting Interval:", Style2)
    Excel_Summary.write(inc + 26, 0, " ---HIGH END = <30 Secs", Style2)
    Excel_Summary.write(inc + 27, 0, " ---MID END = 31 to less than 45 Secs", Style2)
    Excel_Summary.write(inc + 28, 0, " ---LOW END = 46 to less than 60 Secs", Style2)

    # SAVE IN EXCEL FILE
    Excel_Workbook.save(File_Name_Excel)
    done = messagebox.showinfo("Results Data..",
                               "Results Stored in Excel File :" + str(File_Name_Excel) + " With " + str(
                                   inc) + " of Combinations...Please check Excel file")
    print(File_Name_Excel)
    ##OPEN EXCEL FILE IF CHECK BOX CHECKED
    if Excel_Check.get() == 1:
        os.system('start EXCEL.EXE ' + str(File_Name_Excel))

def Updates(Tab7):
    os.system('start NOTEPAD.EXE ' + 'Updates/Elect_Engg_Suite_1.0.txt')


#####################################################################
################### BASIC SCREEN SETTINGS ###########################
#####################################################################
root.geometry("1400x800")
tabControl = ttk.Notebook(root)

##### TAB SETTINGS #####
Tab1 = Frame(tabControl,bg='aliceblue')
Tab2 = Frame(tabControl,bg='wheat')
Tab3 = Frame(tabControl,bg='cornsilk')
Tab4 = Frame(tabControl,bg='beige')
Tab5 = Frame(tabControl,bg='ivory')
Tab6 = Frame(tabControl,bg='mistyrose')
Tab7=Frame(tabControl,bg='lightblue')
Tab8=Frame(tabControl,bg='beige')

tabControl.add(Tab1, text=' System Fault Level Analysis ',padding=5)
tabControl.add(Tab2, text=' Lightning Risk Level Analysis ',padding=5)
tabControl.add(Tab3, text=' Cable Details Query ',padding=5)
tabControl.add(Tab4, text=' Capacitor Bank Design ',padding=5)
tabControl.add(Tab5, text=' Protection Relay Query ',padding=5)
tabControl.add(Tab6, text=' Elevator Design Up Peak  ',padding=5)
tabControl.add(Tab7, text=' Updates and Revisions ',padding=5)
tabControl.add(Tab8, text=' Voltage Drop Estimation ',padding=5)

tabControl.pack(expand=1, fill="both")

#########################################################################################
###################### TAB-1 SYSTEM FAULT LEVEL ANALYSIS ################################
#########################################################################################

##### BASIC PROJECT DETAILS #####

Basic_Proj_Lbl = Label(Tab1, text='Basic Project Details:', bg="lightblue", fg="red", relief=RIDGE, width=25).grid(
    row=0, column=0)
Name_Proj_Lbl = Label(Tab1, text='Name of Project:', relief=RIDGE, width=25).grid(row=1, column=0)
Name_Proj_Entry = Entry(Tab1, bg='yellow', relief=RIDGE, width=30)
Name_Proj_Entry.grid(row=1, column=1)

##### AUTO DATE RETRIVAL #####
Current_Time = datetime.datetime.now()
Hr = Current_Time.hour
Min = Current_Time.minute
Sec = Current_Time.second
date_today = date.today()

##### AUTO DATE ENTRY #####
Date_Lbl = Label(Tab1, text='Date:', relief=RIDGE, width=25).grid(row=2, column=0)
Date_Entry = Label(Tab1, bg='lightblue', relief=RIDGE, width=25, text=str(Hr) + ":" + str(Min) + ":" + str(Sec)).grid(
    row=2, column=1)

##### CALCULATION PARAMETERS #####
Calc_Parameters = Label(Tab1, text='Calculation Parameters:', bg="lightblue", fg="red", relief=RIDGE, width=25).grid(
    row=3, column=0)

##### HT VOLTAGE #####
HT_Voltage_Lbl = Label(Tab1, text='HT Voltage:', relief=RIDGE, width=25).grid(row=4, column=0)
HT_Voltage_Entry = Entry(Tab1, bg='yellow', relief=RIDGE, width=30)
HT_Voltage_Entry.grid(row=4, column=1)

##### HT FAULT LEVEL MVA #####
Fault_Level_HT_Lbl = Label(Tab1, text='Fault Level at HT Incoming(MVA)', relief=RIDGE, width=25).grid(row=5, column=0)
Fault_Level_HT_Entry = Entry(Tab1, bg='yellow', relief=RIDGE, width=30)
Fault_Level_HT_Entry.grid(row=5, column=1)

##### HT TRANSFORMER RATING #####
Transformer_Rating_Lbl = Label(Tab1, text='Transformer Rating(MVA)', relief=RIDGE, width=25).grid(row=6, column=0)
Transformer_Rating_Entry = Entry(Tab1, bg='yellow', relief=RIDGE, width=30)
Transformer_Rating_Entry.grid(row=6, column=1)

##### HT TRANSFORMER IMPEDENCE #####
Transformer_Impedence_Lbl = Label(Tab1, text='Transformer Impedence(%)', relief=RIDGE, width=25).grid(row=7, column=0)
Transformer_Impedence_Entry = Entry(Tab1, bg='yellow', relief=RIDGE, width=30)
Transformer_Impedence_Entry.grid(row=7, column=1)

##### HT BASE KVA #####
HT_BASEKVA_Lbl = Label(Tab1, text='Base MVA HT(Brkr+XMer Pri)', relief=RIDGE, width=25).grid(row=8, column=0)
HT_BASEKVA_Estimated = Entry(Tab1, bg='yellow', relief=RIDGE, width=30)
HT_BASEKVA_Estimated.grid(row=8, column=1)

##### HT BASE KV #####
HT_BASEKV_Lbl = Label(Tab1, text='Base KV HT(Brkr+XMer Pri)', relief=RIDGE, width=25).grid(row=9, column=0)
HT_BASEKV_Estimated = Entry(Tab1, bg='yellow', relief=RIDGE, width=30)
HT_BASEKV_Estimated.grid(row=9, column=1)

##### BASE KVA LT #####
LT_BASEKVA_Lbl = Label(Tab1, text='Base MVA LT(XMer Sec to Down)', relief=RIDGE, width=25).grid(row=10, column=0)
LT_BASEKVA_Estimated = Entry(Tab1, bg='yellow', relief=RIDGE, width=30)
LT_BASEKVA_Estimated.grid(row=10, column=1)

##### BASE KV LT #####
LT_BASEKV_Lbl = Label(Tab1, text='Base V LT - Volts', relief=RIDGE, width=25).grid(row=11, column=0)
LT_BASEKV_Entry = Entry(Tab1, bg='yellow', relief=RIDGE, width=30)
LT_BASEKV_Entry.grid(row=11, column=1)

########## LEVEL A CALCULATIONS ##########
Cable_Main = pd.read_csv("Data\\Cable_Data_Small.csv")
Cable_Sizes = Cable_Main["Cable Size"]

Section_A_Lbl = Message(Tab1, text='A. Fault Level from HT Incoming to HT Breaker', bg="deepskyblue", fg="red",
                        relief=RIDGE, width=190, font='Calbri 10 bold').grid(row=12, column=0)

##### HT CABLE TYPE #####
HT_Cable_Size_Lbl = Label(Tab1, text='HT Cable Size- Sq mm)', relief=RIDGE, width=25).grid(row=13, column=0)
Search_HT_Cable_Size = tkinter.StringVar()
HT_Cable_Size = ttk.Combobox(Tab1, width=25, state="readonly", values=Cable_Sizes.tolist(),
                             textvariable=Search_HT_Cable_Size)
HT_Cable_Size.grid(row=13, column=1)
HT_Cable_Size.bind('<<ComboboxSelected>>', HT_Cable_Selected_A)

##### RESISTANCE AND REACTANCE VALUE DISPLAY #####
Res_Reac_Lbl_A = Label(Tab1, text='Res,Reac Ohms/Km=', relief=RIDGE, width=25).grid(row=14, column=0)
Res_Reac_A = Label(Tab1, text='----', relief=RIDGE, width=25)
Res_Reac_A.grid(row=14, column=1)

##### CABLE RUNS #####
A_HT_Cable_Runs_Lbl = Label(Tab1, text='HT Cable Runs)', relief=RIDGE, width=25).grid(row=15, column=0)
A_HT_Cable_Runs_Entry = Entry(Tab1, bg='yellow', relief=RIDGE, width=30)
A_HT_Cable_Runs_Entry.grid(row=15, column=1)

##### A HT CABLE LENGTH #####
A_HT_Cable_Length_Lbl = Label(Tab1, text='HT Cable Length (m)', relief=RIDGE, width=25).grid(row=16, column=0)
A_HT_Cable_Length_Entry = Entry(Tab1, bg='yellow', relief=RIDGE, width=30)
A_HT_Cable_Length_Entry.grid(row=16, column=1)

# A HT CABLE IMPEDENCE
A_HT_Cable_Impedence_Lbl = Label(Tab1, text='HT Cable Impedence (Z1)', relief=RIDGE, width=25).grid(row=17, column=0)
A_HT_Cable_Impedence_Calculation = Label(Tab1, text='', relief=RIDGE, width=25)
A_HT_Cable_Impedence_Calculation.grid(row=17, column=1)

# U REACTANCE OF BREAKER(Xpu)
A_U_Reactance_Lbl = Label(Tab1, text='HT Breaker U Reactance(Xpu/Z2)', relief=RIDGE, width=25).grid(row=18, column=0)
A_U_Reactance_Lbl_Calculation = Label(Tab1, text='', relief=RIDGE, width=25)
A_U_Reactance_Lbl_Calculation.grid(row=18, column=1)

# TOTAL IMPEDENCE
A_Total_Impedence_Lbl = Label(Tab1, text='Total Impedence(Z1+Z2)', relief=RIDGE, width=25).grid(row=19, column=0)
A_Total_Impedence_Calculation = Label(Tab1, text='', relief=RIDGE, width=25)
A_Total_Impedence_Calculation.grid(row=19, column=1)

# FAULT MVA AT HT CB
A_Fault_MVA_HTCB_Lbl = Label(Tab1, text='Fault MVA at HT CB', relief=RIDGE, width=25).grid(row=20, column=0)
A_Fault_MVA_HTCB_Calculation = Label(Tab1, text='', relief=RIDGE, width=25)
A_Fault_MVA_HTCB_Calculation.grid(row=20, column=1)

# FAULT MVA AT LOCATION A
A_Fault_Lbl = Label(Tab1, text='Fault Current at Level-A', relief=RIDGE, width=25).grid(row=21, column=0)
A_Fault_Calculation = Label(Tab1, text='', relief=RIDGE, width=25)
A_Fault_Calculation.grid(row=21, column=1)

# FAULT LEVEL AT LEVEL B
Section_B_Lbl = Message(Tab1, text="B. Fault Level from HT Breaker to X'mer Primary", bg="deepskyblue", fg="red",
                        relief=RIDGE, width=190, font='Calbri 10 bold').grid(row=22, column=0)

# HT CABLE TYPE- LEVEL B
Cable_Size_Lbl_B = Label(Tab1, text='HT Cable Size- Sq mm)', relief=RIDGE, width=25).grid(row=23, column=0)
Search_HT_Cable_Size_B = tkinter.StringVar()
HT_Cable_Size_B = ttk.Combobox(Tab1, width=25, state="readonly", values=Cable_Sizes.tolist(),
                               textvariable=Search_HT_Cable_Size_B)
HT_Cable_Size_B.grid(row=23, column=1)
HT_Cable_Size_B.bind('<<ComboboxSelected>>', HT_Cable_Selected_B)

# RESISTANCE AND REACTANCE VALUE DISPLAY-B
Res_Reac_Lbl_B = Label(Tab1, text='Res,Reac Ohms/Km=', relief=RIDGE, width=25).grid(row=24, column=0)
Res_Reac_B = Label(Tab1, text='----', relief=RIDGE, width=25)
Res_Reac_B.grid(row=24, column=1)

# CABLE RUNS-B
B_Cable_Runs_Lbl = Label(Tab1, text='HT Cable Runs)', relief=RIDGE, width=25).grid(row=25, column=0)
B_Cable_Runs_Entry = Entry(Tab1, bg='yellow', relief=RIDGE, width=30)
B_Cable_Runs_Entry.grid(row=25, column=1)

# A HT CABLE LENGTH-B
B_Cable_Length_Lbl = Label(Tab1, text='HT Cable Length (m)', relief=RIDGE, width=25).grid(row=26, column=0)
B_Cable_Length_Entry = Entry(Tab1, bg='yellow', relief=RIDGE, width=30)
B_Cable_Length_Entry.grid(row=26, column=1)

# TOTAL IMPEDENCE
B_Total_Impedence_Lbl = Label(Tab1, text='Total Impedence(Z3+Z4+Z5)', relief=RIDGE, width=25).grid(row=27, column=0)
B_Total_Impedence_Calculation = Label(Tab1, text='', relief=RIDGE, width=25)
B_Total_Impedence_Calculation.grid(row=27, column=1)

# FAULT MVA AT HT CB
B_Fault_MVA_Lbl = Label(Tab1, text='Fault MVA at HT CB', relief=RIDGE, width=25).grid(row=28, column=0)
B_Fault_MVA_Calculation = Label(Tab1, text='', relief=RIDGE, width=25)
B_Fault_MVA_Calculation.grid(row=28, column=1)

# FAULT CURRENT AT LOCATION B
B_Fault_Current_Lbl = Label(Tab1, text='Fault Current at Level-B', relief=RIDGE, width=25).grid(row=29, column=0)
B_Fault_Current_Calculation = Label(Tab1, text='', relief=RIDGE, width=25)
B_Fault_Current_Calculation.grid(row=29, column=1)

############################# LEVEL -C #######################################
# FAULT LEVEL AT LEVEL C
Section_C_Lbl = Message(Tab1, text="C. Fault Level from X'mer Primary to Secondary", bg="deepskyblue", fg="red",
                        relief=RIDGE, width=190, font='Calbri 10 bold').grid(row=0, column=2)

# TOTAL IMPEDENCE-LEVEL C
C_Total_Impedence_Lbl = Label(Tab1, text='Total Impedence(Z6+Z7)', relief=RIDGE, width=25).grid(row=1, column=2)
C_Total_Impedence_Calculation = Label(Tab1, text='', relief=RIDGE, width=25)
C_Total_Impedence_Calculation.grid(row=1, column=3)

# FAULT MVA AT TRANSFORMER
C_Fault_MVA_Lbl = Label(Tab1, text='Fault MVA at Transformer ', relief=RIDGE, width=25).grid(row=2, column=2)
C_Fault_MVA_Calculation = Label(Tab1, text='', relief=RIDGE, width=25)
C_Fault_MVA_Calculation.grid(row=2, column=3)

# FAULT CURRENT AT LOCATION C
C_Fault_Current_Lbl = Label(Tab1, text='Fault Current at Level-C', relief=RIDGE, width=25).grid(row=3, column=2)
C_Fault_Current_Calculation = Label(Tab1, text='', relief=RIDGE, width=25)
C_Fault_Current_Calculation.grid(row=3, column=3)

############################# LEVEL - D ########################################
Section_D_Lbl = Message(Tab1, text="D. Fault Level at LT side SS to Downstream", bg="deepskyblue", fg="red",
                        relief=RIDGE, width=190, font='Calbri 10 bold').grid(row=4, column=2)

# LT CABLE TYPE- LEVEL D
Cable_Size_Lbl_D = Label(Tab1, text='LT Cable Size- Sq mm)', relief=RIDGE, width=25).grid(row=5, column=2)
Search_HT_Cable_Size_D = tkinter.StringVar()
LT_Cable_Size_D = ttk.Combobox(Tab1, width=25, state="readonly", values=Cable_Sizes.tolist(),
                               textvariable=Search_HT_Cable_Size_D)
LT_Cable_Size_D.grid(row=5, column=3)
LT_Cable_Size_D.bind('<<ComboboxSelected>>', HT_Cable_Selected_D)

# RESISTANCE AND REACTANCE VALUE DISPLAY-D
Res_Reac_Lbl_D = Label(Tab1, text='Res,Reac Ohms/Km=', relief=RIDGE, width=25).grid(row=6, column=2)
Res_Reac_D = Label(Tab1, text='----', relief=RIDGE, width=25)
Res_Reac_D.grid(row=6, column=3)

# LT CABLE RUNS-D
D_Cable_Runs_Lbl = Label(Tab1, text='LT Cable Runs)', relief=RIDGE, width=25).grid(row=7, column=2)
D_Cable_Runs_Entry = Entry(Tab1, bg='yellow', relief=RIDGE, width=30)
D_Cable_Runs_Entry.grid(row=7, column=3)

# LT CABLE LENGTH-D
D_Cable_Length_Lbl = Label(Tab1, text='LT Cable Length (m)', relief=RIDGE, width=25).grid(row=8, column=2)
D_Cable_Length_Entry = Entry(Tab1, bg='yellow', relief=RIDGE, width=30)
D_Cable_Length_Entry.grid(row=8, column=3)

# TOTAL IMPEDENCE- D
D_Total_Impedence_Lbl = Label(Tab1, text='Total Impedence(Z8+Z9+Z10)', relief=RIDGE, width=25).grid(row=9, column=2)
D_Total_Impedence_Calculation = Label(Tab1, text='', relief=RIDGE, width=25)
D_Total_Impedence_Calculation.grid(row=9, column=3)

# FAULT MVA AT LEVEL D
D_Fault_MVA_Lbl = Label(Tab1, text='Fault MVA at Main LT Panel', relief=RIDGE, width=25).grid(row=10, column=2)
D_Fault_MVA_Calculation = Label(Tab1, text='', relief=RIDGE, width=25)
D_Fault_MVA_Calculation.grid(row=10, column=3)

# FAULT CURRENT AT LOCATION D
D_Fault_Current_Lbl = Label(Tab1, text='Fault Current at Level-D', relief=RIDGE, width=25).grid(row=11, column=2)
D_Fault_Current_Calculation = Label(Tab1, text='', relief=RIDGE, width=25)
D_Fault_Current_Calculation.grid(row=11, column=3)

############################# LEVEL - E ########################################
Section_E_Lbl = Message(Tab1, text="E. Fault Level from LT Panel to Sub Panel", bg="deepskyblue", fg="red",
                        relief=RIDGE, width=190, font='Calbri 10 bold').grid(row=12, column=2)

# LT CABLE TYPE- LEVEL E
Cable_Size_Lbl_E = Label(Tab1, text='LT Cable Size- Sq mm)', relief=RIDGE, width=25).grid(row=13, column=2)
Search_HT_Cable_Size_E = tkinter.StringVar()
LT_Cable_Size_E = ttk.Combobox(Tab1, width=25, state="readonly", values=Cable_Sizes.tolist(),
                               textvariable=Search_HT_Cable_Size_E)
LT_Cable_Size_E.grid(row=13, column=3)
LT_Cable_Size_E.bind('<<ComboboxSelected>>', HT_Cable_Selected_E)

# RESISTANCE AND REACTANCE VALUE DISPLAY-E
Res_Reac_Lbl_E = Label(Tab1, text='Res,Reac Ohms/Km=', relief=RIDGE, width=25).grid(row=14, column=2)
Res_Reac_E = Label(Tab1, text='----', relief=RIDGE, width=25)
Res_Reac_E.grid(row=14, column=3)

# LT CABLE RUNS-E
E_Cable_Runs_Lbl = Label(Tab1, text='LT Cable Runs)', relief=RIDGE, width=25).grid(row=15, column=2)
E_Cable_Runs_Entry = Entry(Tab1, bg='yellow', relief=RIDGE, width=30)
E_Cable_Runs_Entry.grid(row=15, column=3)

# LT CABLE LENGTH-E
E_Cable_Length_Lbl = Label(Tab1, text='LT Cable Length (m)', relief=RIDGE, width=25).grid(row=16, column=2)
E_Cable_Length_Entry = Entry(Tab1, bg='yellow', relief=RIDGE, width=30)
E_Cable_Length_Entry.grid(row=16, column=3)

# TOTAL IMPEDENCE- E
E_Total_Impedence_Lbl = Label(Tab1, text='Total Impedence(Z11+Z12+Z13)', relief=RIDGE, width=25).grid(row=17, column=2)
E_Total_Impedence_Calculation = Label(Tab1, text='', relief=RIDGE, width=25)
E_Total_Impedence_Calculation.grid(row=17, column=3)

# FAULT MVA AT LEVEL E
E_Fault_MVA_Lbl = Label(Tab1, text='Fault MVA at LT Panel', relief=RIDGE, width=25).grid(row=18, column=2)
E_Fault_MVA_Calculation = Label(Tab1, text='', relief=RIDGE, width=25)
E_Fault_MVA_Calculation.grid(row=18, column=3)

# FAULT CURRENT AT LOCATION D
E_Fault_Current_Lbl = Label(Tab1, text='Fault Current at Level-E', relief=RIDGE, width=25).grid(row=19, column=2)
E_Fault_Current_Calculation = Label(Tab1, text='', relief=RIDGE, width=25)
E_Fault_Current_Calculation.grid(row=19, column=3)

#################### LEVEL- F ##########################################
Section_F_Lbl = Message(Tab1, text="F. Fault Level from LT Panel to MCC", bg="deepskyblue", fg="red", relief=RIDGE,
                        width=190, font='Calbri 10 bold').grid(row=20, column=2)

# LT CABLE TYPE- LEVEL F
Cable_Size_Lbl_F = Label(Tab1, text='LT Cable Size- Sq mm)', relief=RIDGE, width=25).grid(row=21, column=2)
Search_HT_Cable_Size_F = tkinter.StringVar()
LT_Cable_Size_F = ttk.Combobox(Tab1, width=25, state="readonly", values=Cable_Sizes.tolist(),
                               textvariable=Search_HT_Cable_Size_F)
LT_Cable_Size_F.grid(row=21, column=3)
LT_Cable_Size_F.bind('<<ComboboxSelected>>', HT_Cable_Selected_F)

# RESISTANCE AND REACTANCE VALUE DISPLAY-F
Res_Reac_Lbl_F = Label(Tab1, text='Res,Reac Ohms/Km=', relief=RIDGE, width=25).grid(row=22, column=2)
Res_Reac_F = Label(Tab1, text='----', relief=RIDGE, width=25)
Res_Reac_F.grid(row=22, column=3)

# LT CABLE RUNS-F
F_Cable_Runs_Lbl = Label(Tab1, text='LT Cable Runs)', relief=RIDGE, width=25).grid(row=23, column=2)
F_Cable_Runs_Entry = Entry(Tab1, bg='yellow', relief=RIDGE, width=30)
F_Cable_Runs_Entry.grid(row=23, column=3)

# LT CABLE LENGTH-F
F_Cable_Length_Lbl = Label(Tab1, text='LT Cable Length (m)', relief=RIDGE, width=25).grid(row=24, column=2)
F_Cable_Length_Entry = Entry(Tab1, bg='yellow', relief=RIDGE, width=30)
F_Cable_Length_Entry.grid(row=24, column=3)

# TOTAL IMPEDENCE- F
F_Total_Impedence_Lbl = Label(Tab1, text='Total Impedence(Z14+Z15+Z16)', relief=RIDGE, width=25).grid(row=25, column=2)
F_Total_Impedence_Calculation = Label(Tab1, text='', relief=RIDGE, width=25)
F_Total_Impedence_Calculation.grid(row=25, column=3)

# FAULT MVA AT LEVEL F
F_Fault_MVA_Lbl = Label(Tab1, text='Fault MVA at MCC Panel', relief=RIDGE, width=25).grid(row=26, column=2)
F_Fault_MVA_Calculation = Label(Tab1, text='', relief=RIDGE, width=25)
F_Fault_MVA_Calculation.grid(row=26, column=3)

# FAULT CURRENT AT LOCATION F
F_Fault_Current_Lbl = Label(Tab1, text='Fault Current at Level-E', relief=RIDGE, width=25).grid(row=27, column=2)
F_Fault_Current_Calculation = Label(Tab1, text='', relief=RIDGE, width=25)
F_Fault_Current_Calculation.grid(row=27, column=3)

# IMAGE OF POWER SYSTEM
Power_Image = Image.open("Images/Power_System_1.png")
Power_System_Image = ImageTk.PhotoImage(Power_Image)

Power_System_Lbl = Label(Tab1, image=Power_System_Image)
Power_System_Lbl.image = Power_System_Image
Power_System_Lbl.grid(row=0, column=5, rowspan=155, padx=10)

Fault_Level_Lbl = Label(Tab1, text='Fault Levels in KA', relief=RIDGE, width=25, bg='orange', fg="darkblue",
                        font='Calbri 10 bold').grid(row=0, column=4)
# FAULT LEVEL VALUES
A_Fault_Level = Label(Tab1, text='', relief=RIDGE, width=10, bg='orange', fg="darkblue", font='Calbri 10 bold')
A_Fault_Level.grid(row=1, column=4)
B_Fault_Level = Label(Tab1, text='', relief=RIDGE, width=10, bg='orange', fg="darkblue", font='Calbri 10 bold')
B_Fault_Level.grid(row=4, column=4)
C_Fault_Level = Label(Tab1, text='', relief=RIDGE, width=10, bg='orange', fg="darkblue", font='Calbri 10 bold')
C_Fault_Level.grid(row=9, column=4)
D_Fault_Level = Label(Tab1, text='', relief=RIDGE, width=10, bg='orange', fg="darkblue", font='Calbri 10 bold')
D_Fault_Level.grid(row=14, column=4)
E_Fault_Level = Label(Tab1, text='', relief=RIDGE, width=10, bg='orange', fg="darkblue", font='Calbri 10 bold')
E_Fault_Level.grid(row=19, column=4)
F_Fault_Level = Label(Tab1, text='', relief=RIDGE, width=10, bg='orange', fg="darkblue", font='Calbri 10 bold')
F_Fault_Level.grid(row=22, column=4)

# CALCULATE BUTTON
Calculate_Fault = Button(Tab1, text=' Calculate System Fault ', font='Calbri 9 bold', bg="skyblue")
Calculate_Fault.bind("<ButtonRelease-1>", Calc_System_Fault)
Calculate_Fault.grid(row=1, column=6)

# ABOUT BUTTON
Abt = Button(Tab1, text='About this Application', font='Calbri 9 bold', bg="gold")
Abt.bind("<ButtonRelease-1>", About1)
Abt.grid(row=2, column=6, pady=2)

# QUIT BUTTON
Quitter = Button(Tab1, text='Quit', font='Calbri 9 bold', bg="sandybrown",width=20)
Quitter.bind("<ButtonRelease-1>", Quitt)
Quitter.grid(row=3, column=6, pady=2)

# REPORT CHECKBOX
Excel_Report = IntVar()
Report_Check_Box = Checkbutton(Tab1, text="Save Report to Excel", onvalue=1, offvalue=0, font='Calbri 9 bold',
                               variable=Excel_Report)
Report_Check_Box.grid(row=4, column=6, pady=2)

###################################################################################################
############################# TAB-4 CAPACITOR BANK CALCULATION ####################################
###################################################################################################

Scratch_Board = tkinter.Canvas(Tab4, width=400, height=400, background='lightblue', highlightthickness=2, highlightbackground="black")
Scratch_Board.grid(row=5,column=3)
Widget_Frame = Frame(Tab4)
Widget_Frame.grid(row=0,column=2, sticky="n")

# CANVAS TEXT INITILISING
Results_Lbl = Scratch_Board.create_text(95, 15, fill="red", font="Calbri 8 bold", text="Results:")
Required_KVAR = Scratch_Board.create_text(95, 35, fill="green", font="Calbri 8 bold", text=" ")
Present_Phase_Angle = Scratch_Board.create_text(95, 55, fill="green", font="Calbri 8 bold",text="")
Target_Phase_Angle = Scratch_Board.create_text(95, 75, fill="green", font="Calbri 8 bold", text="")

Capicitor_Heading=Label(Widget_Frame, text="Capacitor Bank Calculation",font='Calbri 12 bold').grid(row=0,column=2)
Capacitor_KW_Lbl = Label(Tab4, text='Enter System KW (Kilo Watt)', relief=RIDGE, width=35).grid(row=1, column=1)
Capacitor_KW_Entry = Entry(Tab4, bg='yellow', relief=RIDGE, width=30)
Capacitor_KW_Entry.grid(row=1, column=2)

Capacitor_Present_PF_Lbl = Label(Tab4, text='Enter Present System PF (Power Factor)', relief=RIDGE, width=35).grid(row=2, column=1)
Capacitor_Present_PF_Entry = Entry(Tab4, bg='yellow', relief=RIDGE, width=30)
Capacitor_Present_PF_Entry.grid(row=2, column=2)

Capacitor_Target_PF_Lbl = Label(Tab4, text='Enter Target System PF (Power Factor)', relief=RIDGE, width=35).grid(row=3, column=1)
Capacitor_Target_PF_Entry = Entry(Tab4, bg='yellow', relief=RIDGE, width=30)
Capacitor_Target_PF_Entry.grid(row=3, column=2)

Capacitor_Voltage_Lbl = Label(Tab4, text='Enter System Voltage(in Volts)', relief=RIDGE, width=35).grid(row=4, column=1)
Capacitor_Voltage_Entry = Entry(Tab4, bg='yellow', relief=RIDGE, width=30)
Capacitor_Voltage_Entry.grid(row=4, column=2)
Vector_Diagram_Lbl = Label(Tab4, text='Vector Diagram', relief=RIDGE, width=33,bg='green',font='Calbri 14 bold').grid(row=4, column=3)

Capacitor_Result_KVAR_Lbl = Label(Tab4, text='Capacitor KVAR Required', relief=RIDGE, width=35,font='Calbri 12 bold').grid(row=6, column=1)
Capacitor_Result = Label(Tab4, text='', relief=RIDGE, width=10, bg='orange', fg="darkblue", font='Calbri 12 bold')
Capacitor_Result.grid(row=6, column=2)

Capacitor_Result_MFarad_Lbl = Label(Tab4, text='Capacitor MicroFarad Required', relief=RIDGE, width=35,font='Calbri 12 bold').grid(row=7, column=1)
Capacitor_Result_MFarad = Label(Tab4, text='', relief=RIDGE, width=10, bg='orange', fg="darkblue", font='Calbri 12 bold')
Capacitor_Result_MFarad.grid(row=7, column=2,pady=5)

# CALCULATE BUTTON
Calculate_KVAR = Button(Tab4, text=' Calculate KVAR Required ', font='Calbri 9 bold', bg="skyblue",width=20)
Calculate_KVAR.bind("<ButtonRelease-1>", Calc_KVAR)
Calculate_KVAR.grid(row=9, column=1,pady=5)

# ABOUT BUTTON
Abt = Button(Tab4, text='About this Application', font='Calbri 9 bold', bg="gold",width=20)
Abt.bind("<ButtonRelease-1>", About1)
Abt.grid(row=10, column=1, pady=5)

# QUIT BUTTON
Quitter = Button(Tab4, text='Quit', font='Calbri 9 bold', bg="sandybrown",width=20)
Quitter.bind("<ButtonRelease-1>", Quitt)
Quitter.grid(row=11, column=1, pady=5)

########################################################################################
#############################TAB-3 CABLE QUERY #########################################
########################################################################################

# CSV FILE LOADING FOR LT AND HT CABLES
Cable_Main_Query = pd.read_csv("Data\\HT_Cables_24_8_2023.csv")
Cable_Sizes_Query = Cable_Main_Query["Size"]

Cable_Main_Query_LT = pd.read_csv("Data\\LT_Cables_30_8_2023.csv")
Cable_Sizes_Query_LT = Cable_Main_Query_LT["Size"]

Cable_Heading=Label(Tab3, text="Cable Database Query",font='Calbri 16 bold',bg="green",fg="yellow").grid(row=0,column=2)
Cable_Results_Lbl = Label(Tab3, text='Select Cable Size >>', relief=RIDGE, width=35,bg="yellow",fg="blue",font='Calbri 14 bold').grid(row=4, column=1)

# SELECTING OF CABLE TYPE HT OR LT
Cable_Type_Selection = tkinter.StringVar()
Cable_Type_Selection.set('HT Cable')
Cable_Type_Selection_Radio_Btn = Radiobutton(Tab3, text='LT Cable', variable=Cable_Type_Selection, value='LT Cable',command=Cable_Type_Sel,font='Calbri 14 bold')
Cable_Type_Selection_Radio_Btn.grid(row=1,column=1)
Cable_Type_Selection_Radio_Btn = Radiobutton(Tab3, text='HT Cable', variable=Cable_Type_Selection, value='HT Cable',command=Cable_Type_Sel,font='Calbri 14 bold')
Cable_Type_Selection_Radio_Btn.grid(row=1,column=2)

##### CABLE TYPE #####
Search_Cable_Size = tkinter.StringVar()
Cable_Size_Query_Combo = ttk.Combobox(Tab3, width=60, state="readonly", values=Cable_Sizes_Query.tolist(),
                             textvariable=Search_Cable_Size,font=("Calbri",11))
Cable_Size_Query_Combo.grid(row=4, column=2)
Cable_Size_Query_Combo.bind('<<ComboboxSelected>>', Cable_Selected_Query)

#FEILD-1
Cable_Feild_1_Lbl = Label(Tab3, text='...', relief=RIDGE, width=35,font='Calbri 14 bold')
Cable_Feild_1_Lbl.grid(row=6, column=1)
Cable_Feild_1_Data = Label(Tab3, text='...', relief=RIDGE, width=35,bg='orange', fg="darkblue", font='Calbri 14 bold')
Cable_Feild_1_Data.grid(row=6, column=2)

#FEILD-2
Cable_Feild_2_Lbl = Label(Tab3, text='...', relief=RIDGE, width=35,font='Calbri 14 bold')
Cable_Feild_2_Lbl.grid(row=8, column=1)
Cable_Feild_2_Data = Label(Tab3, text='...', relief=RIDGE, width=35,bg='orange', fg="darkblue", font='Calbri 14 bold')
Cable_Feild_2_Data.grid(row=8, column=2)

#FEILD-3
Cable_Feild_3_Lbl = Label(Tab3, text='...', relief=RIDGE, width=35,font='Calbri 14 bold')
Cable_Feild_3_Lbl.grid(row=10, column=1)
Cable_Feild_3_Data = Label(Tab3, text='...', relief=RIDGE, width=35,bg='orange', fg="darkblue", font='Calbri 14 bold')
Cable_Feild_3_Data.grid(row=10, column=2)

#FEILD-4
Cable_Feild_4_Lbl = Label(Tab3, text='...', relief=RIDGE, width=35,font='Calbri 14 bold')
Cable_Feild_4_Lbl.grid(row=12, column=1)
Cable_Feild_4_Data = Label(Tab3, text='...', relief=RIDGE, width=35,bg='orange', fg="darkblue", font='Calbri 14 bold')
Cable_Feild_4_Data.grid(row=12, column=2)

#FEILD-5
Cable_Feild_5_Lbl = Label(Tab3, text='...', relief=RIDGE, width=35,font='Calbri 14 bold')
Cable_Feild_5_Lbl.grid(row=14, column=1)
Cable_Feild_5_Data = Label(Tab3, text='...', relief=RIDGE, width=35,bg='orange', fg="darkblue", font='Calbri 14 bold')
Cable_Feild_5_Data.grid(row=14, column=2)

#FEILD-6
Cable_Feild_6_Lbl = Label(Tab3, text='...', relief=RIDGE, width=35,font='Calbri 14 bold')
Cable_Feild_6_Lbl.grid(row=16, column=1)
Cable_Feild_6_Data = Label(Tab3, text='...', relief=RIDGE, width=35,bg='orange', fg="darkblue", font='Calbri 14 bold')
Cable_Feild_6_Data.grid(row=16, column=2)

#FEILD-7
Cable_Feild_7_Lbl = Label(Tab3, text='...', relief=RIDGE, width=35,font='Calbri 14 bold')
Cable_Feild_7_Lbl.grid(row=18, column=1)
Cable_Feild_7_Data = Label(Tab3, text='...', relief=RIDGE, width=35,bg='orange', fg="darkblue", font='Calbri 14 bold')
Cable_Feild_7_Data.grid(row=18, column=2)

#FEILD_8
Cable_Feild_8_Lbl = Label(Tab3, text='...', relief=RIDGE, width=35,font='Calbri 14 bold')
Cable_Feild_8_Lbl.grid(row=20, column=1)
Cable_Feild_8_Data = Label(Tab3, text='...', relief=RIDGE, width=35,bg='orange', fg="darkblue", font='Calbri 14 bold')
Cable_Feild_8_Data.grid(row=20, column=2)

#FEILD-9
Cable_Feild_9_Lbl = Label(Tab3, text='...', relief=RIDGE, width=35,font='Calbri 14 bold')
Cable_Feild_9_Lbl.grid(row=21, column=1)
Cable_Feild_9_Data = Label(Tab3, text='...', relief=RIDGE, width=35,bg='orange', fg="darkblue", font='Calbri 14 bold')
Cable_Feild_9_Data.grid(row=21, column=2,pady=25)

#FEILD-10
Cable_Feild_10_Lbl = Label(Tab3, text='...', relief=RIDGE, width=35,font='Calbri 14 bold')
Cable_Feild_10_Lbl.grid(row=22, column=1)
Cable_Feild_10_Data = Label(Tab3, text='...', relief=RIDGE, width=35,bg='orange', fg="darkblue", font='Calbri 14 bold')
Cable_Feild_10_Data.grid(row=22, column=2,pady=25)

#FEILD-11
Cable_Feild_11_Lbl = Label(Tab3, text='...', relief=RIDGE, width=35,font='Calbri 14 bold')
Cable_Feild_11_Lbl.grid(row=23, column=1)
Cable_Feild_11_Data = Label(Tab3, text='...', relief=RIDGE, width=35,bg='orange', fg="darkblue", font='Calbri 14 bold')
Cable_Feild_11_Data.grid(row=23, column=2,pady=25)

#COST /M
Cable_Gap_Lbl = Label(Tab3, text='', width=2).grid(row=0, column=3)

#IMAGE LOADING
Cable_HT_Image = Image.open("Images/HT_Cable_3_Core.PNG")
Cable_HT_Cable_Image = ImageTk.PhotoImage(Cable_HT_Image)

Cable_HT_Lbl = Label(Tab3,borderwidth=2, relief="solid")
Cable_HT_Lbl.grid(row=0, column=4, rowspan=200, padx=5)

#IMAGE LOADING
Cable_LT_Image = Image.open("Images/LT_Cable_3_Core.PNG")
Cable_LT_Cable_Image = ImageTk.PhotoImage(Cable_LT_Image)

# ABOUT BUTTON
Abt = Button(Tab3, text='About this Application', font='Calbri 9 bold', bg="gold",width=20)
Abt.bind("<ButtonRelease-1>", About1)
Abt.grid(row=32, column=1, pady=5)

# QUIT BUTTON
Quitter = Button(Tab3, text='Quit', font='Calbri 9 bold', bg="sandybrown",width=20)
Quitter.bind("<ButtonRelease-1>", Quitt)
Quitter.grid(row=33, column=1, pady=5)

###############################################################################################
#################################TAB-5 PROTECTION RELAY QUERY #################################
###############################################################################################

# CSV FILE LOADING FOR LT AND HT CABLES
Protect_Main_Query = pd.read_csv("Data\\Protection_Relays_31_8_2023.csv")
Protect_Type_Query = Protect_Main_Query["Relay Number"]
Protect_Desc_Query = Protect_Main_Query["Relay Description"]

Protect_Heading=Label(Tab5, text="Protection Relay Database Query",font='Calbri 16 bold',bg="green",fg="yellow").grid(row=0,column=2)
Protect_Lbl_Text = Label(Tab5, text='... >>', relief=RIDGE, width=35,bg="yellow",fg="blue",font='Calbri 14 bold')
Protect_Lbl_Text.grid(row=4, column=1)

##### PROTECTION RELAY COMBO BOX #####
Search_Protect_Number = tkinter.StringVar()
Protect_Number_Query_Combo = ttk.Combobox(Tab5, width=60, state="readonly", values=Protect_Type_Query.tolist(),
                             textvariable=Search_Protect_Number,font=("Calbri",11))
Protect_Number_Query_Combo.grid(row=4, column=2,pady=15)
Protect_Number_Query_Combo.bind('<<ComboboxSelected>>',Protect_Selected_Query)

# SELECTING OF RELAY NUMBER OR DESCRIPTION
Prot_Type_Selection = tkinter.StringVar()
Prot_Type_Selection.set('Relay Number')
Prot_Type_Selection_Radio_Btn = Radiobutton(Tab5, text='Relay Number', variable=Prot_Type_Selection, value='Relay Number',command=Prot_Type_Sel,font='Calbri 14 bold')
Prot_Type_Selection_Radio_Btn.grid(row=5,column=1)
Prot_Type_Selection_Radio_Btn = Radiobutton(Tab5, text='Relay Description', variable=Prot_Type_Selection, value='Relay Description',command=Prot_Type_Sel,font='Calbri 14 bold')
Prot_Type_Selection_Radio_Btn.grid(row=5,column=2)

#FEILD-1
Protect_Field_1_Lbl = Label(Tab5, text='...', relief=RIDGE, width=35,font='Calbri 14 bold')
Protect_Field_1_Lbl.grid(row=6, column=1)
Protect_Field_1_Data = Label(Tab5, text='...', relief=RIDGE, width=35,bg='orange', fg="darkblue", font='Calbri 14 bold')
Protect_Field_1_Data.grid(row=6, column=2)

#FEILD-2
Protect_Field_2_Lbl = Label(Tab5, text='...', relief=RIDGE, width=35,font='Calbri 14 bold')
Protect_Field_2_Lbl.grid(row=8, column=1)
Protect_Field_2_Data = Label(Tab5, text='...', relief=RIDGE, width=35,bg='orange', fg="darkblue", font='Calbri 14 bold')
Protect_Field_2_Data.grid(row=8, column=2)

#FEILD-3
Protect_Field_3_Lbl = Label(Tab5, text='...', relief=RIDGE, width=35,font='Calbri 14 bold')
Protect_Field_3_Lbl.grid(row=10, column=1)
Protect_Field_3_Data = Label(Tab5, text='...', relief=RIDGE, width=35,bg='orange', fg="darkblue", font='Calbri 14 bold')
Protect_Field_3_Data.grid(row=10, column=2)

#IMAGE LOADING
Protect_Image = Image.open("Images/Protect_Relay_General.png")
Protect_Relay_Image = ImageTk.PhotoImage(Protect_Image)

Protect_Lbl = Label(Tab5,borderwidth=2, relief="solid")
Protect_Lbl.grid(row=0, column=4, rowspan=200, padx=5)


# ABOUT BUTTON
Abt_Protect = Button(Tab5, text='About this Application', font='Calbri 9 bold', bg="gold",width=20)
Abt_Protect.bind("<ButtonRelease-1>", About1)
Abt_Protect.grid(row=20, column=1, pady=5)

# QUIT BUTTON
Quitter_Protect = Button(Tab5, text='Quit', font='Calbri 9 bold', bg="sandybrown",width=20)
Quitter_Protect.bind("<ButtonRelease-1>", Quitt)
Quitter_Protect.grid(row=22, column=1, pady=5)

########################################################################################
#############################TAB-2 LIGHTNING PROTECTION RISK ANALYSIS###################
########################################################################################

Lightning_Heading=Label(Tab2, text="Lightning Protection Risk Analysis(EN 62305-2)",font='Calbri 12 bold',bg="blue",fg="yellow").grid(row=0,column=2)

#PROJECT NAME
Lightning_Proj_Lbl = Label(Tab2, text='Project Name:', relief=RIDGE, width=25,bg='coral').grid(row=4, column=1)
Lightning_Proj_Data_Entry = Entry(Tab2, bg='yellow', relief=RIDGE, width=25)
#Lightning_Proj_Data = Label(Tab2, text='...', relief=RIDGE, width=25,bg='yellow' )
Lightning_Proj_Data_Entry.grid(row=4, column=2)

#DATE OF PROJECT
Lightning_Date_Lbl = Label(Tab2, text='Project Date:', relief=RIDGE, width=25,bg='coral').grid(row=5, column=1)
Lightning_Date_Data = Label(Tab2, text='...', relief=RIDGE, width=25,bg='yellow')
Lightning_Date_Data.grid(row=5, column=2)
Lightning_Date_Data.config(text=str(Hr) + ":" + str(Min) + ":" + str(Sec))

# BUILDING DIMENSIONS
Lightning_Building_Lbl = Label(Tab2, text='Building Dimensions', relief=RIDGE, width=25,font='Calbri 10 bold',bg='lightblue').grid(row=6, column=1)
Lightning_Building_Length_Lbl = Label(Tab2, text=' Building Length in m  ', relief=RIDGE, width=25,bg='gold').grid(row=7, column=1)
Lightning_Building_Length_Entry = Entry(Tab2, bg='yellow', relief=RIDGE, width=25)
Lightning_Building_Length_Entry.grid(row=7, column=2)

Lightning_Building_Bredth_Lbl = Label(Tab2, text=' Building Bredth in m  ', relief=RIDGE, width=25,bg='gold').grid(row=8, column=1)
Lightning_Building_Bredth_Entry = Entry(Tab2, bg='yellow', relief=RIDGE, width=25)
Lightning_Building_Bredth_Entry.grid(row=8, column=2)

Lightning_Building_Height_Lbl = Label(Tab2, text=' Building Height in m  ', relief=RIDGE, width=25,bg='gold').grid(row=9, column=1)
Lightning_Building_Height_Entry = Entry(Tab2, bg='yellow', relief=RIDGE, width=25)
Lightning_Building_Height_Entry.grid(row=9, column=2)

#############################CALCULATING AD ##############################################
Lightning_Building_Ad_Lbl = Label(Tab2, text=' Calculated Ad in m2  ', relief=RIDGE, width=25,bg='gold').grid(row=10, column=1)
Lightning_Building_Ad_Calculated = Label(Tab2, text='...', relief=RIDGE, width=25)
Lightning_Building_Ad_Calculated.grid(row=10, column=2)

#############################LIGHNING STRIKE DENSITY Ng####################################
Lightning_Strike_Density_Lbl = Label(Tab2, text=' Lightning Strike Density in St/km2/Yr ', relief=RIDGE, width=35,bg='pink').grid(row=11, column=1)
Lightning_Strike_Density_Entry = Entry(Tab2, bg='yellow', relief=RIDGE, width=25)
Lightning_Strike_Density_Entry.grid(row=11, column=2)

#####SELECTING DANGER FOR PEOPLE (h)#######################################################
Lightning_Danger_People_Lbl = Label(Tab2, text=' Danger to People (h) ', relief=RIDGE, width=35,bg='lightgreen').grid(row=12, column=1)
Lightning_Danger_People = tkinter.StringVar()
Lightning_Danger_People_Combo = ttk.Combobox(Tab2, width=55, state="readonly", textvariable=Lightning_Danger_People,font=("Calbri",9))
Lightning_Danger_People_Combo.grid(row=12, column=2)
Lightning_Danger_People_Combo['values'] = ('No Particular Danger',
                                           'Low Panic (<=2 Floors,<100 Persons',
                                           'Medium Risk (<1000 persons'
                                           'Difficult to Evacuate (Disabled Persons,hospitals)',
                                           'High Risk of Panic (>1000 Persons)',
                                           'Hazard to Surroundings and Environment ',
                                           'Contamination of Surroundings OR Environment'
                                           )
Lightning_Danger_People_Combo.bind('<<ComboboxSelected>>',Lightning_Danger_People_Sub)

#########################SELECTING OCCUPANCY COEFFECIENT Lf1#######################################################
Lightning_Occupancy_Lbl = Label(Tab2, text=' Occupancy Coeffecient (Lf1) ', relief=RIDGE, width=35,bg='lightgreen').grid(row=13, column=1)
Lightning_Occupancy = tkinter.StringVar()
Lightning_Occupancy_Combo = ttk.Combobox(Tab2, width=55, state="readonly", textvariable=Lightning_Occupancy,font=("Calbri",9))
Lightning_Occupancy_Combo.grid(row=13, column=2)
Lightning_Occupancy_Combo['values'] = ('Structure UnOccupied',
                                       'Structure Normally Occupied',
                                      )
Lightning_Occupancy_Combo.bind('<<ComboboxSelected>>',Lightning_Occupancy_Sub)

#########################SELECTING RELATIVE LOACTION OF SITE Cd#####################################################
Lightning_Rel_Loc_Lbl = Label(Tab2, text=' Relative Location Of Site (Cd) ', relief=RIDGE, width=35,bg='lightgreen').grid(row=14, column=1)
Lightning_Rel_Loc = tkinter.StringVar()
Lightning_Rel_Loc_Combo = ttk.Combobox(Tab2, width=55, state="readonly", textvariable=Lightning_Rel_Loc,font=("Calbri",9))
Lightning_Rel_Loc_Combo.grid(row=14, column=2)
Lightning_Rel_Loc_Combo['values'] = ('Structure Surrounded by Higher Objects or Trees',
                                       'Structure Surrounded by similar or Lower Objects',
                                       'Isolated Structure- No other Objects nearby',
                                       'Isolated Structure on top of Hill or Hillock'
                                      )
Lightning_Rel_Loc_Combo.bind('<<ComboboxSelected>>',Lightning_Rel_Loc_Sub)

#########################SELECTING FIRE RISK rf #################################################################
Lightning_Fire_Risk_Loc_Lbl = Label(Tab2, text=' Fire Risk Of Site (rf) ', relief=RIDGE, width=35,bg='lightgreen').grid(row=15, column=1)
Lightning_Fire_Risk = tkinter.StringVar()
Lightning_Fire_Risk_Combo = ttk.Combobox(Tab2, width=55, state="readonly", textvariable=Lightning_Fire_Risk,font=("Calbri",9))
Lightning_Fire_Risk_Combo.grid(row=15, column=2)
Lightning_Fire_Risk_Combo['values'] = ('Explosion Risk',
                                       'High Risk',
                                       'Ordinary Risk',
                                       'Low Risk'
                                      )
Lightning_Fire_Risk_Combo.bind('<<ComboboxSelected>>',Lightning_Fire_Risk_Sub)

#########################SELECTING ASSOCIATED SERVICES Lf2 ################################################################
Lightning_Ass_Services_Lbl = Label(Tab2, text=' Assocaited Services (Lf2) ', relief=RIDGE, width=35,bg='lightgreen').grid(row=16, column=1)
Lightning_Ass_Services = tkinter.StringVar()
Lightning_Ass_Services_Combo = ttk.Combobox(Tab2, width=55, state="readonly", textvariable=Lightning_Ass_Services,font=("Calbri",9))
Lightning_Ass_Services_Combo.grid(row=16, column=2)
Lightning_Ass_Services_Combo['values'] = ('No Associated Services',
                                       'Gas and Water',
                                       'TV,Communication,Electricity and Radio'
                                         )
Lightning_Ass_Services_Combo.bind('<<ComboboxSelected>>',Lightning_Ass_Services_Sub)

#########################SELECTING ELECTRICAL LINES Ai ################################################################
Lightning_Electrical_Lines_Lbl = Label(Tab2, text=' Presence of Electrical Lines (Ai) ', relief=RIDGE, width=35,bg='lightgreen').grid(row=17, column=1)
Lightning_Electrical_Lines = tkinter.StringVar()
Lightning_Electrical_Lines_Combo = ttk.Combobox(Tab2, width=55, state="readonly", textvariable=Lightning_Electrical_Lines,font=("Calbri",9))
Lightning_Electrical_Lines_Combo.grid(row=17, column=2)
Lightning_Electrical_Lines_Combo['values'] = ('No Electrical Lines',
                                              'Aerial Lines',
                                              'Underground Lines'
                                             )
Lightning_Electrical_Lines_Combo.bind('<<ComboboxSelected>>',Lightning_Electrical_Lines_Sub)

#########################SELECTING LIGHTNING PROTECTION LEVEL Pd ################################################################
Lightning_Protection_Level_Lbl = Label(Tab2, text=' Lightning Protection Level (Pd) ', relief=RIDGE, width=35,bg='lightgreen').grid(row=18, column=1)
Lightning_Protection_Level = tkinter.StringVar()
Lightning_Protection_Level_Combo = ttk.Combobox(Tab2, width=55, state="readonly", textvariable=Lightning_Protection_Level,font=("Calbri",9))
Lightning_Protection_Level_Combo.grid(row=18, column=2)
Lightning_Protection_Level_Combo['values'] = ('None',
                                              'Protection Level-IV',
                                              'Protection Level-III',
                                              'Protection Level-II',
                                              'Protection Level-I'
                                             )
Lightning_Protection_Level_Combo.bind('<<ComboboxSelected>>',Lightning_Protection_Level_Sub)

#########################SELECTING SURGE ARRESTORS Pi ################################################################
Lightning_Surge_Arrestor_Lbl = Label(Tab2, text=' Instalation Of Surge Arrestors(Pi) ', relief=RIDGE, width=35,bg='lightgreen').grid(row=19, column=1)
Lightning_Surge_Arrestor = tkinter.StringVar()
Lightning_Surge_Arrestor_Combo = ttk.Combobox(Tab2, width=55, state="readonly", textvariable=Lightning_Surge_Arrestor,font=("Calbri",9))
Lightning_Surge_Arrestor_Combo.grid(row=19, column=2)
Lightning_Surge_Arrestor_Combo['values'] = (  'None',
                                              'Neccesary I imp >=12.5 KA',
                                              'Neccesary I imp = 75/(mxn)',
                                              'Neccesary I imp =100/(mxn)'
                                             )
Lightning_Surge_Arrestor_Combo.bind('<<ComboboxSelected>>',Lightning_Surge_Arrestor_Sub)

#################################RESULTS HEADING##################################################
Lightning_Results_Gap=Label(Tab2).grid(row=20,column=1,pady=5)
Lightning_Results=Label(Tab2, text="Risk Analysis Results(EN 62305-2)",font='Calbri 12 bold',bg="lime",fg="red",relief=RIDGE).grid(row=21,column=1)
Lightning_Heading_RT1=Label(Tab2, text="RT1 - Loss to Human Life (Max 1 in 100000)",font='Calbri 12 bold',bg="khaki",fg="red",width=38,relief=RIDGE).grid(row=22,column=1)
Lightning_Heading_RT2=Label(Tab2, text="RT2 - Loss to Services to Public (Max 1 in 1000)",font='Calbri 12 bold',bg="khaki",fg="red",width=38,relief=RIDGE).grid(row=23,column=1)
Lightning_Heading_RT3=Label(Tab2, text="RT3 - Loss to Cultural Heritage (Max 1 in 1000)",font='Calbri 12 bold',bg="khaki",fg="red",width=38,relief=RIDGE).grid(row=24,column=1)

######################RESULTS RT1,RT2,RT3#########################################################
Lightning_RT1_Result = Label(Tab2, text='...', relief=RIDGE, width=25)
Lightning_RT1_Result.grid(row=22, column=2)

Lightning_RT2_Result = Label(Tab2, text='...', relief=RIDGE, width=25)
Lightning_RT2_Result.grid(row=23, column=2)

Lightning_RT3_Result = Label(Tab2, text='...', relief=RIDGE, width=25)
Lightning_RT3_Result.grid(row=24, column=2)

Lightning_Results_Gap1=Label(Tab2).grid(row=25,column=1,pady=5)

Lightning_Help = Label(Tab2, text='<  Help  >', relief=RIDGE, width=50,bg='pink',fg='blue',font='Calbri 9 bold')
Lightning_Help.grid(row=11, column=3)

Lightning_Help_Box = Label(Tab2, text='....', relief=RIDGE, width=50,bg='tan',fg='blue',font='Calbri 9 bold',wraplength=300)
Lightning_Help_Box.grid(row=12, column=3)

######CALCULATE RISK BUTTON######
Calculate_Lightning_Risk = Button(Tab2, text=' Calculate Lightning Risk ', font='Calbri 9', bg="skyblue",width=20)
Calculate_Lightning_Risk.bind("<ButtonRelease-1>", Calc_Lightning_Risk)
Calculate_Lightning_Risk.grid(row=26, column=1,pady=5)

#####ABOUT BUTTON######
Abt = Button(Tab2, text='About this Application', font='Calbri 9 bold', bg="gold",width=20)
Abt.bind("<ButtonRelease-1>", About1)
Abt.grid(row=27, column=1, pady=5)

######QUIT BUTTON#####
Quitter = Button(Tab2, text='Quit', font='Calbri 9 bold', bg="sandybrown",width=20)
Quitter.bind("<ButtonRelease-1>", Quitt)
Quitter.grid(row=28, column=1, pady=5)

############################################################################################################
###############################TAB-6 ELEVATOR DESIGN #######################################################
############################################################################################################
Excel_Workbook = xlwt.Workbook()
#SETTING OF FONT0 FOR EXCEL
Font0=xlwt.Font()
Font0.name="Calibri"
Font0.bold=True
Font0.height=600

#SETTING FONT1
Font1=xlwt.Font()
Font1.name="Calibri"
Font1.bold=True
Font1.height=200

#SETTING OF FONT 2,3 FOR EXCEL)
Font2=xlwt.Font()
Font2.name="Calibri"
Font2.height=180

Font3=xlwt.Font()
Font3.name="Calibri"
Font3.height=400

Font4=xlwt.Font()
Font4.name="Calibri"
Font4.bold=True
Font4.height=200


#EXCEL STYLE SETTINGS
Style0=xlwt.XFStyle()
pattern = xlwt.Pattern()
pattern.pattern = xlwt.Pattern.SOLID_PATTERN
pattern.pattern_fore_colour = xlwt.Style.colour_map['light_green']
Style0.pattern = pattern
Style0.font=Font0

Style3=xlwt.XFStyle()
pattern = xlwt.Pattern()
pattern.pattern = xlwt.Pattern.SOLID_PATTERN
pattern.pattern_fore_colour = xlwt.Style.colour_map['green']
Style3.pattern = pattern
Style3.font=Font2

Style4=xlwt.XFStyle()
pattern = xlwt.Pattern()
pattern.pattern = xlwt.Pattern.SOLID_PATTERN
pattern.pattern_fore_colour = xlwt.Style.colour_map['yellow']
Style4.pattern = pattern
Style4.font=Font4
Style4.alignment.wrap=1

Style1=xlwt.XFStyle()
Style1.font=Font1
Style2=xlwt.XFStyle()
Style2.font=Font2


Col_main_lbl1=Label(Tab6,text="Category:Basic Project Details",bg="lightblue",fg="red",font=16, relief=RIDGE,width=55).grid(row=0,column=0)

Lbl1=Label(Tab6,text='Name of Project:', relief=RIDGE,width=55).grid(row=1,column=0)
Ent1 = Entry(Tab6,bg='yellow', relief=SUNKEN, width=35)
Ent1.grid(row=1,column=1,padx=0,pady=0)

Lbl2=Label(Tab6,text='Total Population:', relief=RIDGE,width=55).grid(row=2,column=0)
Ent2= Entry(Tab6,bg='yellow', relief=SUNKEN, width=15)
Ent2.grid(row=2,column=1,padx=0,pady=0)


#CATEGORY BUILDING DATA

Col_main_lbl2=Label(Tab6,text="Category: Building Data",bg="lightblue",fg="red",font=16, relief=RIDGE,width=55).grid(row=3,column=0)
Lbl3=Label(Tab6,text='Number of Floors(N) above Terminal Floor:', relief=RIDGE,width=55).grid(row=5,column=0)
Ent3 = Entry(Tab6,bg='yellow', relief=SUNKEN, width=15)
Ent3.grid(row=5,column=1)

#NEWLY ADDED NUMBER OF STOPS
Lbl16=Label(Tab6,text='Number of Stops(S):', relief=RIDGE,width=55).grid(row=6,column=0)
Ent16 = Entry(Tab6,bg='yellow', relief=SUNKEN, width=15)
Ent16.grid(row=6,column=1)


Lbl4=Label(Tab6,text='Floor Height in Meters:', relief=RIDGE,width=55).grid(row=7,column=0)
Ent4 = Entry(Tab6,bg='yellow', relief=SUNKEN, width=15)
Ent4.grid(row=7,column=1)


#CATEGORY ELEVATOR OPERATION DATA
Col_main_lbl3=Label(Tab6,text="Category: Elevator Operation Data",bg="lightblue",fg="red",font=16, relief=RIDGE,width=55).grid(row=8,column=0)
Lbl5=Label(Tab6,text='Door Open Time in Secs:(to)', relief=RIDGE,width=55).grid(row=9,column=0)
Ent5 = Entry(Tab6,bg='yellow', relief=SUNKEN, width=15)
Ent5.grid(row=9,column=1)

Lbl6=Label(Tab6,text='Door Close Time in Secs:(tc)', relief=RIDGE,width=55).grid(row=10,column=0)
Ent6 = Entry(Tab6,bg='yellow', relief=SUNKEN, width=15)
Ent6.grid(row=10,column=1)

Lbl7=Label(Tab6,text='Passenger Transfer Time (tp)n Secs:', relief=RIDGE,width=55).grid(row=11,column=0)
Ent7 = Entry(Tab6,bg='yellow', relief=SUNKEN, width=15)
Ent7.grid(row=11,column=1)

Lbl8=Label(Tab6,text='Extra Time in Secs:(Enter 0 if none)', relief=RIDGE,width=55).grid(row=12,column=0)
Ent8 = Entry(Tab6,bg='yellow', relief=SUNKEN, width=15)
Ent8.grid(row=12,column=1)

Lbl9=Label(Tab6,text='Building Type Selection', relief=RIDGE,width=55).grid(row=14,column=0)

building_type = ttk.Combobox(Tab6, width=27,state="readonly")
building_type.grid(row=14,column=1)
building_type['values'] = (' Multi Tenant Commercial',
                          ' Single Tenant Commercial',
                          ' Residential'
                          )
#MINIMUM ELEVATOR SPEED
Lbl14=Label(Tab6,text='Minimum Elevator Speed(m/s)', relief=RIDGE,width=55).grid(row=15,column=0)
Min_Elev_Speed = ttk.Combobox(Tab6, width=27,state="readonly")
Min_Elev_Speed.grid(row=15,column=1)
Min_Elev_Speed['values'] = ('0.75',
                          '1.0',
                          '1.5',
                          '1.75',
                          '2.0',
                          '2.5',
                          '3.0',
                          '4.0'
                          )

#WAITING INTERVAL CATEGORY
Lbl15=Label(Tab6,text='Waiting Interval Category', relief=RIDGE,width=55).grid(row=16,column=0)
Wait_Int_Cat = ttk.Combobox(Tab6, width=27,state="readonly")
Wait_Int_Cat.grid(row=16,column=1)
Wait_Int_Cat['values'] = ('Excellent',
                          'Very Good',
                          'Good',
                          'Fair',
                          'Poor',
                          'Unsatisfactory'
                          )
#NEWLY ADDED ELEVATOR CAR CAPACITY FACTOR
Lbl17=Label(Tab6,text='Elevator Capacity Factor:(Typically 0.8)', relief=RIDGE,width=55).grid(row=17,column=0)
Ent17 = Entry(Tab6,bg='yellow', relief=SUNKEN, width=15)
Ent17.grid(row=17,column=1)

#NEWLY ADDED ACCELERATION
Lbl18=Label(Tab6,text='Elevator Acceleration in m/s2:', relief=RIDGE,width=55).grid(row=18,column=0)
Ent18 = Entry(Tab6,bg='yellow', relief=SUNKEN, width=15)
Ent18.grid(row=18,column=1)

#CATEGORY BASIC RESULTS
Col_main_lbl4=Label(Tab6,text="Category: Basic Elevator Results:",bg="lightblue",fg="red",font=16, relief=RIDGE,width=55).grid(row=19,column=0)
Lbl9=Label(Tab6,text='Highest Reversal Floor (H):', relief=RIDGE,width=55,bg="lightgreen").grid(row=20,column=0)
Ent9=Label(Tab6,text='', relief=RIDGE,width=13)
Ent9.grid(row=20,column=1)

Lbl10=Label(Tab6,text='Average Number of Stops (S) :', relief=RIDGE,width=55,bg="lightgreen").grid(row=21,column=0)
Ent10=Label(Tab6,text='', relief=RIDGE,width=13)
Ent10.grid(row=21,column=1)

Lbl11=Label(Tab6,text='Total Height of Travel :', relief=RIDGE,width=55,bg="lightgreen").grid(row=22,column=0)
Ent11=Label(Tab6,text='', relief=RIDGE,width=13)
Ent11.grid(row=22,column=1)

Lbl12=Label(Tab6,text='Single Floor Transit Time (tv) in Secs', relief=RIDGE,width=55,bg="lightgreen").grid(row=23,column=0)
Ent12=Label(Tab6,text='', relief=RIDGE,width=13)
Ent12.grid(row=23,column=1,padx=10, pady=1,)

Lbl13=Label(Tab6,text='Total Door Open Close and Flight Time (T) in Secs', relief=RIDGE,width=55,bg="lightgreen").grid(row=24,column=0)
Ent13=Label(Tab6,text='', relief=RIDGE,width=13)
Ent13.grid(row=24,column=1,padx=10, pady=1)

Calculate_Elevator = Button(Tab6,text='Calculate Elevator Data',width=20,bg='skyblue')
Calculate_Elevator.bind("<ButtonRelease-1>", Calc_Elevator)
Calculate_Elevator.grid(row=1,column=3,padx=5, pady=2)

About_Elevator = Button(Tab6,text= 'About',width=20,bg='gold')
About_Elevator.bind("<ButtonRelease-1>", About1)
About_Elevator.grid(row=2,column=3,padx=5, pady=2)

Quit1 = Button(Tab6,text='Quit',width=20,bg='sandybrown')
Quit1.bind("<ButtonRelease-1>", Quitt)
Quit1.grid(row=3,column=3,padx=5, pady=2)

Excel_Check=tkinter.IntVar()
Excel_Display=Checkbutton(Tab6,text="Open Results in Excel",variable=Excel_Check,bg='lightgreen')
Excel_Display.grid(row=4,column=3,padx=5,pady= 2)

###################################################################################################
##################################TAB-7 UPDATES####################################################
###################################################################################################
Updates_Btn = Button(Tab7,text= 'View Updates',width=20,bg='gold')
Updates_Btn.bind("<ButtonRelease-1>", Updates)
Updates_Btn.grid(row=23,column=3,padx=150, pady=50)
Updates_Btn.config(width= 40,height=5)

Quit1 = Button(Tab7,text='Quit',width=20,bg='sandybrown')
Quit1.bind("<ButtonRelease-1>", Quitt)
Quit1.grid(row=23,column=4,padx=150, pady=50)
Quit1.config(width= 40,height=5)

root.mainloop()

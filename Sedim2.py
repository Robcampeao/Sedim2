# -*- coding: cp1252 -*-

##SEDIM 2.0
##Copyright (C) 2017 Roberta Campeao
##
##Based on: SEDIM 1.0 1996 Monica da Hora
##          SEDDISCH 1989 Stevens and Yang
##
##This program is free software: you can redistribute it and/or modify
##it under the terms of the GNU General Public License as published by
##the Free Software Foundation, either version 3 of the License, or
##(at your option) any later version.
##
##This program is distributed in the hope that it will be useful,
##but WITHOUT ANY WARRANTY; without even the implied warranty of
##MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
##GNU General Public License for more details.
##
##You should have received a copy of the GNU General Public License
##along with this program.  If not, see <http://www.gnu.org/licenses/>.
##
##Contact information:  Roberta Campeao - robertacampeao@gmail.com
##                      Monica da Hora - dahora@vm.uff.br

from PyQt4 import QtGui, QtCore # Import the PyQt4 module we'll need
import sys # We need sys so that we can pass argv to QApplication
import inspect
import xlsxwriter
import os

import design # This file holds our MainWindow and all design related things
              # it also keeps events etc that we defined in Qt Designer

import math

def FNL(x):
    return math.log(x,10)

def WriteResultsFaixa(FormulaName, ResultFormula, ib, Fgranlabel, row, sheetSInt, sheetSIng, faixa):
    if faixa == 1:
        sheetSInt.write(row, 0, FormulaName)
        sheetSIng.write(row, 0, FormulaName)
        row+=1
        sheetSInt.write(row, 0, 'Faixa Gran (mm)')
        sheetSIng.write(row, 0, 'Faixa Gran (mm)')

        sheetSInt.write(row, 1, 'Fracao (%)')
        sheetSIng.write(row, 1, 'Fracao (%)')
        
        sheetSInt.write(row, 2, 'C (mg/l)')
        sheetSIng.write(row, 2, 'C (mg/l)')
        
        #sheetSInt.write(row, 3, 'Qs (ton/dia/m)')
        sheetSInt.write(row, 3, 'Qst (t/dia)')
        
        #sheetSIng.write(row, 3, 'Qs (lb/s/ft)')
        sheetSIng.write(row, 3, 'Qst (ton/dia)')
        row+=1

        for ci, qst, i, fg in zip(ResultFormula[3], ResultFormula[4], ib, Fgranlabel):
            sheetSInt.write(row, 0, fg)
            sheetSInt.write(row, 1, i*100)
            sheetSInt.write(row, 2, ci)
            #sheetSInt.write(row, 3, qst*43.2/0.3048)

            sheetSIng.write(row, 0, fg)
            sheetSIng.write(row, 1, i*100)
            sheetSIng.write(row, 2, ci)
            #sheetSIng.write(row, 3, qst)
            row+=1

        sheetSInt.write(row, 0, "Total")
        sheetSInt.write(row, 1, 100.0*sum(ib))    
        sheetSInt.write(row, 2, ResultFormula[0])
        #sheetSInt.write(row, 3, ResultFormula[1]*43.2/0.3048)
        sheetSInt.write(row, 3, ResultFormula[2])

        
        sheetSIng.write(row, 0, "Total")
        sheetSIng.write(row, 1, "100")
        sheetSIng.write(row, 2, ResultFormula[0])
        #sheetSIng.write(row, 3, ResultFormula[1])
        sheetSIng.write(row, 3, ResultFormula[2])
        row+=2
    else:
        sheetSInt.write(row, 0, FormulaName)
        sheetSIng.write(row, 0, FormulaName)
        row+=1
        sheetSInt.write(row, 0, 'C (mg/l)')
        sheetSIng.write(row, 0, 'C (mg/l)')
        
        #sheetSInt.write(row, 1, 'Qs (ton/dia/m)')
        sheetSInt.write(row, 1, 'Qst (t/dia)')
        
        #sheetSIng.write(row, 1, 'Qs (lb/s/ft)')
        sheetSIng.write(row, 1, 'Qst (ton/dia)')
        row+=1
        
        sheetSInt.write(row, 0, ResultFormula[0])
        #sheetSInt.write(row, 1, ResultFormula[1]*43.2/0.3048)
        sheetSInt.write(row, 1, ResultFormula[2])

        sheetSIng.write(row, 0, ResultFormula[0])
        #sheetSIng.write(row, 1, ResultFormula[1])
        sheetSIng.write(row, 1, ResultFormula[2])
        row+=2
    return row

def FallVelocity(D, TEMP, AF):
    DFV = D * 304.8
    SF = TEMP / 10.0
    KT = int(SF)+1
    PT = SF - KT + 1
    DL = FNL(DFV)
    M = 0
    while DFV > AF[0][M]:
        M += 1
    M -=1
    CF = FNL(AF[0][M])
    EF = FNL(AF[0][M+1])
    PD = (DL - CF)/ (EF - CF)
    ZF = []
    for L in range(2):
        K = L + KT
        ZF.append((1 - PD) * FNL(AF[K][M]) + PD * FNL(AF[K][M + 1]))
    RF = (1.0 - PT) * ZF[0] + PT * ZF[1]
    FV = 10.0**RF / 30.48
    return FV

def Laursen(DFT, ib, V, SG, g, Y, XNU, U, DF50, W):
    UGS=0
    C = 0    
    COMP1 = 6*XNU
    CList = []
    UGSList = []
    for D, i in zip(DFT[1:], ib[1:]):
        DELTA  =  11.6*XNU/U
        FVI = (math.sqrt(36.064*D**3+COMP1**2)-COMP1)/D
        RV = U/FVI

        RVL = FNL(RV)
        if RV < 0.3:
            FV = 10.718*RV**0.243
        elif RV < 3.0:
            FV = 10.0**(0.855*RVL+0.62*RVL**2+1.2)
        elif RV < 20:
            FV = 4.773*RV**2.304
        elif RV < 200:
            FV = 10.0**(3.764*RVL-0.803*RVL**2+0.147)
        else:
            FV = 9680.5*RV**0.2531            
        
        RY = D/DELTA
        if RY > 0.1:
            Yc = 0.04
        elif RY < 0.03:
            Yc = 0.16
        else:
            Yc = 0.08
        
        F1 =(D/Y)**1.1667
        F2 = V**2.0/(58.0*Yc*D*(SG-1.0)*g)
        F3 =(DF50/Y)**0.3333
        
        CI = 10000.0*i*F1*(F2*F3-1.0)*FV
        
        if CI < 0:
            CI = 0

        UGS = UGS + 0.0000625*CI*Y*V

        CList.append(CI)
        UGSList.append(0.0000625*CI*Y*V)
        
    C = 16000.0 * UGS/(Y*V)
    
    return [C, UGS, UGS*43.2*W, CList, UGSList]

def EngelundeHansen (GMS, V, S, DF50, g, SG, W, Y):
    UGS = 0.05*GMS*V**2*Y**1.5*S**1.5/(DF50*g**0.5*(SG-1)**2.0)
    C = 16000.0 * UGS/(Y*V)
    return [C, UGS, UGS*43.2*W]    

def Colby (DF50, Y, V, CY, CF, W, TEMP):
    D50 = DF50 * 304.8
    
    if D50 <= 0.1 or D50 >=0.8:
        print('Outside range of validity')
        return -1
    
    VC = 0.4673 * (Y ** 0.1) * (D50 ** 0.333)
    DIFF = V * 0.3048 - VC
    B = 2.5
    
    if DIFF >= 1.0:
        B = 1.453*D50**(-0.138)
    
    X = FNL(Y)
    
    N = 0
    while TEMP >= CY[3][N]:
        N = N + 1
    
    F1 = CY[4][N-1] + CY[5][N-1] * X + CY[6][N-1] * X ** 2.0
    F2 = CY[4][N]   + CY[5][N]   * X + CY[6][N]   * X ** 2.0
    AF = F1 + (F2 - F1) * (FNL(TEMP)-FNL(CY[3][N-1])) / (FNL(CY[3][N])-FNL(CY[3][N-1]))
    AF = 10 ** AF
    
    N = 0    
    while D50 > CY[0][N]:
        N = N + 1

    A = CY[2][N-1]*Y**(CY[1][N-1])        
    F1 = A * DIFF ** B * (1.0 + (AF - 1.0) * CF[N-1]) * 0.672
        
    A = CY[2][N]  *Y**(CY[1][N])
    F2 = A * DIFF ** B * (1.0 + (AF - 1.0) * CF[N])   * 0.672

    UGS = FNL(F1) + (FNL(F2)-FNL(F1)) * (FNL(D50)-FNL(CY[0][N-1]))/(FNL(CY[0][N])-FNL(CY[0][N-1]))
    UGS = 10.0**UGS
    
    C = 16000.0 * UGS/(Y*V)
    
    return [C, UGS, UGS*43.2*W]

def AckersWhite(D, g, SG, XNU, U, V, Y, W):
    DGR = D * ((g*(SG-1)/XNU**2)**0.3333)
    P = FNL(DGR)
    if DGR <= 60.0:
        AN = 1.0 - 0.56 * P
        AA = 0.23/math.sqrt(DGR)+0.14
        AM = 9.66/DGR + 1.34
        CA = 2.86 * P - P **2 - 3.53
        CA = 10.0 ** CA
    else:
        AN = 0.0
        AA = 0.17
        AM = 1.5
        CA = 0.025
    F1 = U**AN/(math.sqrt(g*D*(SG-1)))
    F2 = (V/(math.sqrt(g)*FNL(10*Y/D)))**(1.0-AN)
    F3 = F1*F2/AA-1.0
    if F3 > 0.0:
        GGR = CA * F3 ** AM
        C = (GGR*D*SG*(V/U)**AN)/Y
        C = C * 10.0**6
        UGS = 0.0000625*C*Y*V
        return [C, UGS, UGS*43.2*W]
    else:
        print("Concentracao menor que zero")
        return -1

def YangD50(SoilType, FV50, DF50, TEMP, AF, D, U, XNU, V, S, Y, W):
    FV = FV50
    if DF50 < 0.0328:
        D = DF50
        FV = FallVelocity(D, TEMP, AF)                
    R = U * DF50 / XNU
    F1 = 2.05
    if R < 70:
        F1 = 0.66 + 2.5/ (FNL(R) - 0.06)
    F2 = FNL(FV*DF50/XNU)
    F3 = FNL(U/FV)
    F4 =  V * S / FV - F1 * S
    if F4 > 0:
        if SoilType == 'Sand':
            C = 5.435 - 0.286 * F2 - 0.457 * F3 + (1.799 - 0.409 * F2 - 0.314 * F3) * FNL(F4) #Sand
        elif SoilType == 'Gravel':
            C = 6.681 - 0.633 * F2 - 4.816 * F3 + (2.784 - 0.305 * F2 - 0.282 * F3) * FNL(F4) #Gravel
        else:
            print('Wrong soil type.')
            C = 0.0
    else:
        C = 0
    C = 10.0**C
    UGS = 0.0000625*C*Y*V
    return [C, UGS, UGS*43.2*W]

def YangFT(SoilType, ib, DFT, TEMP, AF, XNU, U, V, S, Y, W):
    C = 0
    UGS = 0
    CList = []
    UGSList = []
    for i in range(1,11):
        if ib[i] > 0:
            D = DFT[i]
            if D < 0.0328:
                FV = FallVelocity(D, TEMP, AF)
            else:
                COMP1 = 6* XNU
                FV = (math.sqrt(36.064*D**3+COMP1**2)-COMP1)/D
            R = U*D/XNU
            F1 = 2.05
            if R < 70:
                F1 = 0.66 + 2.5 / (FNL(R) - 0.06)
            F2 = FNL(FV*D/XNU)
            F3 = FNL(U/FV)
            CI = 0
            F4 = V*S/FV-F1*S
            if F4 > 0:
                if SoilType == 'Gravel' or (SoilType == 'Mixture' and i>5):
                    CI = 6.681-0.633*F2-4.816*F3+(2.784-0.305*F2-0.282*F3)*FNL(F4)
                else:
                    CI = 5.435-0.286*F2-0.457*F3+(1.799-0.409*F2-0.314*F3)*FNL(F4)
                CI = 10**CI * ib[i]
            C += CI
            UGSI = 0.0000625*CI*Y*V
            UGS += UGSI
            
            CList.append(CI)
            UGSList.append(0.0000625*CI*Y*V)
        else:
            CList.append(0.0)
            UGSList.append(0.0)
            
    return [C, UGS, UGS*43.2*W, CList, UGSList]
            
def Schoklitsch(DFT, ib, S, Y, V, W):
    UGS=0
    CI = 0
    CList = []
    UGSList = []
    for D, i in zip(DFT, ib):
        F1 = 25.0* S**1.5 *Y*V
        F2 = 1.6 * S**0.17
        F3 = math.sqrt(D)
        X = F1/F3 - F2*F3
        if X <= 0:
            UGSI = 0
        else:
            UGSI = i * X
            
        UGS = UGS + UGSI
                
        CI = 16000.0 * UGSI/(Y*V)
        
        CList.append(CI)
        UGSList.append(UGSI)
        
    C = 16000.0 * UGS/(Y*V)
    
    return [C, UGS, UGS*43.2*W, CList, UGSList]

def Kalinske(DFT, ib, AK0, AK1, AK2, AK3, AK4, AK5, Y, V, S, W):
    S1 = 0
    TEMP2 = []
    CList = []
    UGSList = []
    for D, i in zip(DFT, ib):
        TEMP2.append(i/D)
        S1 += i/D
    C = 0
    UGS = 0
    T0 = 62.4 * Y * S
    F1 = 25.28 * T0**0.5 / S1
    for D, i, t in zip(DFT, ib, TEMP2):
        if i > 0:
            T1 = 12 * D
            X = T1 / T0
            F2 = AK0 + AK1 * X + AK2*X**2 + AK3 * X**3 + AK4*X**4 + AK5*X**5
            F2 = 10**F2
            UGSI = F1 * T1 * t * F2
            UGS = UGS + UGSI
            CI = 16000.0 * UGSI / (Y * V)
            C = C + CI
            
            CList.append(CI)
            UGSList.append(0.0000625*CI*Y*V)
        else:
            CList.append(0.0)
            UGSList.append(0.0)
  
    return [C, UGS, UGS*43.2*W, CList, UGSList]
            
def MeyerPetereMuller (DF90, Dm, Y, V, S, W):
    D90 = DF90 * 304.8
    nM = (1.486*Y**(0.667)* S**0.5)/V
    QsQ = 1
    F1= 0.368*QsQ*(D90**(1.0/6.0)/nM)**1.5*Y*S - 0.0698*Dm
    UGS= abs(F1)**1.5
    C = 16000.0 * UGS/(Y*V)
    return [C, UGS, UGS*43.2*W]

def RottnerOld(GMS, SG, g, Y, V, DF50, W):
    termo1 = GMS*math.sqrt((SG-1.0)*g*Y**3.0)
    termo2 = (V/math.sqrt((SG-1.0)*g*Y) * (0.667*(DF50/Y)**1.5 + 0.14) - 0.778*(DF50/Y)**1.5)**3.0
    UGS = termo1*termo2
    C = 16000.0 * UGS/(Y*V)
    return [C, UGS, UGS*43.2*W]

def Rottner(GMS, SG, g, Y, V, DF50, W):
    R = (DF50/Y)**0.667
    F1 = V/(7.286*math.sqrt(Y))
    F2 = 0.667*R + 0.14
    UGS = 1204.8 * Y**1.5 * (F1*F2-0.778*R)**3
    C = 16000.0 * UGS/(Y*V)
    return [C, UGS, UGS*43.2*W]

def Toffaleti(TEMP, Y, S, V, DF65, g, XNU, DFT, DIP, ib, AF, NLD, W):
    TDF = 1.8*TEMP+32.0
    ZV = 0.1198+0.00048*TDF
    CZ = 260.67-0.667*TDF
    YA = Y / 11.24
    YB = Y / 2.5
    CV = 1.0 + ZV
    SI = S * Y * CZ
    U3 = V**3/(XNU*g*S)
    U2 = V/math.sqrt(DF65*g*S)
    F1 = math.log(U3)
    F2 = 4.083*math.log(U2)-3.76
    F3 = 1.864 * F1 - 9.09

    CList = []
    UGSList = []
    
    if F3 < F2:
        U1 = F3
    else:        
        FI = (F2+9.09)/1.864
        FI = (F1 - FI)*0.43429
        if FI >= 1.7:
            U1 = F2 + 0.4
        else:
            F6 = FI * 10.0
            for F5 in range(1,18):
                F1 = F5 - F6
                if F5 < F6:
                    continue  
            J = F5-1
            F1 = 1.0 - F1
            F5 = DIP[J]+F1*(DIP[J+1]-DIP[J])
            U1 = F2 + F5
    
    AM = 10.0 * V / U1  
    PAM = ((XNU * 100000)**0.3333)/AM
    F1 = 100000 * PAM * S * DF65/g
    T = (0.051 + 0.00009*TDF)*1.1
    
    if PAM <= 0.5:
        A = 9.8/(PAM**1.515)
    elif PAM <= 0.66:
        A = 41.0*PAM ** 0.55
    elif PAM <= 0.72:
        A = 228.0 * PAM ** 4.68
    elif PAM <= 1.3:
        A = 49.0
    else:
        A = 23.5 * PAM ** 2.8

    if F1 <= 0.25:
        pass
    elif F1 <= 0.35:
        A = A*5.2*F1**1.19
    else:
        A = A*0.5/F1**1.05

    if A < 16:
        A = 16

    CT = 0
    UGS = 0
    
    for i in range(1,len(ib)):
        if i <= 1:
            GFB = 1.905/(T*A/(V**2))**1.667
            if ib[i]>0:
                if i <= 7:
                    D=DFT[i]
                    FV = FallVelocity(D, TEMP, AF)
                else:
                    FV=1.6
            else:
                continue
        else:
            GFA = GFB
            GFB = GFA/3.175
            if ib[i]>0:
                if i <= 7:
                    D=DFT[i]
                    FV = FallVelocity(D, TEMP, AF)
                else:
                    FV=1.6
            else:
                CList.append(0.0)
                UGSList.append(0.0)
                continue
                
        ZOM = FV * V/SI
        if ZOM < (1.5*ZV):
            ZOM = 1.5 * ZV
        F1 = 0.756 * ZOM - ZV
        F2 = ZOM - ZV
        F3 = 1.5 * ZOM - ZV
        F4 = 1.0 - F1
        F5 = 1.0 - F2
        F6 = 1.0 - F3
        YAF4 = YA**F4
        C = ib[i]*W
        DD = 2.0 * DFT[i]
        DDF4 = DD ** F4
        UD = CV * V * (DD/Y) ** ZV
        X = F4 * GFB / (YAF4 - DDF4)
        UGSI = X*DDF4
        UBL = UGSI/(43.2*UD*DD)
                
        if UBL > 100.0:
            UGSI = UGSI*100.0/UBL

        UGSI = C * UGSI
        
        if NLD != 2:
            GA = UGSI + C*GFB
            C = C*X
            YAF2 = YA**(F2-F1)
            YAF5 = YA**F5
            CF5 = C/F5
            YBF3 = YB**(F3-F2)
            YBF6 = YB**F6
            CF6 = C/F6
            CF4 = C/ F4
            GB = CF5 * YAF2 * (YB**F5 - YAF5)
            GC = CF6 * YAF2 * YBF3 * (Y**F6 - YBF6)
            UGSI = GA + GB + GC

        UGSI = UGSI/(43.2*W)
        UGS = UGS + UGSI
        CI = 16000 * UGSI/(Y*V)
        CT = CT + CI

        CList.append(CI)
        UGSList.append(0.0000625*CI*Y*V)
            
    return [CT, UGS, UGS*43.2*W, CList, UGSList]

class MainApp(QtGui.QMainWindow, design.Ui_MainWindow):
    def __init__(self):
        super(self.__class__, self).__init__()
        self.setupUi(self)

        self.CalcularPushButton.clicked.connect(self.RunProgram)
        self.SIntRadioButton.setChecked(True)
        self.SIngRadioButton.clicked.connect(self.ChangeToSIng)
        self.SIntRadioButton.clicked.connect(self.ChangeToSInt)
        self.CalcularDeclivPushButton.clicked.connect(self.CalcularDeclividade)
        self.actionSalvar.triggered.connect(self.Salvar)
        self.actionAbrir.triggered.connect(self.Abrir)
        self.actionSair.triggered.connect(self.Sair)
        self.ArrasteSuspensaoCheckBox.clicked.connect(self.ToogleArrasteSuspensao)
        self.ArrasteCheckBox.clicked.connect(self.ToogleArraste)
        self.actionSobre.triggered.connect(self.Sobre)

    def Salvar(self):
        print("Salvando...")
        Local = str(self.LocalLineEdit.text())
        settings = QtCore.QSettings(Local+'.ini', QtCore.QSettings.IniFormat) 

        for name, obj in inspect.getmembers(self):
            if isinstance(obj, QtGui.QComboBox):
                name   = obj.objectName()      # get combobox name
                index  = obj.currentIndex()    # get current index from combobox
                text   = obj.itemText(index)   # get the text for current index
                settings.setValue(name, text)   # save combobox selection to registry

            if isinstance(obj, QtGui.QLineEdit):
                name = obj.objectName()
                value = obj.text()
                settings.setValue(name, value)    # save ui values, so they can be restored next time

            if isinstance(obj, QtGui.QCheckBox):
                name = obj.objectName()
                state = obj.checkState()
                settings.setValue(name, state)

        print('Done!')
                            
    def Abrir(self):
        print('Abrindo...')
        try:
            ArquivodeDados = QtGui.QFileDialog.getOpenFileName()
            settings = QtCore.QSettings(ArquivodeDados, QtCore.QSettings.IniFormat)
            for name, obj in inspect.getmembers(self):
                if isinstance(obj, QtGui.QComboBox):
                    index  = obj.currentIndex()    # get current region from combobox
                    name   = obj.objectName()
                    value = unicode(settings.value(name).toString())  

                    if value == "":
                        continue

                    index = obj.findText(value)   # get the corresponding index for specified string in combobox

                    if index == -1:  # add to list if not found
                        obj.insertItems(0,[value])
                        index = obj.findText(value)
                        obj.setCurrentIndex(index)
                    else:
                        obj.setCurrentIndex(index)   # preselect a combobox value by index    

                if isinstance(obj, QtGui.QLineEdit):
                    name = obj.objectName()
                    value = unicode(settings.value(name).toString())  # get stored value from registry
                    obj.setText(value)  # restore lineEditFile

                if isinstance(obj, QtGui.QCheckBox):
                    name = obj.objectName()
                    value = int(settings.value(name).toString())  # get stored value from registry
                    obj.setCheckState(value)   # restore checkbox
        except:
            print('Error opening file.')
        
    def Sair(self):
        sys.stderr.write('\r')
        if QtGui.QMessageBox.question(None, '', "Are you sure you want to quit?",
                                QtGui.QMessageBox.Yes | QtGui.QMessageBox.No,
                                QtGui.QMessageBox.No) == QtGui.QMessageBox.Yes:
            QtGui.QApplication.quit()

    def Sobre(self):
        QtGui.QMessageBox.information(None, 'Copyright', "SEDIM 2.0  Copyright (C) 2017  ROBERTA CAMPEAO \n \nEste software foi desenvolvido na Dissertação apresentada ao Curso de Pós Graduação em Engenharia de Biossistemas da Universidade Federal Fluminense, com apoio financeiro da Fundação CAPES. \n \nBaseado em: HORA, Mônica de A. G. M.. Avaliação do Transporte de Sedimentos da Sub-bacia do Ribeirão do Rato, Região Noroeste do Estado do Paraná. Dissertação (Mestrado em Engenharia Civil) -  Universidade Federal do Rio de Janeiro, Rio de Janeiro, 1996. 287p. \n \nContato: Roberta Campeão - robertacampeao@gmail.com, Mônica da Hora - dahora@vm.uff.br \n  \nThis program comes with ABSOLUTELY NO WARRANTY. \nThis is free software, and you are welcome to redistribute it \nunder certain conditions; For details see readme file.")
        os.system('Manual.pdf')
        
    def ToogleArrasteSuspensao(self):
        if self.ArrasteSuspensaoCheckBox.isChecked():
            self.LaursenCheckBox.setCheckState(1)
            self.EngelundeHansenCheckBox.setCheckState(1)
            self.ColbyCheckBox.setCheckState(1)
            self.AckersWhiteD50CheckBox.setCheckState(1)
            self.AckersWhiteD35CheckBox.setCheckState(1)
            self.YangSandD50CheckBox.setCheckState(1)
            self.YangSandFTCheckBox.setCheckState(1)
            self.YangGravelD50CheckBox.setCheckState(1)
            self.YangGravelFTCheckBox.setCheckState(1)
            self.YangMixCheckBox.setCheckState(1)
##            self.ToffaletiCheckBox.setCheckState(1)
        if not self.ArrasteSuspensaoCheckBox.isChecked():
            self.LaursenCheckBox.setCheckState(0)
            self.EngelundeHansenCheckBox.setCheckState(0)
            self.ColbyCheckBox.setCheckState(0)
            self.AckersWhiteD50CheckBox.setCheckState(0)
            self.AckersWhiteD35CheckBox.setCheckState(0)
            self.YangSandD50CheckBox.setCheckState(0)
            self.YangSandFTCheckBox.setCheckState(0)
            self.YangGravelD50CheckBox.setCheckState(0)
            self.YangGravelFTCheckBox.setCheckState(0)
            self.YangMixCheckBox.setCheckState(0)
##            self.ToffaletiCheckBox.setCheckState(0)

    def ToogleArraste(self):
        if self.ArrasteCheckBox.isChecked():
            self.SchoklitschCheckBox.setCheckState(1)
            self.KallinskeCheckBox.setCheckState(1)
            self.MeyerPeterICheckBox.setCheckState(1)
            self.RottnerCheckBox.setCheckState(1)
        if not self.ArrasteCheckBox.isChecked():
            self.SchoklitschCheckBox.setCheckState(0)
            self.KallinskeCheckBox.setCheckState(0)
            self.MeyerPeterICheckBox.setCheckState(0)
            self.RottnerCheckBox.setCheckState(0)

    def ChangeToSIng(self):
        self.LarguraUnidLabel.setText('ft')
        self.ProfUnidLabel.setText('ft')
        self.VelocUnidLabel.setText('ft/s')
        self.DeclivUnidLabel.setText('ft/ft')
        self.TempUnidLabel.setText(u'ºC')
        self.D35UnidLabel.setText('mm')
        self.D50UnidLabel.setText('mm')
        self.D65UnidLabel.setText('mm')
        self.D90UnidLabel.setText('mm')
        self.VazaoUnidLabel.setText('ft3/s')
        self.RaioUnidLabel.setText('ft')
        self.AreaUnidLabel.setText('ft2')

    def ChangeToSInt(self):
        self.LarguraUnidLabel.setText('m')
        self.ProfUnidLabel.setText('m')
        self.VelocUnidLabel.setText('m/s')
        self.DeclivUnidLabel.setText('m/m')
        self.TempUnidLabel.setText(u'ºC')
        self.D35UnidLabel.setText('mm')
        self.D50UnidLabel.setText('mm')
        self.D65UnidLabel.setText('mm')
        self.D90UnidLabel.setText('mm')
        self.VazaoUnidLabel.setText('m3/s')
        self.RaioUnidLabel.setText('m')
        self.AreaUnidLabel.setText('m2')

    def CalcularDeclividade(self):
        if self.SIntRadioButton.isChecked():            
            q = float(self.VazaoLineEdit.text()) #Vazao linear m3/s
            n = float(self.RugosidadeLineEdit.text()) #coeficiente de rugosidade
            R = float(self.RaioLineEdit.text())  #Raio hidraulico m
            A = float(self.AreaLineEdit.text())  #Area da secao m2
            V = float(self.VelocLineEdit.text()) #m/s velocidade media
            S = (q * n /(A*R**(0.666)))**2 #declividade = gradiente de energia m/m
            self.DeclivLineEdit.setText(str(S))
            
        if self.SIngRadioButton.isChecked():
            q = float(self.VazaoLineEdit.text()) * 0.3048**3.0  #Vazao linear m3/s
            n = float(self.RugosidadeLineEdit.text()) #coeficiente de rugosidade
            R = float(self.RaioLineEdit.text()) * 0.3048 #Raio hidraulico m
            V = float(self.VelocLineEdit.text()) * 0.3048 #m/s velocidade media
            A = float(self.AreaLineEdit.text()) * 0.3048 ** 2.0  #Area da secao m2
            S = (q * n / (A.R**(0.666)))**2 #declividade = gradiente de energia m/m
            self.DeclivLineEdit.setText(str(S))

    def RunProgram(self):
        try:
            if self.SIntRadioButton.isChecked():
                g = 32.1725 #ft/s^2 aceleracao da gravidade
                Y = float(self.ProfLineEdit.text()) * 1000.0 / 304.8 #ft profundidade
                V = float(self.VelocLineEdit.text()) * 1000.0 / 304.8 #ft/s velocidade media
                DF35 = float(self.D35LineEdit.text()) / 304.8 #ft
                DF50 = float(self.D50LineEdit.text()) / 304.8 #ft
                DF65 = float(self.D65LineEdit.text()) / 304.8 #ft
                DF90 = float(self.D90LineEdit.text()) / 304.8 #ft
                W = float(self.LarguraLineEdit.text())* 1000.0 / 304.8 #ft Largura do rio

            if self.SIngRadioButton.isChecked():
                g = 32.1725 #ft/s^2 aceleracao da gravidade
                Y = float(self.ProfLineEdit.text()) #ft profundidade
                V = float(self.VelocLineEdit.text()) #ft/s velocidade media
                DF35 = float(self.D35LineEdit.text()) / 304.8 #ft
                DF50 = float(self.D50LineEdit.text()) / 304.8 #ft
                DF65 = float(self.D65LineEdit.text()) / 304.8 #ft
                DF90 = float(self.D90LineEdit.text()) / 304.8 #ft
                W = float(self.LarguraLineEdit.text()) #ft Largura do rio

            ib = [float(self.Faixa1LineEdit.text())/100, float(self.Faixa2LineEdit.text())/100,
                  float(self.Faixa3LineEdit.text())/100, float(self.Faixa4LineEdit.text())/100,
                  float(self.Faixa5LineEdit.text())/100, float(self.Faixa6LineEdit.text())/100,
                  float(self.Faixa7LineEdit.text())/100, float(self.Faixa8LineEdit.text())/100,
                  float(self.Faixa9LineEdit.text())/100, float(self.Faixa10LineEdit.text())/100,
                  float(self.Faixa11LineEdit.text())/100] #em porcentagem

            Local = str(self.LocalLineEdit.text())
            GMS = 165.36 #lb/ft^3 Peso especifico do sedimento
            SG = 2.65 #densidade relativa do sedimento adm
            TEMP = float(self.TempLineEdit.text()) #Celsius
            S = float(self.DeclivLineEdit.text())  #declividade = gradiente de energia ft/ft

        except:
            print("Verifique os dados de entrada.")
            QtGui.QMessageBox.warning(self, 'Erro!', "Verifique os dados de entrada.")
            return
        
        if sum(ib)>1:
            QtGui.QMessageBox.warning(self, 'Erro!', "Soma das faixas maior que 100%")
            
        U = math.sqrt(g*Y*S)
        
        COMP1 = 1.0334+0.03672*TEMP+0.0002058*TEMP**2
        XNU = 0.00002/COMP1

        COMP2 = 6 * XNU
        FV50 =  (math.sqrt(36.064*DF50**3+COMP2**2)-COMP2)/DF50
        
        CY = [[0.1000, 0.2000,  0.3000,  0.400,  0.8000,  0.0000,  0.0000],
          [0.6100, 0.4800,  0.3000,  0.300,  0.3000,  0.0000,  0.0000],
          [1.4530, 1.3290,  1.4000,  1.260,  1.0990,  0.0000,  0.0000],
          [0.0100, 5.0000, 10.0000, 15.600, 20.0000, 30.0000, 40.0000],
          [0.1057, 0.0845,  0.0469,  0.000, -0.0277, -0.0654, -0.1155],
          [0.0735, 0.0166,  0.0014,  0.000, -0.0164, -0.0610, -0.0763],
          [0.0118, 0.0202,  0.0135,  0.000,  0.0000,  0.0000,  0.0000]]
    
        CF = [0.64, 1.0, 1.0, 0.88, 0.2]

        AF = [[0.00001,	0.06,	0.10,	0.20,	0.4,	0.80,	1.50,	2.00,	3.00,	7.00,	8.00,	9.00,	10.0],
          [0.0010,	0.24,	0.60,   1.80,	4.6,	9.50,	16.1,	19.9,	25.3,	39.5,	41.5,	43.5,	45.0],
          [0.0001,	0.32,	0.76,	2.20,	5.3,	10.5,	16.9,	20.3,	25.6,	39.5,	41.5,	43.5,	45.0],
          [0.0001,	0.40,   0.92,	2.50,	5.8,	11.0,	17.5,	20.7,	25.9,	39.5,	41.5,	43.5,	45.0],
          [0.0001,	0.49,	1.10,	2.85,	6.3,	11.6,	17.9,	21.1,	26.2,	39.5,	41.5,	43.5,	45.0],
          [0.0001,	0.57,	1.26,	3.20,	6.7,	12.0,	18.1,	21.5,	26.5,	39.5,	41.5,	43.5,	45.0]]

        DIP = [0,0.37,0.71,0.99,1.21,1.34,1.41,1.38,1.27,1.11,0.94,0.78,0.65,0.55,0.49,0.45,0.42,0.4]
    
        AK0 = -0.068
        AK1 = -1.1328
        AK2 = 0.94
        AK3 = -1.206
        AK4 = 0.567
        AK5 = -0.0975
    
        DFT = []
        Fgran = [0.016, 0.062, 0.125, 0.25, 0.5, 1.0, 2.0, 4.0, 8.0, 16.0, 32.0, 64.0]

        Fgranlabel = ['0.0 - 0.062', '0.062 - 0.125', '0.125 - 0.25',
                      '0.25 - 0.5', '0.5 - 1.0', '1.0 - 2.0', '2.0 - 4.0', '4.0 - 8.0',
                      '8.0 - 16.0', '16.0 - 32.0', '32.0 - 64.0']
        
        for i in range(len(Fgran)-1):
            DFT.append(math.sqrt(Fgran[i]*Fgran[i+1])/304.8)

        Dm = 0
        for D, i in zip(DFT, ib):
            Dm = Dm + D*304.8*i

        wbk = xlsxwriter.Workbook(Local+'.xlsx')
        sheetSInt = wbk.add_worksheet("Sist Internacional")
        sheetSIng = wbk.add_worksheet("Sist Ingles")
        
        row = 0

        ErrorMessages = []
            
        if self.LaursenCheckBox.isChecked():
            try:
                #print "Laursen", Laursen(DFT, ib, V, SG, g, Y, XNU, U, DF50, W)
                ResultLaursen = Laursen(DFT, ib, V, SG, g, Y, XNU, U, DF50, W)
                row = WriteResultsFaixa('Laursen', ResultLaursen, ib, Fgranlabel[1:], row, sheetSInt, sheetSIng, 1)
            except Exception as e:
                ErrorMessages.append(e)

        if self.EngelundeHansenCheckBox.isChecked():
            try:
                #print "EngelundeHansen", EngelundeHansen(GMS, V, S, DF50, g, SG, W, Y)
                ResultEngelundeHansen = EngelundeHansen(GMS, V, S, DF50, g, SG, W, Y)
                row = WriteResultsFaixa('EngelundeHansen', ResultEngelundeHansen, ib, Fgranlabel[1:], row, sheetSInt, sheetSIng, 0)
            except Exception as e:
                ErrorMessages.append(e)

        if self.ColbyCheckBox.isChecked():
            try:
                #print "Colby", Colby(DF50, Y, V, CY, CF, W, TEMP)
                ResultColby = Colby(DF50, Y, V, CY, CF, W, TEMP)
                row = WriteResultsFaixa('Colby', ResultColby, ib, Fgranlabel[1:], row, sheetSInt, sheetSIng, 0)
            except Exception as e:
                ErrorMessages.append(e)

        if self.AckersWhiteD35CheckBox.isChecked():
            try:
                #print "AckersWhiteD35", AckersWhite(DF35, g, SG, XNU, U, V, Y, W)
                ResultAckersWhite = AckersWhite(DF35, g, SG, XNU, U, V, Y, W)
                row = WriteResultsFaixa('AckersWhiteD35', ResultAckersWhite, ib, Fgranlabel[1:], row, sheetSInt, sheetSIng, 0)
            except Exception as e:
                ErrorMessages.append(e)

        if self.AckersWhiteD50CheckBox.isChecked():
            try:
                #print "AckersWhiteD50", AckersWhite(DF50, g, SG, XNU, U, V, Y, W)
                ResultAckersWhite = AckersWhite(DF50, g, SG, XNU, U, V, Y, W)
                row = WriteResultsFaixa('AckersWhiteD50', ResultAckersWhite, ib, Fgranlabel[1:], row, sheetSInt, sheetSIng, 0)
            except Exception as e:
                ErrorMessages.append(e)

        if self.YangSandD50CheckBox.isChecked():
            try:
                #print "YangSandD50", YangD50('Sand', FV50, DF50, TEMP, AF, D, U, XNU, V, S, Y, W)
                ResultYangSandD50 = YangD50('Sand', FV50, DF50, TEMP, AF, D, U, XNU, V, S, Y, W)
                row = WriteResultsFaixa('YangSandD50', ResultYangSandD50, ib, Fgranlabel[1:], row, sheetSInt, sheetSIng, 0)
            except Exception as e:
                ErrorMessages.append(e)

        if self.YangSandFTCheckBox.isChecked():
            try:
                #print "YangSandFT", YangFT('Sand', ib, DFT, TEMP, AF, XNU, U, V, S, Y, W)
                ResultYangSandFT = YangFT('Sand', ib, DFT, TEMP, AF, XNU, U, V, S, Y, W)
                row = WriteResultsFaixa('YangSandFT', ResultYangSandFT, ib, Fgranlabel[1:], row, sheetSInt, sheetSIng, 1)
            except Exception as e:
                ErrorMessages.append(e)

        if self.YangGravelD50CheckBox.isChecked():
            try:
                #print "YangGravelD50", YangD50('Gravel', FV50, DF50, TEMP, AF, D, U, XNU, V, S, Y, W)
                ResultYangGravelD50 = YangD50('Gravel', FV50, DF50, TEMP, AF, D, U, XNU, V, S, Y, W)
                row = WriteResultsFaixa('YangGravelD50', ResultYangGravelD50, ib, Fgranlabel[1:], row, sheetSInt, sheetSIng, 0)
            except Exception as e:
                ErrorMessages.append(e)

        if self.YangGravelFTCheckBox.isChecked():
            try:
                #print "YangGravelFT", YangFT('Gravel', ib, DFT, TEMP, AF, XNU, U, V, S, Y, W)
                ResultYangGravelSF = YangFT('Gravel', ib, DFT, TEMP, AF, XNU, U, V, S, Y, W)
                row = WriteResultsFaixa('YangGravelFT', ResultYangGravelSF, ib, Fgranlabel[1:], row, sheetSInt, sheetSIng, 1)
            except Exception as e:
                ErrorMessages.append(e)

        if self.YangMixCheckBox.isChecked():
            try:
                #print "YangMixFT", YangFT('Mixture', ib, DFT, TEMP, AF, XNU, U, V, S, Y, W)
                ResultYangMixFT = YangFT('Mixture', ib, DFT, TEMP, AF, XNU, U, V, S, Y, W)
                row = WriteResultsFaixa('YangMixFT', ResultYangMixFT, ib, Fgranlabel[1:], row, sheetSInt, sheetSIng, 1)
            except Exception as e:
                ErrorMessages.append(e)

##        if self.ToffaletiCheckBox.isChecked():
##            NLD = 1.0
##            try:
##                print "Toffaleti", Toffaleti(TEMP, Y, S, V, DF65, g, XNU, DFT, DIP, ib, AF, NLD, W)
##                ResultToffaleti = Toffaleti(TEMP, Y, S, V, DF65, g, XNU, DFT, DIP, ib, AF, NLD, W)
##                row = WriteResultsFaixa('Toffaleti', ResultToffaleti, ib, Fgranlabel[1:], row, sheetSInt, sheetSIng, 1)
##            except Exception as e:
##                ErrorMessages.append(e)

        if self.SchoklitschCheckBox.isChecked():
            try:
                #print "Schoklitsch", Schoklitsch(DFT, ib, S, Y, V, W)
                ResultSchoklitsch = Schoklitsch(DFT, ib, S, Y, V, W)
                row = WriteResultsFaixa('Schoklitsch', ResultSchoklitsch, ib, Fgranlabel, row, sheetSInt, sheetSIng, 1)
            except Exception as e:
                ErrorMessages.append(e)

        if self.KallinskeCheckBox.isChecked():
            try:
                #print "Kalinske", Kalinske(DFT, ib, AK0, AK1, AK2, AK3, AK4, AK5, Y, V, S, W)
                ResultKalinske = Kalinske(DFT, ib, AK0, AK1, AK2, AK3, AK4, AK5, Y, V, S, W)
                row = WriteResultsFaixa('Kalinske', ResultKalinske, ib, Fgranlabel, row, sheetSInt, sheetSIng, 1)
            except Exception as e:
                ErrorMessages.append(e)

        if self.MeyerPeterICheckBox.isChecked():
            try:
                #print "MeyerPetereMuller", MeyerPetereMuller(DF90, Dm, Y, V, S, W)
                ResultMeyerPetereMuller = MeyerPetereMuller(DF90, Dm, Y, V, S, W)
                row = WriteResultsFaixa('MeyerPetereMuller', ResultMeyerPetereMuller, ib, Fgranlabel, row, sheetSInt, sheetSIng, 0)
            except Exception as e:
                ErrorMessages.append(e)

        if self.RottnerCheckBox.isChecked():
            try:
                #print "Rottner", Rottner(GMS, SG, g, Y, V, DF50, W)
                ResultRottner = Rottner(GMS, SG, g, Y, V, DF50, W)
                row = WriteResultsFaixa('Rottner', ResultRottner, ib, Fgranlabel, row, sheetSInt, sheetSIng, 0)
            except Exception as e:
                ErrorMessages.append(e)

        if row == 0:
            QtGui.QMessageBox.warning(self, 'Erro!', "Selecione um metodo.")
            wbk.close()
            return

        wbk.close()

        if ErrorMessages != []:
            errorhand = open('Erros.txt', 'w')
            print('Error messages:')
            for e in ErrorMessages:
                print(e)
                print >>errorhand, e
            QtGui.QMessageBox.information(self, 'Aviso:', "Checar mensagens de erro!")
        else:
            QtGui.QMessageBox.information(self, 'Sucesso!', "Resultados exportados!")

def main():
    app = QtGui.QApplication(sys.argv)  # A new instance of QApplication
    form = MainApp()                    # We set the form to be our MainApp
    form.show()                         # Show the form
    app.exec_()                         # and execute the app

if __name__ == '__main__':
    main()

import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from matplotlib import ticker

#Caso 1 Subsaturado
def F_caso1(Np, Bo, Boi=None):
    F = Np*Bo
    return F

def Eo_caso1(Bo, Boi, Rsi, Rs, Bg):
    Eo = Bo - Boi + ((Rsi - Rs)*Bg)
    return Eo

def Efw_caso1(Boi, Cf, Swi, Cw, delta_P):
    Efw = Boi*((Cw*Swi + Cf)/(1-Swi))*delta_P
    return Efw

#FCaso 3 Capa de gas y Rp corregido
def Eg_caso3(Bti, Bg, Bgi):
    Eg = Bti*((Bg/Bgi)-1)
    return Eg

def Eo_caso3(Bt, Bti, Rsi =None, Rs =None, Bg=None):
    Eo = Bt - Bti + ((Rsi - Rs) * Bg)
    return Eo
def F_caso3(Np, Bt, Rp, Rsi, Bg):
    F = Np*(Bt + ((Rp - Rsi)*Bg))
    return F

def Rp_corregido(N, Eo, Eg, m, Np, Bt, Rsi, Bg):
    Rp_co = ((N*Eo) + (N*m*Eg) + (Np*Rsi*Bg) - (Np*Bt))/(Np*Bg)
    return Rp_co


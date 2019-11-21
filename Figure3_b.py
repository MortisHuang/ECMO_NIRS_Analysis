# -*- coding: utf-8 -*-
"""
Created on Thu Nov 21 15:08:43 2019

@author: Mortis
"""

import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
plt.rcParams["font.family"] = "Times New Roman"
sns.set_style("whitegrid")

#%% VA
read = pd.read_excel(r".\VA\VA_ALL_BAR.xlsx")
VA_NC=read.loc[read["Type"]=="NC"]
VA_NC=VA_NC.replace('NC', "Non-cannulation")
VA_C=read.loc[read["Type"]=="C"]
VA_C=VA_C.replace("C", "Cannulation")
#%%
VA_C_average = pd.DataFrame([np.average(VA_C.loc[VA_C["Stage"]=="Stage 1"]["Average HbO2"]),
                             np.average(VA_C.loc[VA_C["Stage"]=="Stage 2"]["Average HbO2"]),
                             np.average(VA_C.loc[VA_C["Stage"]=="Stage 3"]["Average HbO2"]),
                             np.average(VA_C.loc[VA_C["Stage"]=="Stage 4"]["Average HbO2"]),
                             np.average(VA_C.loc[VA_C["Stage"]=="Stage 5"]["Average HbO2"])],columns=["Values"])
VA_C_average["Type"] = ["Cannulation","Cannulation","Cannulation","Cannulation","Cannulation"]
VA_C_average["Stage"] = ["1","2","3","4","5"]

VA_NC_average = pd.DataFrame([np.average(VA_NC.loc[VA_NC["Stage"]=="Stage 1"]["Average HbO2"]),
                             np.average(VA_NC.loc[VA_NC["Stage"]=="Stage 2"]["Average HbO2"]),
                             np.average(VA_NC.loc[VA_NC["Stage"]=="Stage 3"]["Average HbO2"]),
                             np.average(VA_NC.loc[VA_NC["Stage"]=="Stage 4"]["Average HbO2"]),
                             np.average(VA_NC.loc[VA_NC["Stage"]=="Stage 5"]["Average HbO2"])],columns=["Values"])
VA_NC_average["Type"] = ["Non-cannulation","Non-cannulation","Non-cannulation","Non-cannulation","Non-cannulation"]
VA_NC_average["Stage"] = ["1","2","3","4","5"]
#%%

fig,ax = plt.subplots(1, 2, figsize=(8, 6))
plt.rcParams["font.family"] = "Times New Roman"
g1=sns.factorplot(x="Type", y="Values", hue='Stage', data=VA_NC_average,palette=sns.color_palette("Paired")[1:], kind='bar',ci=None,ax=ax[0],legend=False)
ax[0].set_ylabel(r'$\Delta HbO_2$ ($\mu$mol/L) ')    
ax[0].set_xlabel('')
ax[0].legend().set_visible(False)
ax[0].set_ylim(0,5)

g2=sns.factorplot(x="Type", y="Values", hue='Stage', data=VA_C_average,palette=sns.color_palette("Paired")[1:],kind='bar',ci=None,ax=ax[1],legend=False)
ax[1].set_ylabel('')    
ax[1].set_xlabel('')
ax[1].legend().set_visible(False)
ax[1].set_ylim(0,5)

fig.legend(["Stage 1","Stage 2","Stage 3","Stage 4","Stage 5"],loc="lower center",ncol=5)
fig.savefig('VA_Group_Bar.png',dpi=360)
#plt.close('all')
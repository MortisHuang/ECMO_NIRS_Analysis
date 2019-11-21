# -*- coding: utf-8 -*-
"""
Created on Wed Nov 20 14:01:22 2019

@author: Mortis

請先更改Excel原始檔中的第66行，把column名標註上去
"""

import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
sns.set_style("ticks")
sns.set( font="Times New Roman")
plt.rcParams["font.family"] = "Times New Roman"

from openpyxl import load_workbook
file_name = "20180928-39184015.xlsx"
pa_number = file_name.split('-')[1].split('.')[0]
wb = load_workbook(filename=r"./VA/{}".format(file_name), 
                   read_only=True)

ws = wb['Export 1']
sampling_rate = 50    #要按照原始檔中紀錄的Sampling Rate 來修改
# Read the cell values into a list of lists
data_rows = []
for row in ws:
    data_cols = []
    for cell in row:
        data_cols.append(cell.value)
    data_rows.append(data_cols)

df = pd.DataFrame(data_rows[65:])
colname=list(df.iloc[0])
df = df.iloc[1:]
df.columns = colname

mav_para = 1800

CH1_HbO2_mav = df.CH1_HbO2.ewm(span=mav_para, adjust=True).mean()
CH2_HbO2_mav = df.CH2_HbO2.ewm(span=mav_para, adjust=True).mean()
CH3_HbO2_mav = df.CH3_HbO2.ewm(span=mav_para, adjust=True).mean()

CH1_HHb_mav = df.CH1_HHb.ewm(span=mav_para, adjust=True).mean()
CH2_HHb_mav = df.CH2_HHb.ewm(span=mav_para, adjust=True).mean()
CH3_HHb_mav = df.CH3_HHb.ewm(span=mav_para, adjust=True).mean()

HbO2_mav_data=pd.concat([CH1_HbO2_mav, CH2_HbO2_mav, CH3_HbO2_mav],axis=1)

HbO2_average_a = pd.DataFrame(HbO2_mav_data.iloc[:,[0,2,4]].mean(1))   #如果檔名是39184015以外的，要把[0,2,4]跟[1,3,5]互換
HbO2_average_a.columns=["HbO2_average_a"]

HbO2_average_b = pd.DataFrame(HbO2_mav_data.iloc[:,[1,3,5]].mean(1))   #如果檔名是39184015以外的，要把[0,2,4]跟[1,3,5]互換
HbO2_average_b.columns=["HbO2_average_b"]

HHb_mav_data=pd.concat([CH1_HHb_mav, CH2_HHb_mav, CH3_HHb_mav],axis=1)

HHb_average_a = pd.DataFrame(HHb_mav_data.iloc[:,[0,2,4]].mean(1))    #如果檔名是39184015以外的，要把[0,2,4]跟[1,3,5]互換
HHb_average_a.columns=["HHb_average_a"]

HHb_average_b = pd.DataFrame(HHb_mav_data.iloc[:,[1,3,5]].mean(1))    #如果檔名是39184015以外的，要把[0,2,4]跟[1,3,5]互換
HHb_average_b.columns=["HHb_average_b"]

TSI_a = df.iloc[:,[2]].ewm(span=mav_para, adjust=True).mean()        #如果檔名是39184015以外的，要把[2]跟[12]互換
TSI_b = df.iloc[:,[12]].ewm(span=mav_para, adjust=True).mean()

#%% Both side
sns.set_style("ticks")
#sns.set(font="Times New Roman")
plt.rcParams["font.family"] = "Times New Roman"
fig = plt.figure(figsize=(8,6))
ax1 = fig.add_subplot(111)
minor=ax1.plot(df.loc[:,["time(ms)"]].div(1000),HbO2_average_a,alpha=1,label="$\Delta HbO_2$", color='r', linewidth=3.0)
major=ax1.plot(df.loc[:,["time(ms)"]].div(1000),HbO2_average_b,alpha=1, color='r', linewidth=2.5, linestyle = '-.')
ax1.plot(df.loc[:,["time(ms)"]].div(1000),HHb_average_a,alpha=1 ,label="$\Delta HHb$"  , color='b', linewidth=3.0)
ax1.plot(df.loc[:,["time(ms)"]].div(1000),HHb_average_b,alpha=1 , color='b', linewidth=2.5, linestyle = '-.')
ax1.axvspan(0, 300, alpha=0.5, color=sns.color_palette("Paired")[0])
ax1.axvspan(301, 901, alpha=0.5, color=sns.color_palette("Paired")[1])
ax1.axvspan(901, 1501, alpha=0.5, color=sns.color_palette("Paired")[2])
ax1.axvspan(1501, 2101, alpha=0.5, color=sns.color_palette("Paired")[3])
ax1.axvspan(2101, 2701, alpha=0.5, color=sns.color_palette("Paired")[4])
ax1.axvspan(2701, 3300, alpha=0.5, color=sns.color_palette("Paired")[5])
handles, labels = ax1.get_legend_handles_labels()
ax1.legend(["Non-cannulation Side","Cannulation Side"],loc = 'best', ncol=2)
ax1.set_xlim((0, np.max(df.loc[:,["time(ms)"]].div(1000))[0]))
#ax1.set_ylim((-8.0,4.5))
ax1.set_xlabel("Time (s)")
ax1.set_ylabel(r'$\mu$mol/L')
#leg1 = ax1.legend([minor,major],['max','min'], loc='lower right')
ax2 = ax1.twinx()
ax2.plot(df.loc[:,["time(ms)"]].div(1000),TSI_a,alpha=0.8,label=r'$TSI(\%)$',color='g', linewidth=3.0)
ax2.plot(df.loc[:,["time(ms)"]].div(1000),TSI_b,alpha=0.8,color='g', linewidth=3.0, linestyle = '-.')
ax2.set_ylabel('%',color='g')
ax2.set_ylim((35,75))
ax2.tick_params(axis='y', labelcolor='g')
#fig.suptitle("Data of patient VA Both side",size = 15)
fig.legend(loc='upper center', ncol=3,bbox_to_anchor=(0.5, -0.08), bbox_transform=ax1.transAxes)
plt.savefig("{}_NIRS_signal plot.png".format(pa_number),dpi=360)

#%% Bar



#HbO2_average_a
stage1_HbO_nc = np.mean(HbO2_average_a[301*sampling_rate:900*sampling_rate])[0]
stage2_HbO_nc = np.mean(HbO2_average_a[901*sampling_rate:1500*sampling_rate])[0]
stage3_HbO_nc = np.mean(HbO2_average_a[1501*sampling_rate:2100*sampling_rate])[0]
stage4_HbO_nc = np.mean(HbO2_average_a[2101*sampling_rate:2700*sampling_rate])[0]
stage5_HbO_nc = np.mean(HbO2_average_a[2701*sampling_rate:3300*sampling_rate])[0]

#HbO2_average_b
stage1_HbO_c = np.mean(HbO2_average_b[301*sampling_rate:900*sampling_rate])[0]
stage2_HbO_c = np.mean(HbO2_average_b[901*sampling_rate:1500*sampling_rate])[0]
stage3_HbO_c = np.mean(HbO2_average_b[1501*sampling_rate:2100*sampling_rate])[0]
stage4_HbO_c = np.mean(HbO2_average_b[2101*sampling_rate:2700*sampling_rate])[0]
stage5_HbO_c = np.mean(HbO2_average_b[2701*sampling_rate:3300*sampling_rate])[0]

#%%Merge
df_HbO2_plot_data = pd.concat([HbO2_average_a,HbO2_average_b],axis=1)
df_HHb_plot_data = pd.concat([HHb_average_a,HHb_average_b],axis=1)
df_TSI_plot_data = pd.concat([TSI_a,TSI_b],axis=1)
df_TSI_plot_data.columns = ["TSI % a","TSI % b"]
df_stage_average = pd.DataFrame([stage1_HbO_nc,stage2_HbO_nc,stage3_HbO_nc,stage4_HbO_nc,stage5_HbO_nc,
                                 stage1_HbO_c,stage2_HbO_c,stage3_HbO_c,stage4_HbO_c,stage5_HbO_c])
    
df_stage_average.columns=["Average HbO2"]
    
df_stage_average["Stage"] = ["NC Stage 1","NC Stage 2","NC Stage 3","NC Stage 4","NC Stage 5",
                             "C Stage 1", "C Stage 2", "C Stage 3", "C Stage 4", "C Stage 5"]

#%% save files

writer = pd.ExcelWriter('{}_NIRS_HbO2_plot.xlsx'.format(pa_number))
df_HbO2_plot_data.to_excel(writer,'Sheet 1',float_format='%.4f') 
writer.save()

writer = pd.ExcelWriter('{}_NIRS_HHb_plot.xlsx'.format(pa_number))
df_HHb_plot_data.to_excel(writer,'Sheet 1',float_format='%.4f') 
writer.save()

writer = pd.ExcelWriter('{}_NIRS_TSI_plot.xlsx'.format(pa_number))
df_TSI_plot_data.to_excel(writer,'Sheet 1',float_format='%.4f') 
writer.save()

writer = pd.ExcelWriter('{}_NIRS_Bar_plot.xlsx'.format(pa_number))
df_stage_average.to_excel(writer,'Sheet 1',float_format='%.4f') 
writer.save()


#%%
y_max=np.ceil(np.amax(df_stage_average["Average HbO2"]))
y_min=np.amin(df_stage_average["Average HbO2"])
if y_min > 0:
    y_min =0
    
fig,ax = plt.subplots(2,1,figsize=(8,6.2))


ax1 = ax[0]
#ax1.bar(0, VA_M_bar_data.Values.iloc[0],0.5, align='edge', alpha=0.8, color=sns.color_palette("Paired")[0],label='Baseine')
ax1.bar(0.5, stage1_HbO_nc,1, align='edge', alpha=1, color=sns.color_palette("Paired")[1],label='Stage 1')
ax1.bar(1.5, stage2_HbO_nc,1, align='edge', alpha=1, color=sns.color_palette("Paired")[2],label='Stage 2')
ax1.bar(2.5, stage3_HbO_nc,1, align='edge', alpha=1, color=sns.color_palette("Paired")[3],label='Stage 3')
ax1.bar(3.5, stage4_HbO_nc,1, align='edge', alpha=1, color=sns.color_palette("Paired")[4],label='Stage 4')
ax1.bar(4.5, stage5_HbO_nc,1, align='edge', alpha=1, color=sns.color_palette("Paired")[5],label='Stage 5')
ax1.set_xlim((0,5.5))
ax1.set_ylim((y_min,y_max))
ax1.spines['right'].set_visible(False)
ax1.spines['top'].set_visible(False)
ax1.spines['bottom'].set_visible(False)
ax1.set_xticks([])
ax1.set_title("Non-cannulation Side")


ax2 = ax[1]
#ax2.bar(0, VA_m_bar_data.Values.iloc[0],0.5, align='edge', alpha=0.8, color=sns.color_palette("Paired")[0],label='Baseine')
ax2.bar(0.5, stage1_HbO_c,1, align='edge', alpha=1, color=sns.color_palette("Paired")[1],label='Stage 1')
ax2.bar(1.5, stage2_HbO_c,1, align='edge', alpha=1, color=sns.color_palette("Paired")[2],label='Stage 2')
ax2.bar(2.5, stage3_HbO_c,1, align='edge', alpha=1, color=sns.color_palette("Paired")[3],label='Stage 3')
ax2.bar(3.5, stage4_HbO_c,1, align='edge', alpha=1, color=sns.color_palette("Paired")[4],label='Stage 4')
ax2.bar(4.5, stage5_HbO_c,1, align='edge', alpha=1, color=sns.color_palette("Paired")[5],label='Stage 5')
ax2.set_xlim((0,5.5))
ax2.set_ylim((y_min,y_max))
ax2.spines['right'].set_visible(False)
ax2.spines['top'].set_visible(False) 
ax2.spines['bottom'].set_visible(False)
ax2.set_xticks([])
ax2.set_title("Cannulation Side")

fig.legend(["Stage 1","Stage 2","Stage 3","Stage 4","Stage 5"],loc="lower center",ncol=5)
fig.text(0.04, 0.5, r'$\Delta HbO_2$ ($\mu$mol/L)', va='center', rotation='vertical')
fig.savefig("{}_NIRS_bar plot.png".format(pa_number),dpi=360)


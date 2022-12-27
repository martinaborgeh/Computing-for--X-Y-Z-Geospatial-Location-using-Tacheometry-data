import math as m
import csv
def Tacheo_Traver(Table_data, initial_Bearing, initial_Northings, initial_Eastings, backstation_initial_northings,
                  backstation_initial_eastings):
    bearing=None
    if ',' in initial_Bearing:
        bearing=initial_Bearing.split(',')
        bearing[0]=float(bearing[0])
        bearing[1]=float(bearing[1])/float(60)
        bearing[2]=float(bearing[2])/float(3600)
        bearing=sum(bearing)
    else:
        bearing=float(initial_Bearing)
    inst_stn = Table_data[0][0]
    Bearing = [bearing]
    HCR_ANGLE = ['']
    bear = bearing
    Average_HD =[]
    if Table_data[0][-6]!='' and  Table_data[0][-5]!='':
        Average_HD.append(round((float(Table_data[0][-6]) + (float(Table_data[0][-5]))) / 2, 4))
    elif Table_data[0][-6]!='' and Table_data[0][-5]=='':
        Average_HD.append(float(Table_data[0][-6]))
    elif Table_data[0][-6]=='' and Table_data[0][-5]!='':
        Average_HD.append(float(Table_data[0][-5]))
    Latitude = ['']
    Northings = [backstation_initial_northings]
    Departure = ['']
    Eastings = [float(backstation_initial_eastings)]
    point_name = [Table_data[0][1]]
    stn = [inst_stn]
    for i, item in enumerate(Table_data):
        if item[-2] == '' and item[3] == '':
            stn.append(inst_stn)
            pn = item[2]
            point_name.append(pn)
            LL = None
            RR = None
            if item[5]!='' and item[4]!='':
                if ',' in item[5] and ',' in Table_data[i - 1][5] and ',' in item[4] and ',' in Table_data[i - 1][4]:
                    back_ll=item[5].split(',')
                    back_ll[0]=float(back_ll[0])
                    back_ll[1]=float(back_ll[1])/float(60)
                    back_ll[2]=float(back_ll[2])/float(3600)
                    back_ll=sum(back_ll)
                    back_rr=item[4].split(',')
                    back_rr[0] = float(back_rr[0])
                    back_rr[1] = float(back_rr[1]) / float(60)
                    back_rr[2] = float(back_rr[2]) / float(3600)
                    back_rr = sum(back_rr)
                    forward_ll=Table_data[i - 1][5].split(',')
                    forward_ll[0] = float(forward_ll[0])
                    forward_ll[1] = float(forward_ll[1]) / float(60)
                    forward_ll[2] = float(forward_ll[2]) / float(3600)
                    forward_ll = sum(forward_ll)
                    forward_rr=Table_data[i - 1][4].split(',')
                    forward_rr[0] = float(forward_rr[0])
                    forward_rr[1] = float(forward_rr[1]) / float(60)
                    forward_rr[2] = float(forward_rr[2]) / float(3600)
                    forward_rr = sum(forward_rr)
                    LL=back_ll-forward_ll
                    RR=back_rr-forward_rr
                else:
                    LL=float(item[5]) - float(Table_data[i - 1][5])
                    RR=float(item[4]) - float(Table_data[i - 1][4])
                if LL < float(0) and RR > float(0):
                    HCR_ANGLE.append(round(((LL + float(360)) + RR) / 2, 4))
                elif LL > float(0) and RR < float(0):
                    HCR_ANGLE.append(round((LL + (RR + float(360))) / 2, 4))
                elif LL == float(0) and RR <= float(180):
                    HCR_ANGLE.append(round(RR, 4))
                elif LL == float(0) and RR > float(180):
                    HCR_ANGLE.append(round(RR, 4))
                elif RR == float(0) and LL <= float(180):
                    HCR_ANGLE.append(round(LL, 4))
                elif RR == float(0) and LL >= float(180):
                    HCR_ANGLE.append(round(LL, 4))
                elif LL > float(0) and RR > float(0):
                    HCR_ANGLE.append(round((LL + RR) / 2, 4))
                elif LL < float(0) and RR < float(0):
                    HCR_ANGLE.append(round(((LL + float(360)) + (RR + float(360))) / 2, 4))
            elif item[5]!='' and item[4]=='':
                if ',' in item[5] and ',' in Table_data[i - 1][5]:
                    back_ll = item[5].split(',')
                    back_ll[0] = float(back_ll[0])
                    back_ll[1] = float(back_ll[1]) / float(60)
                    back_ll[2] = float(back_ll[2]) / float(3600)
                    back_ll = sum(back_ll)
                    forward_ll = Table_data[i - 1][5].split(',')
                    forward_ll[0] = float(forward_ll[0])
                    forward_ll[1] = float(forward_ll[1]) / float(60)
                    forward_ll[2] = float(forward_ll[2]) / float(3600)
                    forward_ll = sum(forward_ll)
                    LL=back_ll-forward_ll
                else:
                    LL=float(item[5])-float(Table_data[i - 1][5])

                if LL < float(0):
                    HCR_ANGLE.append(LL + float(360))
                elif LL >= float(0):
                    HCR_ANGLE.append(LL)
            elif item[5]=='' and item[4]!='':
                if ',' in item[4] and Table_data[i - 1][4]:
                    back_rr = item[4].split(',')
                    back_rr[0] = float(back_rr[0])
                    back_rr[1] = float(back_rr[1]) / float(60)
                    back_rr[2] = float(back_rr[2]) / float(3600)
                    back_rr = sum(back_rr)
                    forward_rr = Table_data[i - 1][4].split(',')
                    forward_rr[0] = float(forward_rr[0])
                    forward_rr[1] = float(forward_rr[1]) / float(60)
                    forward_rr[2] = float(forward_rr[2]) / float(3600)
                    forward_rr = sum(forward_rr)
                    RR=back_rr-forward_rr
                else:
                    RR=float(item[4])-float(Table_data[i - 1][4])
                if RR < float(0):
                    HCR_ANGLE.append(RR + float(360))
                elif RR >= float(0):
                    HCR_ANGLE.append(RR)
            if item[-6] != '' and item[-5] != '':
                Average_HD.append(round((float(item[-6]) + (float(item[-5]))) / 2, 4))
            elif item[-6] != '' and item[-5] == '':
                Average_HD.append(float(item[-6]))
            elif item[-6] == '' and item[-5] != '':
                Average_HD.append(float(item[-5]))
            bearing = bear + HCR_ANGLE[-1]
            if float(0) < bearing <= float(360):
                Bearing.append(bearing)
                bear = bearing
                Latitude.append(round(Average_HD[-1] * m.cos(m.radians(bear)),4))
                Departure.append(round(Average_HD[-1] * m.sin(m.radians(bear)),4))
                Northings.append(round(Latitude[-1] + float(initial_Northings),4))
                Eastings.append(round(Departure[-1] + float(initial_Eastings),4))
            elif bearing < float(0):
                Bearing.append(bearing + float(360))
                bear = bearing + float(360)
                Latitude.append(round(Average_HD[-1] * m.cos(m.radians(bear)),4))
                Departure.append(round(Average_HD[-1] * m.sin(m.radians(bear)),4))
                Northings.append(round(Latitude[-1] + float(initial_Northings),4))
                Eastings.append(round(Departure[-1] + float(initial_Eastings),4))
            elif float(360) < bearing <= float(720):
                bear = bearing - float(360)
                Bearing.append(bearing - float(360))
                Latitude.append(round(Average_HD[-1] * m.cos(m.radians(bear)),4))
                Departure.append(round(Average_HD[-1] * m.sin(m.radians(bear)),4))
                Northings.append(round(Latitude[-1] + float(initial_Northings),4))
                Eastings.append(round(Departure[-1] + float(initial_Eastings),4))
            elif float(720) < bearing <= float(1440):
                Bearing.append(bearing - float(720))
                bear = bearing - float(720)
                Latitude.append(round(Average_HD[-1] * m.cos(m.radians(bear)),4))
                Departure.append(round(Average_HD[-1] * m.sin(m.radians(bear)),))
                Northings.append(round(Latitude[-1] + float(initial_Northings),4))
                Eastings.append(round(Departure[-1] + float(initial_Eastings),))
        elif item[-2] != '' and item[3] != '':
            initial_Northings = Northings[-1]
            initial_Eastings = Eastings[-1]

            inst_stn=item[0]
            if bear <= float(180):
                bear = bear + float(180)
            else:
                bear = bear - float(180)
        elif item[3]!='' and item[-2]=='':
            continue
    return list(zip(stn, HCR_ANGLE, Average_HD, Bearing, Latitude, Departure, Northings, Eastings, point_name))


def Tacheo_Heightning(data, instrument_elevation, backstation_elevation):
    elevation = [float(backstation_elevation)]
    average_vertical_dist = []
    if data[0][-4]!='' and  data[0][-3]!='':
        average_vertical_dist .append(round((float(data[0][-4]) + (float(data[0][-3]))) / 2, 4))
    elif data[0][-4]!='' and data[0][-3]=='':
        average_vertical_dist .append(float(data[0][-4]))
    elif data[0][-4]=='' and data[0][-3]!='':
        average_vertical_dist .append(float(data[0][-3]))
    average_prism_height = float(data[0][-1])
    instrument_height = float(data[0][3])
    verical_angle = ['']
    for item in data:
        if item[-2] == '' and item[3] == '':
            if item[-4] != '' and item[-3] != '':
                average_vertical_dist.append(round((float(item[-4]) + (float(item[-3]))) / 2, 4))
            elif item[-4] != '' and item[-3] == '':
                average_vertical_dist.append(float(item[-4]))
            elif item[-4] == '' and item[-3] != '':
                average_vertical_dist.append(float(item[-3]))
            average_prism_height = float(item[-1])
            elevation.append(float(instrument_elevation) + instrument_height + average_vertical_dist[-1] - \
                             average_prism_height)
            if item[6]!='' and item[7]!='':
                VCR_LL=None
                VCR_RR=None
                if ',' in item[6] and ',' in item[7]:
                    VCR_LL=item[6].split(',')
                    VCR_LL[0]=float(VCR_LL[0])
                    VCR_LL[1]=float(VCR_LL[1])/float(60)
                    VCR_LL[2]=float(VCR_LL[2])/float(3600)
                    VCR_LL=sum(VCR_LL)
                    VCR_RR = item[7].split(',')
                    VCR_RR[0] = float(VCR_RR[0])
                    VCR_RR[1] = float(VCR_RR[1]) / float(60)
                    VCR_RR[2] = float(VCR_RR[2]) / float(3600)
                    VCR_RR = sum(VCR_RR)
                else:
                    VCR_LL = float(item[6])
                    VCR_RR = float(item[7])
                if float(0) < VCR_LL < float(270) and float(0) < VCR_RR < float(270):
                    verical_angle.append(round(((float(90) - VCR_LL) + (float(90) - VCR_RR)) / 2, 4))
                elif VCR_LL >= float(270) and VCR_RR >= float(270):
                    verical_angle.append(round(((VCR_LL - float(270)) + (VCR_RR - float(270))) / 2, 4))
                elif float(0) < VCR_LL < float(270) and VCR_RR > float(270):
                    verical_angle.append(round(((float(90) - VCR_LL) + (VCR_RR - float(270))) / 2, 4))
                elif float(0) < VCR_RR < float(270) and VCR_LL > float(270):
                    verical_angle.append(round(((float(90) - VCR_RR) + (VCR_LL - float(270))) / 2, 4))
            elif item[6]!='' and item[7]=='':
                if ',' in item[6]:
                    VCR_LL = item[6].split(',')
                    VCR_LL[0] = float(VCR_LL[0])
                    VCR_LL[1] = float(VCR_LL[1]) / float(60)
                    VCR_LL[2] = float(VCR_LL[2]) / float(3600)
                    VCR_LL = sum(VCR_LL)
                else:
                    VCR_LL=float(item[6])
                if float(0)<VCR_LL < float(270):
                    verical_angle.append(float(90)-VCR_LL)
                elif VCR_LL>=float(270):
                    verical_angle.append(VCR_LL-float(270))
            elif item[6]=='' and item[7]!='':
                if ',' in item[7]:
                    VCR_RR = item[7].split(',')
                    VCR_RR[0] = float(VCR_RR[0])
                    VCR_RR[1] = float(VCR_RR[1]) / float(60)
                    VCR_RR[2] = float(VCR_RR[2]) / float(3600)
                    VCR_RR = sum(VCR_RR)
                else:
                    VCR_RR=float(item[7])
                if float(0)<VCR_RR < float(270):
                    verical_angle.append(float(90)-VCR_RR)
                elif VCR_RR>=float(270):
                    verical_angle.append(VCR_RR-float(270))
        elif item[-2] != '' and item[3] != '':
            if item[-4] != '' and item[-3] != '':
                average_vertical_dist.append(round((float(item[-4]) + (float(item[-3]))) / 2, 4))
            elif item[-4] != '' and item[-3] == '':
                average_vertical_dist.append(float(item[-4]))
            elif item[-4] == '' and item[-3] != '':
                average_vertical_dist.append(float(item[-3]))
            instrument_height = float(item[3])
            instrument_elevation = elevation[-1]
        else:
            continue
    return list(zip(average_vertical_dist, elevation,verical_angle))


def Convert_decimal_Bearing_to_degree(angle):
    Deg = []
    Min = []
    Sec = []
    for item in angle:
        item1 = float(item[5])
        Deg.append(int(item1))
        Min.append(int((item1 - int(item1)) * 60))
        Sec.append(int((((item1 - int(item1)) * 60) - int(((item1 - int(item1)) * 60))) * 60))
    return list(zip(Deg, Min, Sec))
def Convert_decimal_HCR_ANGLE_to_degree(angle):
    Deg = []
    Min = []
    Sec = []
    for item in angle:
        if item[1]!='':
            item1 = float(item[1])
            Deg.append(int(item1))
            Min.append(int((item1 - int(item1)) * 60))
            Sec.append(int((((item1 - int(item1)) * 60) - int(((item1 - int(item1)) * 60))) * 60))
        else:
            Deg.append('')
            Min.append('')
            Sec.append('')
    return list(zip(Deg, Min, Sec))
def Convert_decimal_VCR_ANGLE_to_degree(angle):
    Deg = []
    Min = []
    Sec = []
    for item in angle:
        if item[2]!='' :
            if float(item[2])>float(0):
                item1 = float(item[2])
                Deg.append(int(item1))
                Min.append(int((item1 - int(item1)) * 60))
                Sec.append(int((((item1 - int(item1)) * 60) - int(((item1 - int(item1)) * 60))) * 60))
            elif float(item[2])<float(0):
                item1 = abs(float(item[2]))
                Deg.append(int(-item1))
                Min.append(int((item1 - int(item1)) * 60))
                Sec.append(int((((item1 - int(item1)) * 60) - int(((item1 - int(item1)) * 60))) * 60))
            elif float(item[2])==float(0):
                Deg.append(0)
                Min.append(0)
                Sec.append(0)
        else:
            Deg.append('')
            Min.append('')
            Sec.append('')
            continue
    return list(zip(Deg, Min, Sec))
def Import_data_from_excel(file):
    rows=[]
    with open(file,'r') as filepath:
        data=csv.reader(filepath)
        header=next(data)
        for i in data:
            rows.append(i)
    return rows














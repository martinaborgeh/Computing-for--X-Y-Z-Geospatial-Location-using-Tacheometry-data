import csv
def Levelling_Computation(Data,initial_elev):
    # print(Data)
    # print(initial_elev)
    # try:
    Elevation = [float(initial_elev)]
    HPC = float(initial_elev) + float(Data[0][0])
    for elev in Data[1:]:
        if elev[0] != '' and elev[1] != '':
            Elevation.append(round(HPC - float(elev[1]), 3))
            HPC = Elevation[-1] + float(elev[0])
        elif elev[0] != '' and elev[2] != '':
            Elevation.append(round(HPC - float(elev[2]), 3))
            HPC = Elevation[-1] + float(elev[0])
        elif elev[0] == '' and elev[1] != '':
            Elevation.append(round(HPC - float(elev[1]), 3))
        elif elev[0] == '' and elev[2] != '':
            Elevation.append(round(HPC - float(elev[2]), 3))
    # except Exception as e:
    #     print(e)
    return Elevation








# def cal():

#
#     for elev in L[1:]:
#         if elev[0]!='':
#             Elevation.append(round(HPC-float(elev[1]),3))
#             HPC=Elevation[-1]+float(elev[0])
#         else:
#             Elevation.append(round(HPC-float(elev[1]),3))
#     return Elevation
# print(cal())




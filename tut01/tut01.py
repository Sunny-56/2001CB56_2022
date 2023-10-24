import pandas as pd
import csv

def func():

    l1 = []
    l2 = []
    l3 = []

    # opening file in read mode
    uu_sum = vv_sum = ww_sum = 0.0
    numb = 29746.0
    with open('octant_input.csv', 'r') as file:
        reader = csv.reader(file)

        for row in reader:
            if (row[1] != 'U'):
                uu_sum += float(row[1])
            if (row[2] != 'V'):
                vv_sum += float(row[2])
            if (row[3] != 'W'):
                ww_sum += float(row[3])

    # u Average , v Average, w Average is calculated.

        u_Average = uu_sum/numb
        v_Average = vv_sum/numb
        w_Average = ww_sum/numb
        lst4 = [u_Average]
        lst5 = [v_Average]
        lst6 = [w_Average]

    ln1 = ln2 = ln3 = ln4 = 0
    t = []
    u = []
    v = []
    w = []
    with open('octant_input.csv', 'r') as file:
        reader = csv.reader(file)

        for row in reader:
            if (row[1] != 'U'):
                l1.insert(ln1, float(row[1])-u_Average)
                u.insert(ln1, row[1])
                ln1 += 1
            if (row[2] != 'V'):
                l2.insert(ln2, float(row[2])-v_Average)
                v.insert(ln2, row[2])
                ln2 += 1
            if (row[3] != 'W'):
                l3.insert(ln3, float(row[3])-w_Average)
                w.insert(ln3, row[3])
                ln3 += 1
            if (row[0] != "Time"):
                t.insert(ln4, row[0])
                ln4 += 1

    for i in range(len(l1)-1):
        lst4.append(None)
        lst5.append(None)
        lst6.append(None)

    # creating a dataframe using pandas
    data_frame = pd.DataFrame()
    # storing the lists
    data_frame["Time"] = t
    data_frame["U"] = u
    data_frame["V"] = v
    data_frame["W"] = w
    data_frame["uavg"] = lst4
    data_frame["vavg"] = lst5
    data_frame["wavg"] = lst6
    data_frame["u-uavg"] = l1
    data_frame["v-vavg"] = l2
    data_frame["w-wavg"] = l3
    OCT_LST = []
    len5 = 0
    print(data_frame)
    # counting overall count
    i = 0
    o_cp1 = o_cp2 = o_cp3 = o_cp4 = o_cn1 = o_cn2 = o_cn3 = o_cn4 = 0
    while (i < len(l1)):
        if (l1[i] > 0 and l2[i] > 0):
            if (l3[i] > 0):
                o_cp1 += 1
                OCT_LST.insert(len5, 1)
                len5 += 1
            else:
                o_cn1 += 1
                OCT_LST.insert(len5, -1)
                len5 += 1
        elif (l1[i] < 0 and l2[i] > 0):
            if (l3[i] > 0):
                o_cp2 += 1
                OCT_LST.insert(len5, 2)
                len5 += 1
            else:
                o_cn2 += 1
                OCT_LST.insert(len5, -2)
                len5 += 1
        elif (l1[i] < 0 and l2[i] < 0):
            if (l3[i] > 0):
                o_cp3 += 1
                OCT_LST.insert(len5, 3)
                len5 += 1
            else:
                o_cn3 += 1
                OCT_LST.insert(len5, -3)
                len5 += 1
        elif (l1[i] > 0 and l2[i] < 0):
            if (l3[i] > 0):
                o_cp4 += 1
                OCT_LST.insert(len5, 4)
                len5 += 1
            else:
                o_cn4 += 1
                OCT_LST.insert(len5, -4)
                len5 += 1
        i = i+1


    # we define the func
    def octant_identification(mod):
        j = k = 0
        interval = mod

        while (j < rnd):

            cn1 = cn2 = cn3 = cn4 = cn_1 = cn_2 = cn_3 = cn_4 = 0
            while (k < mod and k < len(l1)):
                if (l1[k] > 0 and l2[k] > 0):
                    if (l3[k] > 0):
                        cn1 += 1
                    else:
                        cn_1 += 1
                elif (l1[k] < 0 and l2[k] > 0):
                    if (l3[k] > 0):
                        cn2 += 1
                    else:
                        cn_2 += 1
                elif (l1[k] < 0 and l2[k] < 0):
                    if (l3[k] > 0):
                        cn3 += 1
                    else:
                        cn_3 += 1
                elif (l1[k] > 0 and l2[k] < 0):
                    if (l3[k] > 0):
                        cn4 += 1
                    else:
                        cn_4 += 1
                k = k+1

            mod = mod+interval
            # appending the values of octant cn
            val1.append(cn1)
            vaneg1.append(cn_1)
            val2.append(cn2)
            vaneg2.append(cn_2)
            val3.append(cn3)
            vaneg3.append(cn_3)
            val4.append(cn4)
            vaneg4.append(cn_4)
            j = j+1


    data_frame["octant"] = OCT_LST
    mod_list = []
    mod_list.insert(0, None)
    mod_list.insert(1, "User input")
    for i in range(len(l1)-2):
        mod_list.insert(i+2, None)
    data_frame[""] = mod_list
    # taking input from user
    print("Enter the mod value")
    mod = int(input())
    L_TMp = []
    L_TMp.insert(0, "Overall Count")

    stng1 = "Mod "
    stng1 += str(mod)
    L_TMp.insert(1, stng1)
    num = 30000
    rnd = int(num/mod)
    MOD_VALUE = num % mod
    len6 = 2
    # creating lists for storing octant count
    val1 = []
    vaneg1 = []
    val2 = []
    vaneg2 = []
    val3 = []
    vaneg3 = []
    val4 = []
    vaneg4 = []
    # storing overall count of octants
    val1.insert(0, o_cp1)
    vaneg1.insert(0, o_cn1)
    val2.insert(0, o_cp2)
    vaneg2.insert(0, o_cn2)
    val3.insert(0, o_cp3)
    vaneg3.insert(0, o_cn3)
    val4.insert(0, o_cp4)
    vaneg4.insert(0, o_cn4)


    val1.insert(1, None)
    vaneg1.insert(1, None)
    val2.insert(1, None)
    vaneg2.insert(1, None)
    val3.insert(1, None)
    vaneg3.insert(1, None)
    val4.insert(1, None)
    vaneg4.insert(1, None)
    octant_identification(mod)


    while len(val1) < len(l1):
        val1.append(None)
        vaneg1.append(None)
        val2.append(None)
        vaneg2.append(None)
        val3.append(None)
        vaneg3.append(None)
        val4.append(None)
        vaneg4.append(None)


    for i in range(rnd):
        stng1 = str(mod*i)
        stng1 += "-"
        stng1 += str(mod*(i+1)-1)
        L_TMp.insert(len6, stng1)
        len6 += 1
    if (MOD_VALUE):
        stng1 = str(num-MOD_VALUE)
        stng1 += "-"
        stng1 += str(num)
        L_TMp.insert(len6, stng1)
        len6 += 1


    while len6 < len(u):
        L_TMp.insert(len6, None)
        len6 += 1
    data_frame["octant ID"] = L_TMp
    # inserting all columns of their octant value
    data_frame["1"] = val1
    data_frame["-1"] = vaneg1
    data_frame["2"] = val2
    data_frame["-2"] = vaneg2
    data_frame["3"] = val3
    data_frame["-3"] = vaneg3
    data_frame["4"] = val4
    data_frame["-4"] = vaneg4
    # printing in csv file
    data_frame.to_csv("octant_output.csv", index=False)

func()

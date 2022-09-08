def cont(x):
    import pandas as pd
    a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18,a19,a20,a21,a22,a23,a24,a25,a26,a27,a28,a29,a30 = [[] for i in range(30)]
    e1,e2,e3,e4,e5,e6,e7,e8,e9,e10,e11,e12,e13,e14,e15,e16,e17,e18,e19,e20,e21,e22,e23,e24,e25,e26,e27,e28,e29,e30 = [[] for i in range(30)]
    names = [a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18,a19,a20,a21,a22,a23,a24,a25,e26,a27,a28,a29,a30]
    es = [e1,e2,e3,e4,e5,e6,e7,e8,e9,e10,e11,e12,e13,e14,e15,e16,e17,e18,e19,e20,e21,e22,e23,e24,e25,e26,e27,e28,e29,e30]
    for i in range(x):
        names[i],es[i] = ref(i)
    while True:
        if len(names[-1]) == 0:
                names.pop()
                es.pop()
        if len(names[-1]) !=0:
                break
    with pd.ExcelWriter("C:/Users/Joshua/Documents/Cu/Project/Data/Readings - written.xlsx") as writer:
        for i in range(len(names)):
            names[i].to_excel(writer, sheet_name=es[i])
    print ("\nVery Successful!")

def ref(i):
        def exc(i):
                import pandas as pd
                acc = ["R600a 30g", "R600a 30g PCM", "R600a 30g PCM + NANO", "R600a 40g", "R600a 40g PCM", "R600a 40g PCM + NANO", "R600a 50g", "R600a 50g PCM", "R600a 50g PCM + NANO", "R600a 60g", "R600a 60g PCM", "R600a 60g PCM + NANO", "R600a 70g", "R600a 70g PCM", "R600a 70g PCM + NANO", "LPG 30g", "LPG 30g PCM", "LPG 30g PCM + NANO", "LPG 40g", "LPG 40g PCM", "LPG 40g PCM + NANO", "LPG 50g", "LPG 50g PCM", "LPG 50g PCM + NANO", "LPG 60g", "LPG 60g PCM", "LPG 60g PCM + NANO", "LPG 70g", "LPG 70g PCM", "LPG 70g PCM + NANO"]
                e = acc[i]
                print ("\n" + e)
                df = pd.read_excel(r'C:\Users\Joshua\Documents\Cu\Project\Data\Readings - Vals.xlsx', sheet_name= e)
                x = list(df[' Pd (psi)'])
                y = list(df['Ps (psi)'])
                return x,y,e
        x,y,e = exc(i)
        def inter(d,s,e):
                if len(d)!=len(s):
                        return "Pd & Ps asymmetrical"
                def val(table, pressure):
                        check = sorted(table)
                        temp = 0
                        if pressure in check:
                                return round(table[pressure], 1)
                        for i in range(len(check)):
                                if pressure < check[i]:
                                        diff = (table[check[i]] - table[check[i-1]])/ (check[i]-check[i-1])
                                        sd = table[check[i-1]] + ((pressure - check[i-1])*diff)
                                        return round(sd,1)
                                        
                def mix2(table1, table2, pressure, r1_per=0, r2_per=0):
                        x = val(table1, pressure)
                        y = val(table2, pressure)
                        return round((r1_per*x) + (r2_per*y),1)
                    
                def manss(a,b,c,d,f,g,e):
                        import pandas as pd
                        mans, mans2 = [],[]
                        for i in range(len(c)):
                                mans.append(mix2(a,b,c[i],f,g))
                                mans2.append(mix2(a,b,d[i],f,g))
                        df1 = pd.DataFrame({'Td':mans,'Ts':mans2})
                        return df1,e
                    
                R600a_psi = {-14: -70, -13.8: -66, -13.5: -62, -13.2: -58, -12.8: -54, -12.3: -50, -11.6: -46, -11: -42, -10.1: -38, -9.1: -34, -8: -30, -7.3: -28, -6.6: -26, -5.8: -24, -5.1: -22, -4.2: -20, -3.3: -18, -2.3: -16, -1.2: -14, -0.1: -12, 1: -10, 2.3: -8, 3.6: -6, 5: -4, 6.5: -2, 8.1: 0, 9.7: 2, 11: 4, 13: 6, 15: 8, 17: 10, 19: 12, 22: 14, 24: 16, 27: 18, 29: 20, 32: 22, 35: 24, 38: 26, 41: 28, 44: 30, 47: 32, 51: 34, 55: 36, 58: 38, 62: 40, 66: 42, 71: 44, 75: 46, 80: 48, 85: 50, 90: 52, 95: 54, 100: 56, 106: 58, 111: 60, 143: 70, 180: 80, 223: 90, 273: 100, 331: 110, 397: 120, 472: 130}
                R600a_kpa = {-97: -70, -95: -66, -93: -62, -91: -58, -88: -54, -85: -50, -80: -46, -75: -42, -70: -38, -63: -34, -55: -30, -50: -28, -45: -26, -40: -24, -35: -22, -29: -20, -23: -18, -16: -16, -9: -14, -1: -12, 7: -10, 16: -8, 25: -6, 35: -4, 45: -2, 56: 0, 67: 2, 79: 4, 92: 6, 105: 8, 119: 10, 134: 12, 150: 14, 166: 16, 183: 18, 201: 20, 220: 22, 239: 24, 260: 26, 281: 28, 303: 30, 327: 32, 351: 34, 376: 36, 403: 38, 430: 40, 458: 42, 488: 44, 519: 46, 551: 48, 584: 50, 618: 52, 653: 54, 690: 56, 728: 58, 768: 60, 986: 70, 1243: 80, 1541: 90, 1885: 100, 2281: 110, 2735: 120, 3257: 130}
                R290_psi = {-11.1: -70, -10.2: -66, -9.1: -62, -7.8: -58, -6.3: -54, -4.5: -50, -2.4: -46, 0.1: -42, 2.9: -38, 6: -34, 9.6: -30, 12: -28, 14: -26, 16: -24, 18: -22, 21: -20, 23: -18, 26: -16, 29: -14, 32: -12, 35: -10, 39: -8, 42: -6, 46: -4, 50: -2, 54: 0, 58: 2, 63: 4, 68: 6, 73: 8, 78: 10, 83: 12, 89: 14, 94: 16, 100: 18, 107: 20, 113: 22, 120: 24, 127: 26, 134: 28, 142: 30, 150: 32, 158: 34, 166: 36, 175: 38, 184: 40, 193: 42, 203: 44, 213: 46, 223: 48, 234: 50, 245: 52, 256: 54, 268: 56, 280: 58, 292: 60, 361: 70, 440: 80, 531: 90}
                R290_kpa = {-77.0: -70, -71.0: -66, -63.0: -62, -54.0: -58, -43.0: -54, -31.0: -50, -16.0: -46, 0.5: -42, 20.0: -38, 42.0: -34, 67.0: -30, 80.0: -28, 95.0: -26, 110.0: -24, 126.0: -22, 143.0: -20, 161.0: -18, 180.0: -16, 200.0: -14, 222.0: -12, 244.0: -10, 267.0: -8, 292.0: -6, 318.0: -4, 345.0: -2, 373.0: 0, 403.0: 2, 434.0: 4, 466.0: 6, 500.0: 8, 535.0: 10, 572.0: 12, 610.0: 14, 650.0: 16, 692.0: 18, 735.0: 20, 780.0: 22, 827.0: 24, 875.0: 26, 926.0: 28, 978.0: 30, 1032.0: 32, 1088.0: 34, 1146.0: 36, 1206.0: 38, 1268.0: 40, 1332.0: 42, 1399.0: 44, 1468.0: 46, 1539.0: 48, 1612.0: 50, 1688.0: 52, 1766.0: 54, 1847.0: 56, 1930.0: 58, 2015.0: 60, 2485.0: 70, 3031.0: 80, 3663.0: 90}
                R134a_kpa = {-64.0: -46, -55.0: -42, -44.0: -38, -32.0: -34, -17.0: -30, -9.0: -28, 0.3: -26, 10.0: -24, 20.0: -22, 31.0: -20, 43.0: -18, 56.0: -16, 69.0: -14, 84.0: -12, 99.0: -10, 116.0: -8, 133.0: -6, 151.0: -4, 171.0: -2, 191.0: 0, 213.0: 2, 236.0: 4, 261.0: 6, 286.0: 8, 313.0: 10, 342.0: 12, 372.0: 14, 403.0: 16, 436.0: 18, 470.0: 20, 507.0: 22, 544.0: 24, 584.0: 26, 626.0: 28, 669.0: 30, 714.0: 32, 761.0: 34, 811.0: 36, 862.0: 38, 915.0: 40, 971.0: 42, 1029.0: 44, 1089.0: 46, 1152.0: 48, 1217.0: 50, 1284.0: 52, 1354.0: 54, 1427.0: 56, 1502.0: 58, 1580.0: 60, 2015.0: 70, 2532.0: 80, 3143.0: 90}
                R134a_psi = {-19.0: -46, -16.3: -42, -13.1: -38, -9.4: -34, -5.0: -30, -2.5: -28, 0.1: -26, 1.4: -24, 2.9: -22, 4.6: -20, 6.3: -18, 8.1: -16, 10.0: -14, 12.0: -12, 14.0: -10, 17.0: -8, 19.0: -6, 22.0: -4, 25.0: -2, 28.0: 0, 31.0: 2, 34.0: 4, 38.0: 6, 41.0: 8, 45.0: 10, 50.0: 12, 54.0: 14, 58.0: 16, 63.0: 18, 68.0: 20, 73.0: 22, 79.0: 24, 85.0: 26, 91.0: 28, 97.0: 30, 103.0: 32, 110.0: 34, 118.0: 36, 125.0: 38, 133.0: 40, 141.0: 42, 149.0: 44, 158.0: 46, 167.0: 48, 176.0: 50, 186.0: 52, 196.0: 54, 207.0: 56, 218.0: 58, 229.0: 60, 292.0: 70, 367.0: 80, 456.0: 90}
                s_num = {'LPG 30g': 72, 'LPG 30g PCM': 55, 'LPG 30g PCM + NANO': 54, 'LPG 40g': 72, 'LPG 40g PCM': 57, 'LPG 40g PCM + NANO': 55, 'LPG 50g': 71, 'LPG 50g PCM': 64, 'LPG 50g PCM + NANO': 64, 'LPG 60g': 66, 'LPG 60g PCM': 61, 'LPG 60g PCM + NANO': 57, 'LPG 70g': 84, 'LPG 70g PCM': 62, 'LPG 70g PCM + NANO': 64, 'R600a 30g': 75, 'R600a 30g PCM': 51, 'R600a 30g PCM + NANO': 69, 'R600a 40g': 77, 'R600a 40g PCM': 56, 'R600a 40g PCM + NANO': 53, 'R600a 50g': 78, 'R600a 50g PCM': 71, 'R600a 50g PCM + NANO': 53, 'R600a 60g': 76, 'R600a 60g PCM': 59, 'R600a 60g PCM + NANO': 57, 'R600a 70g': 68, 'R600a 70g PCM': 61, 'R600a 70g PCM + NANO': 62}

                if e[0:3] == "R60":
                    func = "1"
                else:
                    func = "2"
                tabs = {"290": [R290_psi, R290_kpa], "600": [R600a_psi, R600a_kpa], "134": [R134a_psi, R134a_kpa]}
                a = tabs["600"][0]
                
                if func == "1":
                        def anss(b,c,e):
                                import pandas as pd
                                ans = []
                                ans2 = []
                                for i in range(len(b)):
                                        ans.append(val(a, b[i]))
                                        ans2.append(val(a, c[i]))
                                df1 = pd.DataFrame({'Td':ans,'Ts':ans2})
                                return df1,e
                        return anss(d,s,e)

                b = tabs["290"][0]
                f = .6
                g = .4
                if func == "2":
                        return manss(a,b,d,s,g,f,e)

        return inter(x,y,e)

asd = int(input("How many tables do you want to run.\nTotal is 30\n"))
cont(asd)

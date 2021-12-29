def cont(x):
    import pandas as pd
    a1,a2,a3,a4,a5,a6,a7,a8,a9,a10,a11,a12,a13,a14,a15,a16,a17,a18,a19,a20,a21,a22,a23,a24,a25,a26,a27,a28,a29,a30 = [],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[]
    e1,e2,e3,e4,e5,e6,e7,e8,e9,e10,e11,e12,e13,e14,e15,e16,e17,e18,e19,e20,e21,e22,e23,e24,e25,e26,e27,e28,e29,e30 = [],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[]
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
    with pd.ExcelWriter("C:/Users/Joshua/Documents/Cu/Project/Data/Readings - writtenCOP.xlsx") as writer:
        for i in range(len(names)):
            names[i].to_excel(writer, sheet_name=es[i])
    print ("\nVery Successful!")

def ref(i):
        def exc(i):
                import pandas as pd
                acc = ["R600a 30g", "R600a 30g PCM", "R600a 30g PCM + NANO", "R600a 40g", "R600a 40g PCM", "R600a 40g PCM + NANO", "R600a 50g", "R600a 50g PCM", "R600a 50g PCM + NANO", "R600a 60g", "R600a 60g PCM", "R600a 60g PCM + NANO", "R600a 70g", "R600a 70g PCM", "R600a 70g PCM + NANO", "LPG 30g", "LPG 30g PCM", "LPG 30g PCM + NANO", "LPG 40g", "LPG 40g PCM", "LPG 40g PCM + NANO", "LPG 50g", "LPG 50g PCM", "LPG 50g PCM + NANO", "LPG 60g", "LPG 60g PCM", "LPG 60g PCM + NANO", "LPG 70g", "LPG 70g PCM", "LPG 70g PCM + NANO"]
                e = acc[i]
                print ("\n" + e)
                df = pd.read_excel(r'C:\Users\Joshua\Documents\Cu\Project\Data\Readings - COP.xlsx', sheet_name= e)
                x = list(df['COP'])
                df1 = pd.DataFrame({'COP':x})
                return df1,e
        return exc(i)

asd = int(input("How many tables do you want to run.\nTotal is 30\n"))
cont(asd)

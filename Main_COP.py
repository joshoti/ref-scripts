def cont(x):
    import pandas as pd
    names = [list() for i in range(30)]
    es = [list() for i in range(30)]
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

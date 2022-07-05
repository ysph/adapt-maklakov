import data as ps

def getPPPoints(res):
    if res >= 104:
        return 1
    elif res <= 103 and res >= 87:
        return 2
    elif res <= 86 and res >= 72:
        return 3
    elif res <= 71 and res >= 57:
        return 4
    elif res <= 56 and res >= 43:
        return 5
    elif res <= 42 and res >= 36:
        return 6
    elif res <= 35 and res >= 29:
        return 7
    elif res <= 28 and res >= 23:
        return 8
    elif res <= 22 and res >= 19:
        return 9
    elif res <= 18:
        return 10
    return -1
def getNPUPoints(res):
    if res >= 68:
        return 1
    elif res <= 67 and res >= 52:
        return 2
    elif res <= 51 and res >= 41:
        return 3
    elif res <= 40 and res >= 30:
        return 4
    elif res <= 29 and res >= 20:
        return 5
    elif res <= 19 and res >= 15:
        return 6
    elif res <= 14 and res >= 10:
        return 7
    elif res <= 9 and res >= 8:
        return 8
    elif res <= 7 and res >= 6:
        return 9
    elif res <= 5:
        return 10
    return -1
def getKPPoints(res):
    if res >= 25:
        return 1
    elif res <= 24 and res >= 22:
        return 2
    elif res <= 21 and res >= 19:
        return 3
    elif res <= 18 and res >= 16:
        return 4
    elif res <= 15 and res >= 14:
        return 5
    elif res <= 13 and res >= 12:
        return 6
    elif res <= 11 and res >= 10:
        return 7
    elif res <= 9 and res >= 8:
        return 8
    elif res <= 7 and res >= 6:
        return 9
    elif res <= 5:
        return 10
    return -1
def getMNPoints(res):
    if res >= 18:
        return 1
    elif res <= 17 and res >= 16:
        return 2
    elif res <= 15 and res >= 14:
        return 3
    elif res <= 13 and res >= 12:
        return 4
    elif res <= 11 and res >= 10:
        return 5
    elif res == 9:
        return 6
    elif res == 8 or res == 7:
        return 7
    elif res == 6:
        return 8
    elif res == 5:
        return 9
    elif res <= 4:
        return 10
    return -1
def npuGroup(pts):
    if pts <= 2:
        return "4 группа"
    elif pts <= 5 and pts >= 3:
        return "3 группа"
    elif pts <= 8 and pts >= 6:
        return "2 группа"
    elif pts <= 10 and pts >= 9:
        return "1 группа"
    return -1
def testPsych(data):
    # красный - плюс, зеленое - минус

    truth = 0
    for key, value in ps.truth.items():
        if data[key] == '-':
            truth += 1
    res1 = "1. Достоверность (Д): " + str(truth) + ' из 10. '
    if truth < 10:
        res1 += "Результат - достоверно.\n"
    else:
        res1 += "Результат - не достоверно.\n"
   
    minuses = 0
    pluses = 0
    for key, value in ps.personal_potential.items():
        if data[key] == value == '-':
            minuses += 1
        elif data[key] == value == '+':
            pluses += 1
    points = minuses + pluses
    r2 = points
    
    minuses = 0
    pluses = 0
    for key, value in ps.psych_stability.items():
        if data[key] == value == '-':
            minuses += 1
        elif data[key] == value == '+':
            pluses += 1
    points = minuses + pluses
    r1 = points
    
    minuses = 0
    pluses = 0
    for key, value in ps.communication_ability.items():
        if data[key] == value == '-':
            minuses += 1
            #print(key,data[key])
        if data[key] == value == '+':
            pluses += 1
            #print(key,data[key])
    points = minuses + pluses
    r3 = points
    
    minuses = 0
    pluses = 0
    for key, value in ps.moral_norm.items():
        if data[key] == value == '-':
            minuses += 1
        if data[key] == value == '+':
            pluses += 1
    points = minuses + pluses
    r4 = points

    return truth, getNPUPoints(r1), getPPPoints(r2), getKPPoints(r3), getMNPoints(r4), npuGroup(getNPUPoints(r1))
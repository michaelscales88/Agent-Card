class Agent():

    def __init__(self, name):
        self.agent = name
        self.totalTime = 0
        self.dndTime = 0
        self.excessHours = 0

    def __str__(self):
        print("self.totalTime: " , self.totalTime)
        print("self.dndTime: " , self.dndTime)
        
    def getTotalTime(self):
        return self.totalTime - self.excessHours

    def getDndTime(self):
        return self.dndTime

    def setTotalTime(self, newTime):
        self.totalTime = newTime
        if(self.totalTime > 623988): 
            self.excessHours = self.totalTime - 623988

    def setDndTime(self, newDndTime):
        self.dndTime = newDndTime

    def getPercentDnd(self):
        totalTime = self.totalTime - self.excessHours
        DND = self.dndTime
        value = 0
        if ((totalTime - DND) >= 0):
            value = ((totalTime - DND) / totalTime) * 100
        return value

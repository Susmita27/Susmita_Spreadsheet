class Matrix(object):
    """A simple Matrix class in Python"""

    def __init__(self, rows, columns):
        self.width = columns
        self.height = rows
        self.data = []
        # self.data = [[0] * columns] * rows
        for i in range(self.height):
            self.data.append([])
            for j in range(self.width):
                self.data[i].append(0)

    def setElementAt(self, x, y, value):
        self.data[x][y] = value

    def getElementAt(self, x, y):
        return self.data[x][y]

    def __str__(self):
        result = []
        for row in self.data:
            for cell in row:
                result.append(str(cell))
                result.append(" ")
            result.append("\n")
        string = ''.join(result)
        return string
    
    



